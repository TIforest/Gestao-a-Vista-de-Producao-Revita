import type { Env } from "../types";
import { encodeSharingUrl } from "./shareLink";
import { parseWorkbook, parseGraphRangeValues } from "./parseExcel";
import { daysAgoLocalISOStart } from "./date";

// O painel não guarda o histórico completo da planilha — só uma janela
// recente (cobre o dia atual + margem de segurança pra virada de turno/dia).
// A produção acumulada do mês fica num contador à parte (producao_mensal),
// incrementado conforme cada linha nova aparece, sem precisar reprocessar
// milhares de linhas antigas a cada sincronização.
export const RETENTION_DAYS = 3;

// Quantas linhas (de trás pra frente) pedir na sincronização automática.
// A planilha de origem já passa de 11 mil linhas (todo o histórico desde
// junho) e cresce ~150-200 linhas/dia — baixar e processar o arquivo
// inteiro (como fazíamos antes, via SheetJS) estourava o limite de CPU do
// Worker sempre que o arquivo mudava. Com a API de Range do Graph, pedimos
// só as últimas N linhas diretamente à Microsoft (ela computa o recorte do
// lado dela) — folga generosa sobre RETENTION_DAYS pra cobrir dias mais
// cheios e retomadas depois de uma parada.
const LINHAS_RECENTES = 1500;

interface TokenResponse {
  access_token: string;
  expires_in: number;
}

// O SharePoint volta e meia responde 503 "Something went wrong, tente
// novamente" por instabilidade passageira do próprio serviço — sem isso, uma
// única falha desse tipo já marcava a sincronização como "erro" na tela.
async function fetchWithRetry(url: string, init: RequestInit, attempts = 4): Promise<Response> {
  let lastRes: Response | null = null;
  let lastErr: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      const res = await fetch(url, init);
      if (res.ok || res.status < 500) return res;
      lastRes = res;
    } catch (err) {
      lastErr = err;
    }
    if (i < attempts - 1) await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
  }
  if (lastRes) return lastRes;
  throw lastErr;
}

// O D1 volta e meia devolve "storage operation exceeded timeout" — igual o
// 503 do SharePoint, é instabilidade passageira do serviço, não bug nosso.
// Com uma sincronização fazendo dezenas de leituras/gravações em sequência,
// uma falha dessas em qualquer uma já derrubava a sincronização inteira.
async function comRetry<T>(fn: () => Promise<T>, attempts = 3): Promise<T> {
  let lastErr: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (i < attempts - 1) await new Promise((resolve) => setTimeout(resolve, 300 * (i + 1)));
    }
  }
  throw lastErr;
}

async function getGraphToken(env: Env): Promise<string> {
  const res = await fetch(`https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: env.MS_CLIENT_ID,
      client_secret: env.MS_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  if (!res.ok) {
    throw new Error(`Falha ao autenticar no Microsoft Graph (${res.status}): ${await res.text()}`);
  }
  const data = (await res.json()) as TokenResponse;
  return data.access_token;
}

interface DriveItemMeta {
  id: string;
  lastModifiedDateTime: string;
  parentReference: { driveId: string };
}

function colIndexToLetter(n: number): string {
  // n é base-1 (1 = A)
  let s = "";
  let num = n;
  while (num > 0) {
    const rem = (num - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

/** Nome da aba a usar: a configurada em MS_SHEET_NAME se existir na planilha, senão a primeira. */
async function resolverNomeAba(
  driveId: string,
  itemId: string,
  token: string,
  sheetNameConfigurada?: string
): Promise<string> {
  const res = await fetchWithRetry(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=name`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Falha ao listar abas da planilha (${res.status}): ${await res.text()}`);
  const data = (await res.json()) as { value: { name: string }[] };
  const nomes = data.value.map((s) => s.name);
  if (sheetNameConfigurada && nomes.includes(sheetNameConfigurada)) return sheetNameConfigurada;
  const primeira = nomes[0];
  if (!primeira) throw new Error("Planilha sem nenhuma aba.");
  return primeira;
}

/**
 * Busca só as últimas `LINHAS_RECENTES` linhas da planilha via API de Range
 * do Graph — sem baixar o arquivo .xlsx inteiro. Faz 3 chamadas leves:
 * dimensão da área usada (só contagem, sem valores), cabeçalho (1 linha) e
 * o recorte final (N linhas). Cada uma custa bytes/CPU proporcionais só ao
 * que pede, não ao tamanho total da planilha.
 */
async function buscarLinhasRecentesViaRange(
  env: Env,
  token: string,
  driveId: string,
  itemId: string
): Promise<{ headerRow: unknown[]; dataRows: unknown[][] }> {
  const sheetName = await resolverNomeAba(driveId, itemId, token, env.MS_SHEET_NAME);
  const sheetPath = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')`;

  const dimRes = await fetchWithRetry(
    `${sheetPath}/usedRange(valuesOnly=true)?$select=rowCount,columnCount`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!dimRes.ok) throw new Error(`Falha ao consultar tamanho da planilha (${dimRes.status}): ${await dimRes.text()}`);
  const dim = (await dimRes.json()) as { rowCount: number; columnCount: number };
  const lastCol = colIndexToLetter(dim.columnCount);

  const headerRes = await fetchWithRetry(
    `${sheetPath}/range(address='A1:${lastCol}1')?$select=values`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!headerRes.ok) throw new Error(`Falha ao ler cabeçalho da planilha (${headerRes.status}): ${await headerRes.text()}`);
  const headerData = (await headerRes.json()) as { values: unknown[][] };
  const headerRow = headerData.values[0] ?? [];

  const startRow = Math.max(2, dim.rowCount - LINHAS_RECENTES + 1);
  const dataRes = await fetchWithRetry(
    `${sheetPath}/range(address='A${startRow}:${lastCol}${dim.rowCount}')?$select=values`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!dataRes.ok) throw new Error(`Falha ao ler linhas recentes da planilha (${dataRes.status}): ${await dataRes.text()}`);
  const rangeData = (await dataRes.json()) as { values: unknown[][] };

  return { headerRow, dataRows: rangeData.values };
}

export interface SyncResult {
  status: "sem_alteracao" | "sincronizado" | "erro";
  rows: number;
  warnings: string[];
  error?: string;
  sharepointLastModified?: string;
}

/**
 * Busca o arquivo do SharePoint via Graph, compara `lastModifiedDateTime` com o
 * que está salvo em sync_state e, se mudou (ou `force`), baixa, faz parse
 * (só a janela recente) e faz upsert em `apontamentos`, incrementando o
 * acumulado mensal para as linhas que forem genuinamente novas.
 */
export async function runSync(env: Env, opts: { force?: boolean } = {}): Promise<SyncResult> {
  try {
    const token = await getGraphToken(env);
    const shareId = encodeSharingUrl(env.MS_SHARE_URL);

    const metaRes = await fetchWithRetry(
      `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem?$select=id,lastModifiedDateTime,parentReference`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) {
      throw new Error(`Falha ao consultar arquivo no SharePoint (${metaRes.status}): ${await metaRes.text()}`);
    }
    const meta = (await metaRes.json()) as DriveItemMeta;

    const current = await comRetry(() =>
      env.DB.prepare("SELECT sharepoint_last_modified FROM sync_state WHERE id = 1").first<{
        sharepoint_last_modified: string | null;
      }>()
    );

    if (!opts.force && current?.sharepoint_last_modified === meta.lastModifiedDateTime) {
      await comRetry(() =>
        env.DB.prepare(
          "UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'sem_alteracao', last_error = NULL WHERE id = 1"
        ).run()
      );
      return { status: "sem_alteracao", rows: 0, warnings: [], sharepointLastModified: meta.lastModifiedDateTime };
    }

    const { headerRow, dataRows } = await buscarLinhasRecentesViaRange(
      env,
      token,
      meta.parentReference.driveId,
      meta.id
    );

    const windowStart = daysAgoLocalISOStart(RETENTION_DAYS);
    const { rows, warnings } = await parseGraphRangeValues(headerRow, dataRows, windowStart);
    await upsertApontamentos(env, rows);
    await pruneOldApontamentos(env, windowStart);

    await comRetry(() =>
      env.DB.prepare(
        `UPDATE sync_state SET sharepoint_last_modified = ?, last_sync_at = datetime('now'),
         last_sync_status = 'sincronizado', last_sync_rows = ?, last_error = NULL WHERE id = 1`
      )
        .bind(meta.lastModifiedDateTime, rows.length)
        .run()
    );

    return { status: "sincronizado", rows: rows.length, warnings, sharepointLastModified: meta.lastModifiedDateTime };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    try {
      await comRetry(() =>
        env.DB.prepare(
          "UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'erro', last_error = ? WHERE id = 1"
        )
          .bind(message)
          .run()
      );
    } catch {
      // Se nem isso conseguir gravar, o D1 está mesmo fora do ar — a tela vai
      // continuar mostrando o último estado bom conhecido até a próxima tentativa.
    }
    return { status: "erro", rows: 0, warnings: [], error: message };
  }
}

async function upsertApontamentos(env: Env, rows: Awaited<ReturnType<typeof parseWorkbook>>["rows"]): Promise<void> {
  const CHUNK = 50;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const chunk = rows.slice(i, i + CHUNK);

    // Descobre quais hashes já existem, pra só somar no acumulado mensal
    // as linhas genuinamente novas (evita contar de novo em cada sync).
    // Sem retry aqui de propósito: são ~2-3 chamadas ao D1 por bloco de 50
    // linhas, dezenas de blocos por sincronização — tentar de novo cada uma
    // (visto na prática) pode multiplicar o tempo total a ponto do Worker
    // estourar o limite de recursos quando o D1 está mesmo instável, o que é
    // pior do que simplesmente falhar rápido e deixar a próxima tentativa
    // (cron de 1 min, ou o auto-sync do navegador) resolver.
    const placeholders = chunk.map(() => "?").join(",");
    const existing = await env.DB.prepare(`SELECT row_hash FROM apontamentos WHERE row_hash IN (${placeholders})`)
      .bind(...chunk.map((r) => r.row_hash))
      .all<{ row_hash: string }>();
    const existingHashes = new Set(existing.results.map((r) => r.row_hash));
    const novas = chunk.filter((r) => !existingHashes.has(r.row_hash));

    if (novas.length > 0) {
      const porMes = new Map<string, number>();
      for (const row of novas) {
        const anoMes = row.data_hora.slice(0, 7);
        porMes.set(anoMes, (porMes.get(anoMes) ?? 0) + row.peso_seco);
      }
      for (const [anoMes, delta] of porMes) {
        await env.DB.prepare(
          `INSERT INTO producao_mensal (ano_mes, total_peso, linhas_contadas, updated_at)
           VALUES (?, ?, ?, datetime('now'))
           ON CONFLICT(ano_mes) DO UPDATE SET
             total_peso = total_peso + excluded.total_peso,
             linhas_contadas = linhas_contadas + excluded.linhas_contadas,
             updated_at = excluded.updated_at`
        )
          .bind(anoMes, delta, novas.filter((r) => r.data_hora.slice(0, 7) === anoMes).length)
          .run();
      }
    }

    // Só grava as linhas genuinamente novas (mesmo grupo "novas" de cima) —
    // gravar a janela inteira a cada sync (como fazia antes) reescreve
    // centenas de linhas sem mudança nenhuma, e com o auto-sync rodando de
    // 1 em 1 minuto isso sozinho quase estourou a cota de escrita do D1.
    const stmts = novas.map((row) =>
      env.DB.prepare(
        `INSERT INTO apontamentos (row_hash, lote, cliente, numero_fardo, turma, peso_seco, data_hora, maquina, produto)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(row_hash) DO NOTHING`
      ).bind(
        row.row_hash,
        row.lote,
        row.cliente,
        row.numero_fardo,
        row.turma,
        row.peso_seco,
        row.data_hora,
        row.maquina,
        row.produto
      )
    );
    if (stmts.length > 0) await env.DB.batch(stmts);
  }
}

async function pruneOldApontamentos(env: Env, windowStartISO: string): Promise<void> {
  await env.DB.prepare("DELETE FROM apontamentos WHERE data_hora < ?").bind(windowStartISO).run();
}

export { upsertApontamentos };
