import type { Env } from "../types";
import { encodeSharingUrl } from "./shareLink";
import { parseGraphRangeValues, type ParseResult, type TotalDia } from "./parseExcel";
import { addDaysISODate, daysAgoLocalISOStart, todayBrazilISODate } from "./date";
import { garantirSchema } from "./schema";

// O painel não guarda o histórico completo da planilha — só uma janela
// recente (cobre o dia atual + margem de segurança pra virada de turno/dia).
// A produção do mês sai da tabela producao_diaria (um total por dia,
// recalculado da planilha a cada leitura — ver atualizarProducaoDiaria).
export const RETENTION_DAYS = 3;

// Quantas linhas (de trás pra frente) pedir na sincronização automática.
// A planilha de origem já passa de 11 mil linhas (todo o histórico desde
// junho) e cresce ~100-200 linhas/dia — baixar e processar o arquivo
// inteiro (como fazíamos antes, via SheetJS) estourava o limite de CPU do
// Worker sempre que o arquivo mudava. Com a API de Range do Graph, pedimos
// só as últimas N linhas diretamente à Microsoft (ela computa o recorte do
// lado dela). 1500 linhas ≈ 2 semanas: além de cobrir a janela, é até
// onde correções feitas na planilha ainda entram no acumulado do mês.
const LINHAS_RECENTES = 1500;
// Teto pra recontagem manual (/api/recontar) — só roda a pedido.
const LINHAS_RECONTAGEM_MAX = 12000;
// Primeira leitura depois que producao_diaria nasce (ou foi zerada): lê bem
// mais linhas pra montar o histórico do mês de uma vez, sem depender de
// alguém chamar /api/recontar. ~2 meses de planilha.
const LINHAS_BOOTSTRAP = 6000;
let bootstrapFeito = false;

async function precisaBootstrap(env: Env): Promise<boolean> {
  if (bootstrapFeito) return false;
  const row = await env.DB.prepare("SELECT COUNT(*) AS n FROM producao_diaria").first<{ n: number }>();
  if ((row?.n ?? 0) > 0) {
    bootstrapFeito = true;
    return false;
  }
  return true;
}

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
 * Busca só as últimas `linhas` linhas da planilha via API de Range do Graph
 * — sem baixar o arquivo .xlsx inteiro. Faz 3 chamadas leves: dimensão da
 * área usada (só contagem, sem valores), cabeçalho (1 linha) e o recorte
 * final (N linhas). Cada uma custa bytes/CPU proporcionais só ao que pede,
 * não ao tamanho total da planilha.
 */
async function buscarLinhasRecentesViaRange(
  env: Env,
  token: string,
  driveId: string,
  itemId: string,
  linhas: number
): Promise<{ headerRow: unknown[]; dataRows: unknown[][]; desdeInicio: boolean }> {
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

  const startRow = Math.max(2, dim.rowCount - linhas + 1);
  const dataRes = await fetchWithRetry(
    `${sheetPath}/range(address='A${startRow}:${lastCol}${dim.rowCount}')?$select=values`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!dataRes.ok) throw new Error(`Falha ao ler linhas recentes da planilha (${dataRes.status}): ${await dataRes.text()}`);
  const rangeData = (await dataRes.json()) as { values: unknown[][] };

  // desdeInicio: o recorte começou na primeira linha de dados, ou seja, a
  // leitura cobre a planilha inteira (vale pra recontagem).
  return { headerRow, dataRows: rangeData.values, desdeInicio: startRow === 2 };
}

export interface AplicacaoResult {
  novas: number;
  removidas: number;
  diasAtualizados: number;
  coberturaDesde: string | null;
}

export interface SyncResult {
  status: "sem_alteracao" | "sincronizado" | "erro";
  rows: number;
  warnings: string[];
  error?: string;
  sharepointLastModified?: string;
  aplicacao?: AplicacaoResult;
}

/**
 * Busca o arquivo do SharePoint via Graph, compara `lastModifiedDateTime` com o
 * que está salvo em sync_state e, se mudou (ou `force`), lê as últimas N
 * linhas, faz parse e aplica no banco (ver aplicarLeitura).
 *
 * `linhas`: quantas linhas ler de trás pra frente (padrão LINHAS_RECENTES);
 * a recontagem manual passa um número maior pra cobrir o mês inteiro.
 */
export async function runSync(env: Env, opts: { force?: boolean; linhas?: number } = {}): Promise<SyncResult> {
  try {
    await garantirSchema(env);
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

    const bootstrap = opts.linhas === undefined && (await precisaBootstrap(env));
    if (!opts.force && !bootstrap && current?.sharepoint_last_modified === meta.lastModifiedDateTime) {
      await comRetry(() =>
        env.DB.prepare(
          "UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'sem_alteracao', last_error = NULL WHERE id = 1"
        ).run()
      );
      return { status: "sem_alteracao", rows: 0, warnings: [], sharepointLastModified: meta.lastModifiedDateTime };
    }

    const linhas = Math.min(
      LINHAS_RECONTAGEM_MAX,
      Math.max(1, Math.floor(opts.linhas ?? (bootstrap ? LINHAS_BOOTSTRAP : LINHAS_RECENTES)))
    );
    const { headerRow, dataRows, desdeInicio } = await buscarLinhasRecentesViaRange(
      env,
      token,
      meta.parentReference.driveId,
      meta.id,
      linhas
    );

    const windowStart = daysAgoLocalISOStart(RETENTION_DAYS);
    const parsed = await parseGraphRangeValues(headerRow, dataRows, windowStart);
    const aplicacao = await aplicarLeitura(env, parsed, windowStart, desdeInicio);
    if (aplicacao.diasAtualizados > 0) bootstrapFeito = true;

    await comRetry(() =>
      env.DB.prepare(
        `UPDATE sync_state SET sharepoint_last_modified = ?, last_sync_at = datetime('now'),
         last_sync_status = 'sincronizado', last_sync_rows = ?, last_error = NULL WHERE id = 1`
      )
        .bind(meta.lastModifiedDateTime, parsed.rows.length)
        .run()
    );

    return {
      status: "sincronizado",
      rows: parsed.rows.length,
      warnings: parsed.warnings,
      sharepointLastModified: meta.lastModifiedDateTime,
      aplicacao,
    };
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

/**
 * Aplica no banco o que foi lido da planilha, com semântica de CONJUNTO:
 * nos dias que a leitura cobre por inteiro, o banco fica igual à planilha —
 * linhas novas entram, linhas que sumiram ou foram corrigidas na planilha
 * (lote, hora, peso, cliente...) saem. Antes o sync só acrescentava, e
 * cada correção feita na planilha virava um "fantasma" que inflava a turma
 * e o mês (foi a causa da divergência com o BI vista em 2026-10-07).
 *
 * Também recalcula producao_diaria (um total por dia) pros dias cobertos —
 * é daí que sai a produção acumulada do mês.
 *
 * `desdeInicio`: a leitura começou na primeira linha de dados (planilha
 * inteira ou upload de arquivo completo). Senão, o dia mais antigo lido
 * pode estar cortado pela metade e não entra na cobertura.
 */
export async function aplicarLeitura(
  env: Env,
  parsed: ParseResult,
  windowStartISO: string,
  desdeInicio: boolean
): Promise<AplicacaoResult> {
  await garantirSchema(env);
  const coberturaDesde = parsed.diaMaisAntigo === null
    ? null
    : desdeInicio
      ? parsed.diaMaisAntigo
      : addDaysISODate(parsed.diaMaisAntigo, 1);

  const { novas, removidas } = await reconciliarApontamentos(env, parsed, windowStartISO, coberturaDesde);
  const diasAtualizados = await atualizarProducaoDiaria(env, parsed.totaisPorDia, coberturaDesde);
  await pruneOldApontamentos(env, windowStartISO);
  return { novas, removidas, diasAtualizados, coberturaDesde };
}

async function reconciliarApontamentos(
  env: Env,
  parsed: ParseResult,
  windowStartISO: string,
  coberturaDesde: string | null
): Promise<{ novas: number; removidas: number }> {
  const existentes = await comRetry(() =>
    env.DB.prepare("SELECT row_hash, data_hora FROM apontamentos WHERE data_hora >= ?")
      .bind(windowStartISO)
      .all<{ row_hash: string; data_hora: string }>()
  );
  const hashesPlanilha = new Set(parsed.rows.map((r) => r.row_hash));
  const hashesBanco = new Set(existentes.results.map((r) => r.row_hash));

  const novas = parsed.rows.filter((r) => !hashesBanco.has(r.row_hash));
  // Só remove em dias que a leitura cobre por inteiro — num dia cortado pela
  // metade, uma linha ausente pode só estar antes do recorte.
  const fantasmas = existentes.results
    .filter((r) => !hashesPlanilha.has(r.row_hash) && coberturaDesde !== null && r.data_hora.slice(0, 10) >= coberturaDesde)
    .map((r) => r.row_hash);

  const CHUNK = 50;
  for (let i = 0; i < novas.length; i += CHUNK) {
    const stmts = novas.slice(i, i + CHUNK).map((row) =>
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
    await env.DB.batch(stmts);
  }

  const DEL_CHUNK = 100;
  for (let i = 0; i < fantasmas.length; i += DEL_CHUNK) {
    const chunk = fantasmas.slice(i, i + DEL_CHUNK);
    const placeholders = chunk.map(() => "?").join(",");
    await env.DB.prepare(`DELETE FROM apontamentos WHERE row_hash IN (${placeholders})`).bind(...chunk).run();
  }

  return { novas: novas.length, removidas: fantasmas.length };
}

/**
 * Grava o total de cada dia coberto pela leitura (de `coberturaDesde` até
 * hoje, ou até o último dia lido, se houver linha com data futura por
 * engano). Dia coberto sem nenhuma linha vira 0 — se a planilha não tem
 * nada naquele dia, o painel também não deve ter.
 */
async function atualizarProducaoDiaria(
  env: Env,
  totaisPorDia: Map<string, TotalDia>,
  coberturaDesde: string | null
): Promise<number> {
  if (coberturaDesde === null) return 0;
  let ultimoDia = todayBrazilISODate();
  for (const dia of totaisPorDia.keys()) if (dia > ultimoDia) ultimoDia = dia;

  const stmts: D1PreparedStatement[] = [];
  for (let dia = coberturaDesde; dia <= ultimoDia; dia = addDaysISODate(dia, 1)) {
    const t = totaisPorDia.get(dia) ?? { peso: 0, linhas: 0 };
    stmts.push(
      env.DB.prepare(
        `INSERT INTO producao_diaria (dia, total_peso, linhas, updated_at)
         VALUES (?, ?, ?, datetime('now'))
         ON CONFLICT(dia) DO UPDATE SET
           total_peso = excluded.total_peso,
           linhas = excluded.linhas,
           updated_at = excluded.updated_at`
      ).bind(dia, t.peso, t.linhas)
    );
    if (stmts.length > 400) break; // trava de segurança contra data absurda
  }
  const CHUNK = 100;
  for (let i = 0; i < stmts.length; i += CHUNK) await env.DB.batch(stmts.slice(i, i + CHUNK));
  return stmts.length;
}

async function pruneOldApontamentos(env: Env, windowStartISO: string): Promise<void> {
  await env.DB.prepare("DELETE FROM apontamentos WHERE data_hora < ?").bind(windowStartISO).run();
}
