import type { Env } from "../types";
import { encodeSharingUrl } from "./shareLink";
import { parseWorkbook } from "./parseExcel";
import { daysAgoLocalISOStart } from "./date";

// O painel não guarda o histórico completo da planilha — só uma janela
// recente (cobre o dia atual + margem de segurança pra virada de turno/dia).
// A produção acumulada do mês fica num contador à parte (producao_mensal),
// incrementado conforme cada linha nova aparece, sem precisar reprocessar
// milhares de linhas antigas a cada sincronização.
export const RETENTION_DAYS = 3;

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

    const contentRes = await fetchWithRetry(
      `https://graph.microsoft.com/v1.0/drives/${meta.parentReference.driveId}/items/${meta.id}/content`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!contentRes.ok) {
      throw new Error(`Falha ao baixar planilha (${contentRes.status}): ${await contentRes.text()}`);
    }
    const buffer = await contentRes.arrayBuffer();

    const windowStart = daysAgoLocalISOStart(RETENTION_DAYS);
    const { rows, warnings } = await parseWorkbook(buffer, env.MS_SHEET_NAME, windowStart);
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
    const placeholders = chunk.map(() => "?").join(",");
    const existing = await comRetry(() =>
      env.DB.prepare(`SELECT row_hash FROM apontamentos WHERE row_hash IN (${placeholders})`)
        .bind(...chunk.map((r) => r.row_hash))
        .all<{ row_hash: string }>()
    );
    const existingHashes = new Set(existing.results.map((r) => r.row_hash));
    const novas = chunk.filter((r) => !existingHashes.has(r.row_hash));

    if (novas.length > 0) {
      const porMes = new Map<string, number>();
      for (const row of novas) {
        const anoMes = row.data_hora.slice(0, 7);
        porMes.set(anoMes, (porMes.get(anoMes) ?? 0) + row.peso_seco);
      }
      for (const [anoMes, delta] of porMes) {
        await comRetry(() =>
          env.DB.prepare(
            `INSERT INTO producao_mensal (ano_mes, total_peso, linhas_contadas, updated_at)
             VALUES (?, ?, ?, datetime('now'))
             ON CONFLICT(ano_mes) DO UPDATE SET
               total_peso = total_peso + excluded.total_peso,
               linhas_contadas = linhas_contadas + excluded.linhas_contadas,
               updated_at = excluded.updated_at`
          )
            .bind(anoMes, delta, novas.filter((r) => r.data_hora.slice(0, 7) === anoMes).length)
            .run()
        );
      }
    }

    const stmts = chunk.map((row) =>
      env.DB.prepare(
        `INSERT INTO apontamentos (row_hash, lote, cliente, numero_fardo, turma, peso_seco, data_hora, maquina, produto)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(row_hash) DO UPDATE SET
           cliente = excluded.cliente,
           peso_seco = excluded.peso_seco,
           produto = excluded.produto`
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
    if (stmts.length > 0) await comRetry(() => env.DB.batch(stmts));
  }
}

async function pruneOldApontamentos(env: Env, windowStartISO: string): Promise<void> {
  await comRetry(() => env.DB.prepare("DELETE FROM apontamentos WHERE data_hora < ?").bind(windowStartISO).run());
}

export { upsertApontamentos };
