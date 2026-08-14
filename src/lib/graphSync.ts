import type { Env } from "../types";
import { encodeSharingUrl } from "./shareLink";
import { parseWorkbook } from "./parseExcel";

interface TokenResponse {
  access_token: string;
  expires_in: number;
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
 * que está salvo em sync_state e, se mudou (ou `force`), baixa, faz parse e faz
 * upsert incremental em `apontamentos` (idempotente via row_hash).
 */
export async function runSync(env: Env, opts: { force?: boolean } = {}): Promise<SyncResult> {
  try {
    const token = await getGraphToken(env);
    const shareId = encodeSharingUrl(env.MS_SHARE_URL);

    const metaRes = await fetch(
      `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem?$select=id,lastModifiedDateTime,parentReference`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) {
      throw new Error(`Falha ao consultar arquivo no SharePoint (${metaRes.status}): ${await metaRes.text()}`);
    }
    const meta = (await metaRes.json()) as DriveItemMeta;

    const current = await env.DB.prepare("SELECT sharepoint_last_modified FROM sync_state WHERE id = 1").first<{
      sharepoint_last_modified: string | null;
    }>();

    if (!opts.force && current?.sharepoint_last_modified === meta.lastModifiedDateTime) {
      await env.DB.prepare(
        "UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'sem_alteracao', last_error = NULL WHERE id = 1"
      ).run();
      return { status: "sem_alteracao", rows: 0, warnings: [], sharepointLastModified: meta.lastModifiedDateTime };
    }

    const contentRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${meta.parentReference.driveId}/items/${meta.id}/content`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!contentRes.ok) {
      throw new Error(`Falha ao baixar planilha (${contentRes.status}): ${await contentRes.text()}`);
    }
    const buffer = await contentRes.arrayBuffer();

    const { rows, warnings } = await parseWorkbook(buffer, env.MS_SHEET_NAME);
    await upsertApontamentos(env, rows);

    await env.DB.prepare(
      `UPDATE sync_state SET sharepoint_last_modified = ?, last_sync_at = datetime('now'),
       last_sync_status = 'sincronizado', last_sync_rows = ?, last_error = NULL WHERE id = 1`
    )
      .bind(meta.lastModifiedDateTime, rows.length)
      .run();

    return { status: "sincronizado", rows: rows.length, warnings, sharepointLastModified: meta.lastModifiedDateTime };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await env.DB.prepare(
      "UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'erro', last_error = ? WHERE id = 1"
    )
      .bind(message)
      .run();
    return { status: "erro", rows: 0, warnings: [], error: message };
  }
}

async function upsertApontamentos(env: Env, rows: Awaited<ReturnType<typeof parseWorkbook>>["rows"]): Promise<void> {
  const CHUNK = 50;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const chunk = rows.slice(i, i + CHUNK);
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
    if (stmts.length > 0) await env.DB.batch(stmts);
  }
}

export { upsertApontamentos };
