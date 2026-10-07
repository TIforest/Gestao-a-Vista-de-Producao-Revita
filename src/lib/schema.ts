import type { Env } from "../types";

// A migration oficial é migrations/0003_producao_diaria.sql (aplicada pelo
// workflow "DB Migrate"). Como o deploy do Worker é automático e a migration
// é manual, o código também garante a tabela em runtime — assim não existe
// uma janela entre o deploy e a migration em que o painel quebra. É
// idempotente e roda uma vez por isolate.
let schemaGarantido = false;

export async function garantirSchema(env: Env): Promise<void> {
  if (schemaGarantido) return;
  await env.DB.prepare(
    `CREATE TABLE IF NOT EXISTS producao_diaria (
       dia TEXT PRIMARY KEY,
       total_peso REAL NOT NULL DEFAULT 0,
       linhas INTEGER NOT NULL DEFAULT 0,
       updated_at TEXT NOT NULL DEFAULT (datetime('now'))
     )`
  ).run();
  schemaGarantido = true;
}
