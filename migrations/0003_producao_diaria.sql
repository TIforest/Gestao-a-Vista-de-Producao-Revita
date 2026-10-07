-- Total de produção por dia (YYYY-MM-DD), recalculado da planilha a cada
-- leitura pros dias que a leitura cobre. A produção acumulada do mês é a
-- soma dos dias do mês. Substitui o contador incremental producao_mensal
-- (que só somava e nunca refletia correções/exclusões feitas na planilha —
-- causa da divergência com o BI em 2026-10-07). A tabela producao_mensal
-- fica no banco por histórico, mas não é mais lida nem escrita.
CREATE TABLE IF NOT EXISTS producao_diaria (
  dia TEXT PRIMARY KEY,
  total_peso REAL NOT NULL DEFAULT 0,
  linhas INTEGER NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
