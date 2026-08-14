-- Acumulado mensal (contador incremental). O painel não guarda o histórico
-- completo de linhas indefinidamente (só uma janela recente, para turno/dia) —
-- a produção do mês é somada aqui conforme cada linha nova aparece na planilha,
-- e não recalculada a partir de milhares de linhas antigas a cada sync.
CREATE TABLE IF NOT EXISTS producao_mensal (
  ano_mes TEXT PRIMARY KEY, -- 'YYYY-MM'
  total_peso REAL NOT NULL DEFAULT 0,
  linhas_contadas INTEGER NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
