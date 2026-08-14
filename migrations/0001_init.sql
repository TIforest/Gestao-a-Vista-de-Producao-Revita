-- Apontamentos de produção (um registro por fardo lançado)
CREATE TABLE IF NOT EXISTS apontamentos (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  row_hash TEXT NOT NULL UNIQUE, -- hash(lote, numero_fardo, maquina, data_hora) para deduplicar no resync
  lote TEXT NOT NULL,
  cliente TEXT NOT NULL DEFAULT '',
  numero_fardo INTEGER,
  turma TEXT NOT NULL,
  peso_seco REAL NOT NULL DEFAULT 0, -- "Soma de Peso Seco 51%"
  data_hora TEXT NOT NULL,           -- ISO 8601, ex: 2026-08-14T13:38:00
  maquina TEXT NOT NULL,             -- Desaguadora 01..04
  produto TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_apontamentos_data_hora ON apontamentos (data_hora);
CREATE INDEX IF NOT EXISTS idx_apontamentos_turma ON apontamentos (turma);
CREATE INDEX IF NOT EXISTS idx_apontamentos_maquina ON apontamentos (maquina);

-- Metas cadastradas pelo gestor. escopo = 'mes' (referencia = 'YYYY-MM') ou 'dia' (referencia = 'YYYY-MM-DD').
CREATE TABLE IF NOT EXISTS metas (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  escopo TEXT NOT NULL CHECK (escopo IN ('mes', 'dia')),
  referencia TEXT NOT NULL,
  valor REAL NOT NULL,
  turnos_por_dia INTEGER NOT NULL DEFAULT 4,
  horas_por_turno REAL NOT NULL DEFAULT 6,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_by TEXT NOT NULL DEFAULT '',
  UNIQUE (escopo, referencia)
);

-- Estado da última sincronização com o SharePoint (linha única, id fixo = 1).
CREATE TABLE IF NOT EXISTS sync_state (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  sharepoint_last_modified TEXT,
  last_sync_at TEXT,
  last_sync_status TEXT NOT NULL DEFAULT 'nunca_sincronizado',
  last_sync_rows INTEGER NOT NULL DEFAULT 0,
  last_error TEXT
);

INSERT OR IGNORE INTO sync_state (id, last_sync_status) VALUES (1, 'nunca_sincronizado');
