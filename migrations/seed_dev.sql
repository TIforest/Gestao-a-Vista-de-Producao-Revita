-- Dados fictícios só para testar o painel localmente (não usar em produção).
INSERT OR IGNORE INTO apontamentos (row_hash, lote, cliente, numero_fardo, turma, peso_seco, data_hora, maquina, produto) VALUES
('seed1', 'CMP140826-75', 'Iguaçu', 20, 'E', 720.79, date('now') || 'T13:38:00', '03', 'Revitacel'),
('seed2', 'IMB140826-76', 'Imbraliti', 18, 'E', 489.51, date('now') || 'T13:36:00', '02', 'Revitacel'),
('seed3', 'IMB140826-76', 'Imbraliti', 17, 'E', 505.68, date('now') || 'T13:19:00', '02', 'Revitacel'),
('seed4', 'CMP140826-75', 'Iguaçu', 19, 'B', 665.42, date('now') || 'T13:15:00', '01', 'Revitacel'),
('seed5', 'CMP140826-74', 'Iguaçu', 15, 'C', 610.10, date('now') || 'T12:55:00', '01', 'Revitacel'),
('seed6', 'IMB140826-70', 'Imbraliti', 10, 'A', 300.00, date('now') || 'T09:00:00', '04', 'Revitacel'),
('seed7', 'IMB140826-71', 'Imbraliti', 11, 'D', 450.00, date('now', '-5 days') || 'T09:00:00', '03', 'Revitacel');

INSERT OR IGNORE INTO metas (escopo, referencia, valor, turnos_por_dia, horas_por_turno, updated_by) VALUES
('dia', date('now'), 130000, 4, 6, 'seed'),
('mes', strftime('%Y-%m','now'), 1500000, 4, 6, 'seed');
