import type { Env } from "../types";

export interface MetaDia {
  referencia: string;
  valor: number;
  turnos_por_dia: number;
  horas_por_turno: number;
}

export interface MetasDerivadas {
  metaDia: number;
  metaMes: number;
  metaTurno: number; // metaDia / turnos_por_dia
  metaHora: number; // metaTurno / horas_por_turno
  turnosPorDia: number;
  horasPorTurno: number;
}

// A meta "vale" a partir da data em que foi cadastrada até a próxima alteração
// — não precisa ser recadastrada todo dia/mês. Por isso buscamos a referência
// mais recente que seja <= a data pedida, em vez de exigir bater exatamente.
export async function getMetaDia(env: Env, dateISO: string): Promise<MetaDia> {
  const row = await env.DB.prepare(
    `SELECT referencia, valor, turnos_por_dia, horas_por_turno FROM metas
     WHERE escopo = 'dia' AND referencia <= ? ORDER BY referencia DESC LIMIT 1`
  )
    .bind(dateISO)
    .first<MetaDia>();
  if (row) return row;
  return {
    referencia: dateISO,
    valor: 0,
    turnos_por_dia: Number(env.TURNOS_POR_DIA_PADRAO) || 4,
    horas_por_turno: Number(env.HORAS_POR_TURNO_PADRAO) || 6,
  };
}

export async function getMetaMes(env: Env, yearMonth: string): Promise<number> {
  const row = await env.DB.prepare(
    "SELECT valor FROM metas WHERE escopo = 'mes' AND referencia <= ? ORDER BY referencia DESC LIMIT 1"
  )
    .bind(yearMonth)
    .first<{ valor: number }>();
  return row?.valor ?? 0;
}

export async function getMetasDerivadas(env: Env, dateISO: string, yearMonth: string): Promise<MetasDerivadas> {
  const [dia, mes] = await Promise.all([getMetaDia(env, dateISO), getMetaMes(env, yearMonth)]);
  const metaTurno = dia.turnos_por_dia > 0 ? dia.valor / dia.turnos_por_dia : 0;
  const metaHora = dia.horas_por_turno > 0 ? metaTurno / dia.horas_por_turno : 0;
  return {
    metaDia: dia.valor,
    metaMes: mes,
    metaTurno,
    metaHora,
    turnosPorDia: dia.turnos_por_dia,
    horasPorTurno: dia.horas_por_turno,
  };
}

export async function upsertMetaDia(
  env: Env,
  dateISO: string,
  valor: number,
  turnosPorDia: number,
  horasPorTurno: number,
  updatedBy: string
): Promise<void> {
  await env.DB.prepare(
    `INSERT INTO metas (escopo, referencia, valor, turnos_por_dia, horas_por_turno, updated_at, updated_by)
     VALUES ('dia', ?, ?, ?, ?, datetime('now'), ?)
     ON CONFLICT(escopo, referencia) DO UPDATE SET
       valor = excluded.valor,
       turnos_por_dia = excluded.turnos_por_dia,
       horas_por_turno = excluded.horas_por_turno,
       updated_at = excluded.updated_at,
       updated_by = excluded.updated_by`
  )
    .bind(dateISO, valor, turnosPorDia, horasPorTurno, updatedBy)
    .run();
}

export async function upsertMetaMes(env: Env, yearMonth: string, valor: number, updatedBy: string): Promise<void> {
  await env.DB.prepare(
    `INSERT INTO metas (escopo, referencia, valor, updated_at, updated_by)
     VALUES ('mes', ?, ?, datetime('now'), ?)
     ON CONFLICT(escopo, referencia) DO UPDATE SET
       valor = excluded.valor,
       updated_at = excluded.updated_at,
       updated_by = excluded.updated_by`
  )
    .bind(yearMonth, valor, updatedBy)
    .run();
}
