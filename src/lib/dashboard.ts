import type { Env } from "../types";
import { dayBoundsLocal, todayBrazilISODate } from "./date";
import { META_TURNO, META_DIA, getMetaPorDesaguadora } from "./metasFixas";

export interface DashboardFilters {
  turma?: string;
  maquina?: string;
  date?: string; // YYYY-MM-DD, default hoje
}

export interface DashboardPayload {
  data: string;
  filtros: { turma: string | null; maquina: string | null };
  turmasDisponiveis: string[];
  desaguadorasDisponiveis: string[];
  producaoMes: number;
  producaoDia: number;
  metaDia: number;
  producaoTurno: number;
  metaTurno: number;
  percentualMetaAtingida: number; // 0..100 (produção turno / meta turno)
  producaoPorTurma: { turma: string; valor: number; meta: number }[];
  producaoPorDesaguadora: { maquina: string; valor: number; meta: number }[];
  ultimosApontamentos: {
    lote: string;
    cliente: string;
    numero_fardo: number | null;
    turma: string;
    peso_seco: number;
    data_hora: string;
    maquina: string;
    produto: string;
  }[];
  sync: {
    ultimaSincronizacao: string | null;
    status: string;
    linhas: number;
    erro: string | null;
  };
}

export async function buildDashboardPayload(env: Env, filters: DashboardFilters): Promise<DashboardPayload> {
  const dateISO = filters.date ?? todayBrazilISODate();
  const yearMonth = dateISO.slice(0, 7);
  const turmasDisponiveis = env.TURMAS.split(",").map((s) => s.trim()).filter(Boolean);
  const desaguadorasDisponiveis = env.DESAGUADORAS.split(",").map((s) => s.trim()).filter(Boolean);

  const turma = filters.turma && turmasDisponiveis.includes(filters.turma) ? filters.turma : null;
  const maquina = filters.maquina && desaguadorasDisponiveis.includes(filters.maquina) ? filters.maquina : null;

  const { start: dayStart, end: dayEnd } = dayBoundsLocal(dateISO);

  // Vem do contador incremental (producao_mensal), não de SUM sobre
  // apontamentos — a tabela de apontamentos só guarda uma janela recente
  // (ver RETENTION_DAYS em graphSync.ts), não o mês inteiro linha a linha.
  const producaoMesRow = await env.DB.prepare("SELECT total_peso AS total FROM producao_mensal WHERE ano_mes = ?")
    .bind(yearMonth)
    .first<{ total: number }>();

  // "Produção Total do Dia": todas as turmas, respeita apenas o filtro de desaguadora.
  const diaConds = ["data_hora >= ?", "data_hora < ?"];
  const diaArgs: (string | number)[] = [dayStart, dayEnd];
  if (maquina) {
    diaConds.push("maquina = ?");
    diaArgs.push(maquina);
  }
  const producaoDiaRow = await env.DB.prepare(
    `SELECT COALESCE(SUM(peso_seco), 0) AS total FROM apontamentos WHERE ${diaConds.join(" AND ")}`
  )
    .bind(...diaArgs)
    .first<{ total: number }>();

  // "Produção Total do Turno": recorte pela turma selecionada (e desaguadora, se houver).
  const turnoConds = [...diaConds];
  const turnoArgs = [...diaArgs];
  if (turma) {
    turnoConds.push("turma = ?");
    turnoArgs.push(turma);
  }
  const producaoTurnoRow = await env.DB.prepare(
    `SELECT COALESCE(SUM(peso_seco), 0) AS total FROM apontamentos WHERE ${turnoConds.join(" AND ")}`
  )
    .bind(...turnoArgs)
    .first<{ total: number }>();

  const producaoPorTurmaRows = await env.DB.prepare(
    `SELECT turma, COALESCE(SUM(peso_seco), 0) AS total FROM apontamentos
     WHERE ${diaConds.join(" AND ")} GROUP BY turma`
  )
    .bind(...diaArgs)
    .all<{ turma: string; total: number }>();

  const porTurmaConds = ["data_hora >= ?", "data_hora < ?"];
  const porTurmaArgs: (string | number)[] = [dayStart, dayEnd];
  if (turma) {
    porTurmaConds.push("turma = ?");
    porTurmaArgs.push(turma);
  }
  const producaoPorDesaguadoraRows = await env.DB.prepare(
    `SELECT maquina, COALESCE(SUM(peso_seco), 0) AS total FROM apontamentos
     WHERE ${porTurmaConds.join(" AND ")} GROUP BY maquina`
  )
    .bind(...porTurmaArgs)
    .all<{ maquina: string; total: number }>();

  const ultimosConds = ["data_hora >= ?", "data_hora < ?"];
  const ultimosArgs: (string | number)[] = [dayStart, dayEnd];
  if (turma) {
    ultimosConds.push("turma = ?");
    ultimosArgs.push(turma);
  }
  if (maquina) {
    ultimosConds.push("maquina = ?");
    ultimosArgs.push(maquina);
  }
  const ultimosRows = await env.DB.prepare(
    `SELECT lote, cliente, numero_fardo, turma, peso_seco, data_hora, maquina, produto
     FROM apontamentos WHERE ${ultimosConds.join(" AND ")}
     ORDER BY data_hora DESC LIMIT 5`
  )
    .bind(...ultimosArgs)
    .all();

  const syncRow = await env.DB.prepare(
    "SELECT last_sync_at, last_sync_status, last_sync_rows, last_error FROM sync_state WHERE id = 1"
  ).first<{ last_sync_at: string | null; last_sync_status: string; last_sync_rows: number; last_error: string | null }>();

  // Sem turma selecionada ("Todas"), o recorte "turno" vira o total do dia
  // inteiro — então compara com a meta do dia, não a de um turno só.
  const metaTurnoAlvo = turma ? META_TURNO : META_DIA;
  const producaoTurno = producaoTurnoRow?.total ?? 0;

  return {
    data: dateISO,
    filtros: { turma, maquina },
    turmasDisponiveis,
    desaguadorasDisponiveis,
    producaoMes: producaoMesRow?.total ?? 0,
    producaoDia: producaoDiaRow?.total ?? 0,
    metaDia: META_DIA,
    producaoTurno,
    metaTurno: metaTurnoAlvo,
    // Sem limite em 100 — se passar da meta, mostra o valor real (ex: 105%).
    percentualMetaAtingida: metaTurnoAlvo > 0 ? (producaoTurno / metaTurnoAlvo) * 100 : 0,
    // Cada turma tem sua própria produção do turno, sempre comparada com a
    // meta fixa do turno (40.000) — não muda com o filtro "Todas"/turma
    // selecionada, que só afeta a meta do gauge acima.
    producaoPorTurma: turmasDisponiveis.map((t) => ({
      turma: t,
      valor: producaoPorTurmaRows.results.find((r) => r.turma === t)?.total ?? 0,
      meta: META_TURNO,
    })),
    producaoPorDesaguadora: desaguadorasDisponiveis.map((m) => ({
      maquina: m,
      valor: producaoPorDesaguadoraRows.results.find((r) => r.maquina === m)?.total ?? 0,
      meta: getMetaPorDesaguadora(m),
    })),
    ultimosApontamentos: ultimosRows.results as DashboardPayload["ultimosApontamentos"],
    sync: {
      ultimaSincronizacao: syncRow?.last_sync_at ?? null,
      status: syncRow?.last_sync_status ?? "nunca_sincronizado",
      linhas: syncRow?.last_sync_rows ?? 0,
      erro: syncRow?.last_error ?? null,
    },
  };
}
