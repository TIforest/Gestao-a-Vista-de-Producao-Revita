// Metas fixas de produção (kg). O painel é só visual/informativo — quem
// decide os números é a operação, não um formulário no painel. Quando um
// valor mudar de verdade (ex: capacidade de uma desaguadora aumentou),
// atualiza aqui e faz o deploy — fica registrado no histórico do git.
//
// Confirmado com o usuário em 2026-08-25 e ajustado em 2026-10-05:
// - Cada desaguadora tem meta própria de 10.000 (não é a meta do turno dividida).
// - Meta do turno = 40.000 (soma das 4 desaguadoras, mas é um valor fixo em si).
// - Meta do dia = meta do turno × turnos por dia. O número de turnos vem da
//   var TURNOS_POR_DIA do wrangler.jsonc (hoje 2: só as turmas A e B operam),
//   igual ao BI original, que mostra 80.000 no gauge do dia. Mudou a
//   operação, muda só a var — sem mexer em código.
// - Meta por hora = meta do turno ÷ 6 (turno tem 6h — fórmula DAX do BI original).
// - Não existe meta do mês.
export const META_POR_DESAGUADORA: Record<string, number> = {
  "01": 10_000,
  "02": 10_000,
  "03": 10_000,
  "04": 10_000,
};

export const META_TURNO = 40_000;
export const HORAS_POR_TURNO = 6;
export const META_HORA = META_TURNO / HORAS_POR_TURNO;

// Usado se a var TURNOS_POR_DIA não existir ou vier inválida.
export const TURNOS_POR_DIA_PADRAO = 2;

export function getMetaDia(turnosPorDia: number): number {
  return META_TURNO * turnosPorDia;
}

export function getMetaPorDesaguadora(maquina: string): number {
  return META_POR_DESAGUADORA[maquina] ?? 0;
}
