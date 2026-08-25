// Metas fixas de produção (kg). O painel é só visual/informativo — quem
// decide os números é a operação, não um formulário no painel. Quando um
// valor mudar de verdade (ex: capacidade de uma desaguadora aumentou),
// atualiza aqui e faz o deploy — fica registrado no histórico do git.
//
// Confirmado com o usuário em 2026-08-25:
// - Cada desaguadora tem meta própria de 10.000 (não é a meta do turno dividida).
// - Meta do turno = 40.000 (soma das 4 desaguadoras, mas é um valor fixo em si).
// - Meta do dia = 120.000 (não é meta do turno × turnos por dia — valor fixo à parte).
// - Não existe meta por hora nem meta do mês.
export const META_POR_DESAGUADORA: Record<string, number> = {
  "01": 10_000,
  "02": 10_000,
  "03": 10_000,
  "04": 10_000,
};

export const META_TURNO = 40_000;
export const META_DIA = 120_000;

export function getMetaPorDesaguadora(maquina: string): number {
  return META_POR_DESAGUADORA[maquina] ?? 0;
}
