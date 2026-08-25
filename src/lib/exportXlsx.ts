import * as XLSX from "xlsx";
import type { DashboardPayload } from "./dashboard";

export function buildExportWorkbook(payload: DashboardPayload, rows: Record<string, unknown>[]): Uint8Array {
  const wb = XLSX.utils.book_new();

  const resumo = [
    ["Gestão à Vista de Produção - Revita"],
    ["Data de referência", payload.data],
    ["Turma filtrada", payload.filtros.turma ?? "Todas"],
    ["Desaguadora filtrada", payload.filtros.maquina ?? "Todas"],
    [],
    ["Produção acumulada do mês", payload.producaoMes],
    ["Produção total do dia", payload.producaoDia],
    ["Meta do dia", payload.metaDia],
    ["Produção total do turno/turma", payload.producaoTurno],
    ["Meta do turno", payload.metaTurno],
    ["Produção média por hora", payload.producaoMediaHora],
    ["Meta por hora", payload.metaHora],
    ["% da meta atingida", `${payload.percentualMetaAtingida.toFixed(1)}%`],
  ];
  const resumoSheet = XLSX.utils.aoa_to_sheet(resumo);
  XLSX.utils.book_append_sheet(wb, resumoSheet, "Resumo");

  const porTurmaSheet = XLSX.utils.json_to_sheet(payload.producaoPorTurma);
  XLSX.utils.book_append_sheet(wb, porTurmaSheet, "Produção x Turma");

  const porDesaguadoraSheet = XLSX.utils.json_to_sheet(payload.producaoPorDesaguadora);
  XLSX.utils.book_append_sheet(wb, porDesaguadoraSheet, "Produção x Desaguadoras");

  const apontamentosSheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, apontamentosSheet, "Apontamentos");

  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(out);
}
