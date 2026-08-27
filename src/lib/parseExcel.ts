import * as XLSX from "xlsx";
import type { Apontamento } from "../types";
import { combineDateTimeCells, todayBrazilISODate } from "./date";
import { sha256Hex } from "./hash";

function toNumber(raw: unknown): number {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === "number") return raw;
  const n = Number(String(raw).trim().replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

// Aliases de cabeçalho reconhecidos (sem acento, maiúsculo, aparados).
// Se a planilha do gestor usar nomes diferentes, ajuste as listas abaixo.
const HEADER_ALIASES: Record<keyof Omit<Apontamento, "row_hash">, string[]> = {
  lote: ["LOTE"],
  cliente: ["CLIENTE"],
  numero_fardo: ["NUMERO DO FARDO", "NUMERO FARDO", "FARDO"],
  turma: ["TURMA"],
  peso_seco: ["SOMA DE PESO SECO 51%", "PESO SECO 51%", "PESO SECO", "PESO LIQUIDO"],
  data_hora: ["HORA DO APONTAMENTO", "DATA HORA", "DATA/HORA", "DATA DO APONTAMENTO"],
  maquina: ["MAQUINA", "DESAGUADORA"],
  produto: ["PRODUTO"],
};
const DATE_ONLY_ALIASES = ["DATA"];

function normalizeHeader(h: string): string {
  return h
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "") // remove acentos (marcas de combinação)
    .trim()
    .toUpperCase();
}

function findColumn(headers: string[], aliases: string[]): number {
  const normalized = headers.map(normalizeHeader);
  for (const alias of aliases) {
    const idx = normalized.findIndex((h) => h === alias || h.includes(alias));
    if (idx !== -1) return idx;
  }
  return -1;
}

function normalizeMaquina(raw: unknown): string {
  const s = String(raw ?? "").trim();
  const digits = s.match(/(\d+)/)?.[1];
  if (!digits) return s.toUpperCase();
  return digits.padStart(2, "0");
}

export interface ParseResult {
  rows: Apontamento[];
  warnings: string[];
}

/**
 * Núcleo compartilhado: dado um "table" já em memória (linha 0 = cabeçalho,
 * demais = dados — não importa se veio do SheetJS lendo um .xlsx binário ou
 * direto da API de Range do Graph como JSON), monta os Apontamento[].
 * Mantido separado do "como conseguir o table" pra reaproveitar entre as
 * duas fontes sem duplicar a lógica de colunas/validação/hash.
 */
async function linhasDeTabela(table: unknown[][], windowStartISO?: string): Promise<ParseResult> {
  const warnings: string[] = [];
  if (table.length < 2) return { rows: [], warnings: ["Planilha vazia ou sem linhas de dados."] };

  const headers = (table[0] as unknown[]).map((h) => String(h ?? ""));
  const col = Object.fromEntries(
    Object.entries(HEADER_ALIASES).map(([field, aliases]) => [field, findColumn(headers, aliases)])
  ) as Record<keyof Omit<Apontamento, "row_hash">, number>;
  const dateOnlyCol = findColumn(headers, DATE_ONLY_ALIASES);

  const missing = Object.entries(col).filter(([, idx]) => idx === -1).map(([field]) => field);
  if (missing.length > 0) {
    warnings.push(
      `Colunas não encontradas no cabeçalho, ajuste os aliases em src/lib/parseExcel.ts: ${missing.join(", ")}`
    );
  }

  const today = todayBrazilISODate();
  const rows: Apontamento[] = [];

  for (let r = 1; r < table.length; r++) {
    const line = table[r] as unknown[];
    if (!line || line.every((c) => c === null || c === "")) continue;

    const get = (idx: number) => (idx === -1 ? null : line[idx] ?? null);

    const dateCell = dateOnlyCol !== -1 ? get(dateOnlyCol) : null;
    const timeCell = get(col.data_hora);
    const dataHora = combineDateTimeCells(dateCell, timeCell, today);
    if (!dataHora) {
      warnings.push(`Linha ${r + 1}: data/hora inválida ou não reconhecida, registro ignorado.`);
      continue;
    }
    // Fora da janela recente (ex.: dias anteriores) — pula ANTES do hash
    // (parte mais cara), sem gerar aviso: é poda normal, não erro de dado.
    // O painel não guarda histórico de linhas, só o suficiente pro dia/turno
    // atual; a produção do mês é somada à parte (ver producao_mensal).
    if (windowStartISO && dataHora < windowStartISO) continue;

    const loteVal = String(get(col.lote) ?? "").trim();
    const turmaVal = String(get(col.turma) ?? "").trim().toUpperCase();
    const maquinaVal = normalizeMaquina(get(col.maquina));
    const numeroFardoRaw = get(col.numero_fardo);
    const numeroFardo = numeroFardoRaw === null ? null : toNumber(numeroFardoRaw);
    const peso = toNumber(get(col.peso_seco));

    if (!loteVal || !turmaVal || !maquinaVal) {
      warnings.push(`Linha ${r + 1}: faltam campos obrigatórios (lote/turma/máquina), registro ignorado.`);
      continue;
    }

    // A chave usa só a DATA (não o horário completo): o horário do
    // apontamento às vezes é corrigido depois na planilha, e usar o
    // timestamp inteiro na chave faria essa correção virar uma linha
    // "nova" (duplicando a produção no acumulado em vez de atualizar).
    // TURMA entra na chave porque alguns lotes usam um rótulo genérico
    // recorrente (ex.: "Prensado") sem número de fardo realmente único —
    // ainda existe uma margem de erro rara (mesmo lote+fardo+turma+máquina
    // duas vezes no mesmo dia), documentado no SETUP.md.
    const rowHash = await sha256Hex(
      `${loteVal}|${numeroFardo ?? ""}|${turmaVal}|${maquinaVal}|${dataHora.slice(0, 10)}`
    );

    rows.push({
      lote: loteVal,
      cliente: String(get(col.cliente) ?? "").trim(),
      numero_fardo: Number.isFinite(numeroFardo) ? (numeroFardo as number) : null,
      turma: turmaVal,
      peso_seco: Number.isFinite(peso) ? peso : 0,
      data_hora: dataHora,
      maquina: maquinaVal,
      produto: String(get(col.produto) ?? "").trim(),
      row_hash: rowHash,
    });
  }

  return { rows, warnings };
}

/**
 * Caminho usado pelo upload manual (/api/upload): recebe o .xlsx binário
 * inteiro e usa o SheetJS pra ler. Mais pesado (processa o arquivo inteiro),
 * mas é só quando alguém sobe um arquivo de próprio punho — não roda toda
 * hora como a sincronização automática (ver parseGraphRangeValues).
 */
export async function parseWorkbook(
  buffer: ArrayBuffer,
  sheetName?: string,
  windowStartISO?: string
): Promise<ParseResult> {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
  const targetSheet = sheetName && workbook.Sheets[sheetName] ? sheetName : workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheet as string];
  if (!sheet) return { rows: [], warnings: [`Aba "${targetSheet}" não encontrada na planilha.`] };

  // raw:true entrega valores nativos (Date para células de data/hora com
  // cellDates, number para células numéricas) em vez de texto formatado —
  // muito mais confiável do que tentar decifrar strings pré-formatadas
  // (formato de data/hora e separador decimal variam por planilha/locale).
  const table: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  return linhasDeTabela(table, windowStartISO);
}

/**
 * Caminho usado pela sincronização automática: recebe só um pedaço da
 * planilha (cabeçalho + últimas N linhas) já como valores, vindo da API de
 * Range do Microsoft Graph — sem baixar/processar o arquivo inteiro. Ver
 * fetchRecentRowsViaGraphRange em graphSync.ts.
 */
export async function parseGraphRangeValues(
  headerRow: unknown[],
  dataRows: unknown[][],
  windowStartISO?: string
): Promise<ParseResult> {
  const table: unknown[][] = [headerRow, ...dataRows];
  return linhasDeTabela(table, windowStartISO);
}
