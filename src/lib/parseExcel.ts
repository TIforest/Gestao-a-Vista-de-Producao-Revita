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

export interface TotalDia {
  peso: number;
  linhas: number;
}

export interface ParseResult {
  /** Linhas dentro da janela recente (com row_hash), prontas pra gravar em `apontamentos`. */
  rows: Apontamento[];
  warnings: string[];
  /** Total de peso/linhas por dia (YYYY-MM-DD) de TODAS as linhas válidas lidas, não só da janela. */
  totaisPorDia: Map<string, TotalDia>;
  /** Menor data entre as linhas lidas — diz até onde a leitura "cobre" a planilha. */
  diaMaisAntigo: string | null;
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
  const totaisPorDia = new Map<string, TotalDia>();
  if (table.length < 2) {
    return { rows: [], warnings: ["Planilha vazia ou sem linhas de dados."], totaisPorDia, diaMaisAntigo: null };
  }

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
  let diaMaisAntigo: string | null = null;
  // Quantas vezes a mesma linha (todos os campos iguais) já apareceu — duas
  // linhas idênticas viram "#1" e "#2" e contam as duas, como o BI faz.
  const ocorrencias = new Map<string, number>();

  for (let r = 1; r < table.length; r++) {
    const line = table[r] as unknown[];
    if (!line || line.every((c) => c === null || c === "")) continue;

    const get = (idx: number) => (idx === -1 ? null : line[idx] ?? null);

    const dateCell = dateOnlyCol !== -1 ? get(dateOnlyCol) : null;
    const timeCell = get(col.data_hora);
    const dataHora = combineDateTimeCells(dateCell, timeCell, today);
    const foraDaJanela = !!windowStartISO && dataHora !== null && dataHora < windowStartISO;
    if (!dataHora) {
      warnings.push(`Linha ${r + 1}: data/hora inválida ou não reconhecida, registro ignorado.`);
      continue;
    }

    const loteVal = String(get(col.lote) ?? "").trim();
    const turmaVal = String(get(col.turma) ?? "").trim().toUpperCase();
    const maquinaVal = normalizeMaquina(get(col.maquina));
    const numeroFardoRaw = get(col.numero_fardo);
    const numeroFardo = numeroFardoRaw === null ? null : toNumber(numeroFardoRaw);
    const peso = toNumber(get(col.peso_seco));
    const pesoFinal = Number.isFinite(peso) ? peso : 0;
    const cliente = String(get(col.cliente) ?? "").trim();
    const produto = String(get(col.produto) ?? "").trim();

    if (!loteVal || !turmaVal || !maquinaVal) {
      // Fora da janela não vale aviso: é histórico antigo, não dado novo errado.
      if (!foraDaJanela) warnings.push(`Linha ${r + 1}: faltam campos obrigatórios (lote/turma/máquina), registro ignorado.`);
      continue;
    }

    // Totais por dia de tudo que foi lido (histórico incluso): é daqui que
    // sai a produção do mês — recalculada da planilha a cada leitura, então
    // correções e exclusões feitas na planilha entram no acumulado.
    const dia = dataHora.slice(0, 10);
    const total = totaisPorDia.get(dia) ?? { peso: 0, linhas: 0 };
    total.peso += pesoFinal;
    total.linhas += 1;
    totaisPorDia.set(dia, total);
    if (diaMaisAntigo === null || dia < diaMaisAntigo) diaMaisAntigo = dia;

    // Fora da janela recente (ex.: dias anteriores) — pula ANTES do hash
    // (parte mais cara), sem gerar aviso: é poda normal, não erro de dado.
    // O painel não guarda histórico de linhas, só o suficiente pro dia/turno.
    if (foraDaJanela) continue;

    // A chave é a linha INTEIRA: qualquer correção na planilha (lote, hora,
    // peso, cliente...) vira uma linha "nova", e a versão antiga some do
    // banco na reconciliação (ver graphSync.aplicarLeitura). É isso que
    // faz o painel bater com o BI, que relê a planilha do zero.
    const chave = [loteVal, numeroFardo ?? "", turmaVal, maquinaVal, dataHora, pesoFinal, cliente, produto].join("|");
    const n = (ocorrencias.get(chave) ?? 0) + 1;
    ocorrencias.set(chave, n);
    const rowHash = await sha256Hex(`${chave}#${n}`);

    rows.push({
      lote: loteVal,
      cliente,
      numero_fardo: Number.isFinite(numeroFardo) ? (numeroFardo as number) : null,
      turma: turmaVal,
      peso_seco: pesoFinal,
      data_hora: dataHora,
      maquina: maquinaVal,
      produto,
      row_hash: rowHash,
    });
  }

  return { rows, warnings, totaisPorDia, diaMaisAntigo };
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
  if (!sheet) {
    return { rows: [], warnings: [`Aba "${targetSheet}" não encontrada na planilha.`], totaisPorDia: new Map(), diaMaisAntigo: null };
  }

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
 * buscarLinhasRecentesViaRange em graphSync.ts.
 */
export async function parseGraphRangeValues(
  headerRow: unknown[],
  dataRows: unknown[][],
  windowStartISO?: string
): Promise<ParseResult> {
  const table: unknown[][] = [headerRow, ...dataRows];
  return linhasDeTabela(table, windowStartISO);
}
