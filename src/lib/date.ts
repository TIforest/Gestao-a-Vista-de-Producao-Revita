// Fábrica opera em horário de Brasília (UTC-3, sem horário de verão desde 2019).
// Evitamos Intl/timezone dinâmico: o offset é fixo e conhecido.
const BR_OFFSET_MS = 3 * 60 * 60 * 1000;

/** "Agora" em horário de Brasília, como Date (os campos UTC do objeto já refletem o horário local). */
export function nowBrazil(): Date {
  return new Date(Date.now() - BR_OFFSET_MS);
}

/** Data local (Brasília) no formato YYYY-MM-DD. */
export function todayBrazilISODate(): string {
  return nowBrazil().toISOString().slice(0, 10);
}

/** ISO local (Brasília) de N dias atrás, à meia-noite — usado como início da janela de retenção. */
export function daysAgoLocalISOStart(days: number): string {
  const d = nowBrazil();
  d.setUTCDate(d.getUTCDate() - days);
  return d.toISOString().slice(0, 10) + "T00:00:00";
}

/** Início/fim (exclusivo) do dia informado (YYYY-MM-DD), em ISO local "sem Z" — comparável a data_hora armazenada. */
export function dayBoundsLocal(dateISO: string): { start: string; end: string } {
  const start = `${dateISO}T00:00:00`;
  const parts = dateISO.split("-").map(Number);
  const y = parts[0] ?? 1970;
  const m = parts[1] ?? 1;
  const d = parts[2] ?? 1;
  const next = new Date(Date.UTC(y, m - 1, d + 1));
  const end = next.toISOString().slice(0, 10) + "T00:00:00";
  return { start, end };
}

/** Início/fim (exclusivo) do mês informado (YYYY-MM), em ISO local "sem Z". */
export function monthBoundsLocal(yearMonth: string): { start: string; end: string } {
  const parts = yearMonth.split("-").map(Number);
  const y = parts[0] ?? 1970;
  const m = parts[1] ?? 1;
  const start = `${yearMonth}-01T00:00:00`;
  const next = new Date(Date.UTC(y, m, 1));
  const end = next.toISOString().slice(0, 7) + "-01T00:00:00";
  return { start, end };
}

interface DateParts {
  y: number;
  m: number; // 1-12
  d: number;
}
interface TimeParts {
  h: number;
  min: number;
  s: number;
}

// Serial de data do Excel -> dias desde a época Unix. 25569 é a distância
// entre a época do Excel (30/12/1899, já compensando o bug do "ano bissexto
// 1900" que o próprio Excel tem) e 01/01/1970. Fórmula padrão, mesma que a
// maioria das bibliotecas (incluindo o SheetJS) usa internamente.
function excelSerialToUTCDate(serial: number): Date {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

/**
 * Extrai ano/mês/dia de uma célula de data. Aceita Date "naive-UTC" (vindo
 * do SheetJS com cellDates+raw:true), texto, ou o número de série do Excel
 * (vindo direto da API de Range do Graph, sem passar pelo SheetJS).
 */
export function dateParts(cell: unknown): DateParts | null {
  if (cell instanceof Date) {
    return { y: cell.getUTCFullYear(), m: cell.getUTCMonth() + 1, d: cell.getUTCDate() };
  }
  if (typeof cell === "number" && Number.isFinite(cell)) {
    // Só a parte inteira (dias) importa aqui — a fração (hora do dia), se
    // houver, fica pro timeParts.
    const dt = excelSerialToUTCDate(Math.floor(cell));
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  if (typeof cell === "string") {
    const s = cell.trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); // ISO: YYYY-MM-DD
    if (m) return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // BR: DD/MM/AAAA
    if (m) return { y: Number(m[3]), m: Number(m[2]), d: Number(m[1]) };
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/); // US: M/D/AA (formato real da planilha Revita)
    if (m) return { y: 2000 + Number(m[3]), m: Number(m[1]), d: Number(m[2]) };
  }
  return null;
}

/** Extrai hora/minuto/segundo de uma célula de hora (Date "naive-UTC", número de série do Excel, ou texto "HH:MM[:SS][ AM/PM]"). */
export function timeParts(cell: unknown): TimeParts | null {
  if (cell instanceof Date) {
    return { h: cell.getUTCHours(), min: cell.getUTCMinutes(), s: cell.getUTCSeconds() };
  }
  if (typeof cell === "number" && Number.isFinite(cell)) {
    // Só a parte fracionária (fração do dia) importa aqui.
    const frac = cell - Math.floor(cell);
    const totalSegundos = Math.round(frac * 86400);
    return { h: Math.floor(totalSegundos / 3600), min: Math.floor((totalSegundos % 3600) / 60), s: totalSegundos % 60 };
  }
  if (typeof cell === "string") {
    const m = cell.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
    if (m) {
      let h = Number(m[1]);
      const ampm = m[4]?.toUpperCase();
      if (ampm === "PM" && h < 12) h += 12;
      if (ampm === "AM" && h === 12) h = 0;
      return { h, min: Number(m[2]), s: m[3] ? Number(m[3]) : 0 };
    }
  }
  return null;
}

const pad2 = (n: number): string => String(n).padStart(2, "0");

/** Combina uma célula de data e uma de hora (ou uma célula única com os dois) em ISO local "sem Z". */
export function combineDateTimeCells(dateCell: unknown, timeCell: unknown, fallbackDateISO?: string): string | null {
  const d = dateParts(dateCell) ?? dateParts(timeCell) ?? (fallbackDateISO ? dateParts(fallbackDateISO) : null);
  if (!d) return null;
  const t = timeParts(timeCell) ?? timeParts(dateCell) ?? { h: 0, min: 0, s: 0 };
  return `${d.y}-${pad2(d.m)}-${pad2(d.d)}T${pad2(t.h)}:${pad2(t.min)}:${pad2(t.s)}`;
}
