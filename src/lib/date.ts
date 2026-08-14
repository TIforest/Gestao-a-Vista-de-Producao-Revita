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

/** Mês local (Brasília) no formato YYYY-MM. */
export function currentBrazilYearMonth(): string {
  return nowBrazil().toISOString().slice(0, 7);
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

/** Converte um valor de célula Excel (Date já parseado pelo SheetJS com cellDates, ou string) para ISO local sem timezone. */
export function excelValueToLocalISO(value: unknown, fallbackDateISO?: string): string | null {
  if (value instanceof Date) {
    // SheetJS entrega o Date "como se fosse UTC" representando o valor local da planilha.
    const iso = value.toISOString();
    return iso.slice(0, 19);
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    // Só hora, ex "13:38:00" — combina com a data de referência (hoje ou coluna "Data" separada).
    if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(trimmed) && fallbackDateISO) {
      const [h, min, s] = trimmed.split(":");
      return `${fallbackDateISO}T${h?.padStart(2, "0")}:${min}:${s ?? "00"}`;
    }
    // Data completa, ex "2026-08-14 13:38:00" ou "14/08/2026 13:38:00"
    const isoLike = trimmed.replace(" ", "T");
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(isoLike)) return isoLike.slice(0, 19);
    const brMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T]?(\d{1,2}:\d{2}(:\d{2})?)?/);
    if (brMatch) {
      const [, dd, mm, yyyy, time] = brMatch;
      const hhmmss = time ?? "00:00:00";
      return `${yyyy}-${mm?.padStart(2, "0")}-${dd?.padStart(2, "0")}T${hhmmss.length === 5 ? hhmmss + ":00" : hhmmss}`;
    }
  }
  return null;
}
