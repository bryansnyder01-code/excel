import type { Computed } from "./formula";
import type { NumberFormat, SheetSpec } from "./types";

/** Currency symbols, the same table as backend/app/services/workbook_builder.py. */
export const CURRENCIES: { code: string; symbol: string; label: string; decimals?: number }[] = [
  { code: "USD", symbol: "$", label: "US dollar" },
  { code: "EUR", symbol: "€", label: "Euro" },
  { code: "GBP", symbol: "£", label: "British pound" },
  { code: "JPY", symbol: "¥", label: "Japanese yen", decimals: 0 },
  { code: "INR", symbol: "₹", label: "Indian rupee" },
  { code: "CNY", symbol: "CN¥", label: "Chinese yuan" },
  { code: "CAD", symbol: "CA$", label: "Canadian dollar" },
  { code: "AUD", symbol: "A$", label: "Australian dollar" },
  { code: "CHF", symbol: "CHF ", label: "Swiss franc" },
  { code: "MXN", symbol: "MX$", label: "Mexican peso" },
  { code: "BRL", symbol: "R$", label: "Brazilian real" },
  { code: "ZAR", symbol: "R", label: "South African rand" },
];

export interface NumberStyle {
  fmt: NumberFormat;
  decimals?: number | null;
  currency?: string | null;
}

export interface Display {
  text: string;
  numeric: boolean;
  negative: boolean;
  error: boolean;
}

const currencyOf = (code?: string | null) => CURRENCIES.find((c) => c.code === code) ?? CURRENCIES[0];

/** Decimal places a format shows by default. */
export function defaultDecimals(fmt: NumberFormat, currency?: string | null): number {
  if (fmt === "currency") return currencyOf(currency).decimals ?? 2;
  if (fmt === "decimal") return 2;
  if (fmt === "percent") return 1;
  return 0;
}

const grouped = (n: number, decimals: number) =>
  new Intl.NumberFormat("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(n);

/** The number format for grid cell (r, c), r >= 1: the cell's own settings over the column's. */
export function numberStyleAt(sheet: SheetSpec, r: number, c: number): NumberStyle {
  if (r === 0) return { fmt: "text" };
  const cell = sheet.rows[r - 1]?.cells[c];
  const col = sheet.columns[c];
  return {
    fmt: cell?.format ?? col?.format ?? "text",
    decimals: cell?.decimals ?? col?.decimals ?? null,
    currency: cell?.currency ?? col?.currency ?? null,
  };
}

/** Render a cell's value (or computed formula result) the way Excel would show it. */
export function formatValue(value: Computed | undefined, style: NumberFormat | NumberStyle): Display {
  const { fmt, decimals, currency } = typeof style === "string" ? { fmt: style } as NumberStyle : style;
  if (value === null || value === undefined || value === "") return { text: "", numeric: false, negative: false, error: false };
  if (typeof value === "object") return { text: value.error, numeric: false, negative: false, error: true };
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE", numeric: false, negative: false, error: false };
  if (typeof value === "string") {
    if (fmt === "date" && /^\d{4}-\d{2}-\d{2}$/.test(value)) return { text: value, numeric: true, negative: false, error: false };
    return { text: value, numeric: false, negative: false, error: false };
  }
  const negative = value < 0;
  const places = decimals ?? defaultDecimals(fmt, currency);
  let text: string;
  switch (fmt) {
    case "currency": {
      const body = `${currencyOf(currency).symbol}${grouped(Math.abs(value), places)}`;
      text = negative ? `-${body}` : body;
      break;
    }
    case "percent":
      text = `${(value * 100).toFixed(places)}%`;
      break;
    case "integer":
      text = grouped(value, 0);
      break;
    case "date":
      text = excelSerialToIso(value);
      break;
    case "decimal":
      text = grouped(value, places);
      break;
    default:
      text = Number.isInteger(value) ? String(value) : String(Math.round(value * 1e6) / 1e6);
  }
  return { text, numeric: true, negative, error: false };
}

function excelSerialToIso(serial: number): string {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  return new Date(ms).toISOString().slice(0, 10);
}

/** Parse what the user typed into a cell: numbers, percentages, booleans, formulas, text. */
export function parseInput(input: string): string | number | boolean | null {
  const s = input.trim();
  if (s === "") return null;
  if (s.startsWith("=")) return s;
  if (/^(true|false)$/i.test(s)) return s.toLowerCase() === "true";
  const plain = s.replace(/CN¥|CA\$|A\$|MX\$|R\$|CHF\s?|[$,€£¥₹]/g, "");
  if (/^-?\d+(\.\d+)?%$/.test(plain)) return parseFloat(plain) / 100;
  if (/^-?\d+(\.\d+)?$/.test(plain)) return parseFloat(plain);
  return input;
}

/** The number format implied by how something was typed: "15%" is a percentage, "€9" is euros. */
export function typedFormat(input: string): { format: NumberFormat; currency?: string } | null {
  const s = input.trim();
  if (s.startsWith("=")) return null;
  if (/^-?[\d,]+(\.\d+)?%$/.test(s)) return { format: "percent" };
  const cur = CURRENCIES.find((c) => c.code !== "USD" && s.replace(/^-/, "").startsWith(c.symbol.trim()) && /\d/.test(s));
  if (cur && /^-?\D+[\d,]+(\.\d+)?$/.test(s)) return { format: "currency", currency: cur.code };
  if (/^-?\$[\d,]+(\.\d+)?$/.test(s)) return { format: "currency" };
  return null;
}

export function columnLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
