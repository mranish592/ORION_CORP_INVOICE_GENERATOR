/**
 * The legacy PDF printed the literal string "INR" because ReportLab's default
 * Helvetica has no rupee glyph. pdfmake's bundled Roboto does, so figures now
 * carry the real symbol. The amount-in-words line still opens with "INR", which
 * is the legally conventional wording on an Indian invoice.
 */
export const RUPEE = "₹";

const moneyFormatter = new Intl.NumberFormat("en-IN", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const quantityFormatter = new Intl.NumberFormat("en-IN", {
  maximumFractionDigits: 3,
});

/** `1234.5` -> `"1,234.50"`. */
export function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return "";
  return moneyFormatter.format(value);
}

/** `1234.5` -> `"₹1,234.50"`. */
export function formatMoney(value: number): string {
  if (!Number.isFinite(value)) return "";
  return `${RUPEE}${moneyFormatter.format(value)}`;
}

/** Quantities are counts, so they keep no forced decimals. */
export function formatQuantity(value: number): string {
  if (!Number.isFinite(value)) return "";
  return quantityFormatter.format(value);
}
