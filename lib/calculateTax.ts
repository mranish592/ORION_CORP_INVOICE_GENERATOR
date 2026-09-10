import { round2 } from "./cells";
import type { InvoiceTotals, LineItem, TaxSlab } from "./types";

/** The only GST rates the invoice knows how to tax. */
export const GST_SLABS = [5, 12, 18] as const;

export type GstSlab = (typeof GST_SLABS)[number];

export function isGstSlab(value: number | null): value is GstSlab {
  return value !== null && (GST_SLABS as readonly number[]).includes(value);
}

/**
 * Port of the legacy `calculate_table`, minus its module-level globals (§7 bug 3):
 * every total lives in this call.
 *
 * A row whose GST% is not exactly 5, 12 or 18 still contributes to the taxable
 * amount and the piece count but generates no tax — the parser raises a warning
 * for those rows so the under-taxing is visible rather than silent.
 *
 * Rounding happens once, after summation, never per row.
 */
export function calculateInvoiceTotals(
  items: LineItem[],
  sameState: boolean,
): InvoiceTotals {
  let taxableAmount = 0;
  let totalPieces = 0;

  const taxByRate = new Map<GstSlab, number>();
  const codesByRate = new Map<GstSlab, string[]>();
  for (const rate of GST_SLABS) {
    taxByRate.set(rate, 0);
    codesByRate.set(rate, []);
  }

  for (const item of items) {
    taxableAmount += item.amount;
    totalPieces += item.quantity;

    if (!isGstSlab(item.gstPercent)) continue;
    const rate = item.gstPercent;
    taxByRate.set(rate, taxByRate.get(rate)! + (rate / 100) * item.amount);

    const codes = codesByRate.get(rate)!;
    if (item.hsn !== "" && !codes.includes(item.hsn)) codes.push(item.hsn);
  }

  const slabs: TaxSlab[] = [];
  let totalTax = 0;
  for (const rate of GST_SLABS) {
    const raw = taxByRate.get(rate)!;
    if (raw === 0) continue;
    totalTax += raw;
    slabs.push({
      rate,
      hsnCodes: codesByRate.get(rate)!,
      taxableValue: round2(raw / (rate / 100)),
      tax: round2(raw),
    });
  }

  return {
    taxableAmount: round2(taxableAmount),
    totalPieces: round2(totalPieces),
    slabs,
    totalTax: round2(totalTax),
    totalAmount: round2(taxableAmount + totalTax),
    sameState,
  };
}

export interface TaxRow {
  label: string;
  amount: number;
}

/**
 * The per-slab rows appended under the line items, and repeated as the rate
 * columns of the tax summary. Intra-state splits each slab into equal CGST and
 * SGST halves; anything else charges IGST at the full slab rate.
 */
export function taxRowsFor(totals: InvoiceTotals): TaxRow[] {
  const rows: TaxRow[] = [];
  for (const slab of totals.slabs) {
    if (totals.sameState) {
      const half = round2(slab.tax / 2);
      rows.push({ label: `CGST ${slab.rate / 2}%`, amount: half });
      rows.push({ label: `SGST ${slab.rate / 2}%`, amount: half });
    } else {
      rows.push({ label: `IGST ${slab.rate}%`, amount: slab.tax });
    }
  }
  return rows;
}
