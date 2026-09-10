/** A single spreadsheet cell as returned by `read-excel-file`. */
export type Cell = string | number | boolean | Date | null | undefined;

/** Raw sheet contents: one array of cells per row, first row included. */
export type SheetData = Cell[][];

/** A problem found while reading a sheet. `location` is human-readable, e.g. "row 8, column G". */
export interface Issue {
  message: string;
  location?: string;
}

export interface ParseResult<T> {
  /** `null` when `errors` is non-empty and no usable document could be built. */
  data: T | null;
  errors: Issue[];
  warnings: Issue[];
}

export interface InvoiceMeta {
  invoiceType: string;
  invoiceNumber: string;
  buyerName: string;
  buyerAddress: string;
  buyerPhone: string;
  buyerGst: string;
  billDate: string;
  deliveryNote: string;
  modeOfPayment: string;
  supplierRef: string;
  otherRef: string;
  buyerOrderNumber: string;
  buyerOrderDate: string;
  despatchDocumentNumber: string;
  deliveryNoteDate: string;
  despatchedThrough: string;
  destination: string;
  term1: string;
  term2: string;
  term3: string;
  /** True when the sheet's same-state cell reads "yes" — drives the CGST/SGST vs IGST split. */
  sameState: boolean;
}

export interface LineItem {
  si: string;
  description: string;
  /** Kept as text so codes such as `40151900` never pick up a float suffix. */
  hsn: string;
  quantity: number;
  /** `null` when the sheet leaves the rate blank; it is displayed but never summed. */
  rate: number | null;
  per: string;
  /** `null` when the sheet's GST cell is blank or not one of 5 / 12 / 18. */
  gstPercent: number | null;
  amount: number;
}

export interface InvoiceSheet {
  meta: InvoiceMeta;
  items: LineItem[];
}

/** One active GST slab, as it appears in the tax summary table. */
export interface TaxSlab {
  rate: 5 | 12 | 18;
  /** Distinct HSN/SAC codes contributing to this slab, in first-seen order. */
  hsnCodes: string[];
  /** Back-computed from the slab tax, matching the legacy `tax_N / rate`. */
  taxableValue: number;
  tax: number;
}

export interface InvoiceTotals {
  /** Sum of the Amount column — the taxable value before GST. */
  taxableAmount: number;
  totalPieces: number;
  slabs: TaxSlab[];
  totalTax: number;
  /** `taxableAmount + totalTax`. */
  totalAmount: number;
  sameState: boolean;
}

export interface PackingLine {
  product: string;
  quantity: number;
}

export interface PackingBox {
  boxNumber: string;
  lines: PackingLine[];
}

export interface PackingSheet {
  invoiceNumber: string;
  shippingAddress: string;
  buyerName: string;
  date: string;
  contactNumber: string;
  boxes: PackingBox[];
}
