"use client";

import { useCallback, useEffect, useRef, useState } from "react";

import { IssueList } from "@/components/IssueList";
import { amountInWords } from "@/lib/amountInWords";
import { buildInvoiceDocDefinition } from "@/lib/buildInvoiceDocDefinition";
import { buildPackingDocDefinition } from "@/lib/buildPackingDocDefinition";
import { calculateInvoiceTotals } from "@/lib/calculateTax";
import { formatMoney, formatQuantity } from "@/lib/format";
import { parseInvoiceSheet } from "@/lib/parseInvoiceSheet";
import { parsePackingSheet } from "@/lib/parsePackingSheet";
import { renderPdfBlob } from "@/lib/pdfClient";
import { readSpreadsheet } from "@/lib/readSpreadsheet";
import { SIGNATURE_DATA_URL } from "@/lib/signatureImage";
import type { InvoiceSheet, InvoiceTotals, Issue, PackingSheet } from "@/lib/types";

export type GeneratorKind = "invoice" | "packing";

type Parsed =
  | { kind: "invoice"; sheet: InvoiceSheet; totals: InvoiceTotals }
  | { kind: "packing"; sheet: PackingSheet };

interface SummaryEntry {
  label: string;
  value: string;
  wide?: boolean;
}

export function GeneratorClient({ kind }: { kind: GeneratorKind }) {
  const [fileName, setFileName] = useState<string | null>(null);
  const [parsed, setParsed] = useState<Parsed | null>(null);
  const [errors, setErrors] = useState<Issue[]>([]);
  const [warnings, setWarnings] = useState<Issue[]>([]);
  const [reading, setReading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);

  const pdfUrlRef = useRef<string | null>(null);

  const releasePdf = useCallback(() => {
    if (pdfUrlRef.current) {
      URL.revokeObjectURL(pdfUrlRef.current);
      pdfUrlRef.current = null;
    }
    setPdfUrl(null);
  }, []);

  useEffect(() => releasePdf, [releasePdf]);

  const handleFile = useCallback(
    async (file: File | null) => {
      releasePdf();
      setParsed(null);
      setErrors([]);
      setWarnings([]);
      setFileName(file?.name ?? null);
      if (!file) return;

      setReading(true);
      try {
        const rows = await readSpreadsheet(file);

        if (kind === "invoice") {
          const result = parseInvoiceSheet(rows);
          setErrors(result.errors);
          setWarnings(result.warnings);
          if (result.data) {
            setParsed({
              kind: "invoice",
              sheet: result.data,
              totals: calculateInvoiceTotals(result.data.items, result.data.meta.sameState),
            });
          }
        } else {
          const result = parsePackingSheet(rows);
          setErrors(result.errors);
          setWarnings(result.warnings);
          if (result.data) setParsed({ kind: "packing", sheet: result.data });
        }
      } catch (error) {
        setErrors([{ message: describeReadFailure(error) }]);
      } finally {
        setReading(false);
      }
    },
    [kind, releasePdf],
  );

  const handleGenerate = useCallback(async () => {
    if (!parsed) return;
    releasePdf();
    setGenerating(true);
    try {
      const doc =
        parsed.kind === "invoice"
          ? buildInvoiceDocDefinition({
              sheet: parsed.sheet,
              totals: parsed.totals,
              signature: SIGNATURE_DATA_URL,
            })
          : buildPackingDocDefinition(parsed.sheet);

      const blob = await renderPdfBlob(doc);
      const url = URL.createObjectURL(blob);
      pdfUrlRef.current = url;
      setPdfUrl(url);
    } catch (error) {
      setErrors((current) => [
        ...current,
        {
          message: `The PDF could not be rendered: ${
            error instanceof Error ? error.message : String(error)
          }`,
        },
      ]);
    } finally {
      setGenerating(false);
    }
  }, [parsed, releasePdf]);

  const summary = parsed ? summarise(parsed) : [];
  const downloadName = parsed ? downloadFileName(parsed) : "document.pdf";

  return (
    <>
      <section className="panel">
        <p className="step-label">Step 1 — choose a sheet</p>
        <div className="dropzone">
          <input
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            aria-label="Spreadsheet to convert"
            onChange={(event) => void handleFile(event.target.files?.[0] ?? null)}
          />
          {fileName ? <p className="filename">{fileName}</p> : null}
        </div>

        {reading ? <p className="hint" style={{ marginTop: 14 }}>Reading the sheet…</p> : null}

        <IssueList title="Fix these before generating" issues={errors} tone="error" />
        <IssueList title="Worth checking" issues={warnings} tone="warning" />
      </section>

      {parsed ? (
        <section className="panel">
          <p className="step-label">Step 2 — check the figures</p>
          <ul className="summary">
            {summary.map((entry) => (
              <li key={entry.label} className={entry.wide ? "wide" : undefined}>
                <span>{entry.label}</span>
                <span className="value">{entry.value}</span>
              </li>
            ))}
          </ul>

          <div className="actions">
            <button
              type="button"
              className="button"
              onClick={() => void handleGenerate()}
              disabled={generating}
            >
              {generating ? "Generating…" : pdfUrl ? "Regenerate PDF" : "Generate PDF"}
            </button>
            {pdfUrl ? (
              <a className="button secondary" href={pdfUrl} download={downloadName}>
                Download {downloadName}
              </a>
            ) : null}
          </div>

          {pdfUrl ? (
            <iframe className="preview" src={pdfUrl} title="Generated PDF preview" />
          ) : null}
        </section>
      ) : null}
    </>
  );
}

function summarise(parsed: Parsed): SummaryEntry[] {
  if (parsed.kind === "packing") {
    const { sheet } = parsed;
    const products = sheet.boxes.reduce((total, box) => total + box.lines.length, 0);
    const pieces = sheet.boxes.reduce(
      (total, box) => total + box.lines.reduce((sum, line) => sum + line.quantity, 0),
      0,
    );
    return [
      { label: "Invoice no.", value: sheet.invoiceNumber || "—" },
      { label: "Buyer", value: sheet.buyerName || "—" },
      { label: "Date", value: sheet.date || "—" },
      { label: "Boxes", value: formatQuantity(sheet.boxes.length) },
      { label: "Product lines", value: formatQuantity(products) },
      { label: "Total pieces", value: formatQuantity(pieces) },
    ];
  }

  const { sheet, totals } = parsed;
  const entries: SummaryEntry[] = [
    { label: "Invoice no.", value: sheet.meta.invoiceNumber || "—" },
    { label: "Buyer", value: sheet.meta.buyerName || "—" },
    { label: "Line items", value: formatQuantity(sheet.items.length) },
    { label: "Total pieces", value: formatQuantity(totals.totalPieces) },
    { label: "Taxable value", value: formatMoney(totals.taxableAmount) },
  ];

  for (const slab of totals.slabs) {
    entries.push({
      label: totals.sameState
        ? `CGST + SGST @ ${slab.rate}%`
        : `IGST @ ${slab.rate}%`,
      value: formatMoney(slab.tax),
    });
  }

  entries.push(
    { label: "Total tax", value: formatMoney(totals.totalTax) },
    { label: "Invoice total", value: formatMoney(totals.totalAmount) },
    { label: "Tax treatment", value: totals.sameState ? "Intra-state (CGST + SGST)" : "Inter-state (IGST)" },
    { label: "Amount in words", value: amountInWords(totals.totalAmount), wide: true },
  );

  return entries;
}

function downloadFileName(parsed: Parsed): string {
  const raw =
    parsed.kind === "invoice"
      ? `invoice-${parsed.sheet.meta.invoiceNumber || "draft"}`
      : `packing-list-${parsed.sheet.invoiceNumber || "draft"}`;
  const safe = raw.replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return `${safe || "document"}.pdf`;
}

function describeReadFailure(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (/zip|spreadsheet|xlsx|central directory/i.test(message)) {
    return "That file could not be read as an .xlsx workbook. Re-save it from Excel or Google Sheets as .xlsx and try again.";
  }
  return `The sheet could not be read: ${message}`;
}
