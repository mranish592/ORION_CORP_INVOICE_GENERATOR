import type { CustomTableLayout } from "pdfmake/interfaces";

/** ReportLab's `colors.grey`, the rule colour used throughout the legacy invoice. */
export const RULE_COLOR = "#808080";
export const RULE_WIDTH = 0.5;

/** A4 is 595.28pt wide; the legacy tables are 540pt, centred. */
export const CONTENT_WIDTH = 540;
export const PAGE_SIDE_MARGIN = (595.28 - CONTENT_WIDTH) / 2;

/** Horizontal breathing room inside each cell, per side. */
export const CELL_PADDING_X = 3;
const CELL_PADDING_Y = 2;

/**
 * Convert the legacy column widths — which, like ReportLab's `_argW`, describe the
 * *full* column including its padding — into the content widths pdfmake expects.
 *
 * pdfmake adds cell padding and rule widths *outside* the declared width, so
 * handing it the raw 540pt figures makes the table wider than the page: the
 * right-hand columns then wrap or fall off it entirely. Scaling keeps the relative
 * proportions of the original layout while making the drawn table exactly
 * {@link CONTENT_WIDTH}.
 */
export function contentWidths(outerWidths: number[]): number[] {
  const columns = outerWidths.length;
  const overhead = columns * CELL_PADDING_X * 2 + (columns + 1) * RULE_WIDTH;
  const scale = (CONTENT_WIDTH - overhead) / CONTENT_WIDTH;
  return outerWidths.map((width) => width * scale);
}

/** The width a table actually occupies once padding and rules are added back. */
export function drawnTableWidth(widths: number[]): number {
  const columns = widths.length;
  return (
    widths.reduce((total, width) => total + width, 0) +
    columns * CELL_PADDING_X * 2 +
    (columns + 1) * RULE_WIDTH
  );
}

const rule = () => RULE_WIDTH;
const noRule = () => 0;
const ruleColor = () => RULE_COLOR;
const paddingX = () => CELL_PADDING_X;
const paddingY = () => CELL_PADDING_Y;

function layout(overrides: CustomTableLayout): CustomTableLayout {
  return {
    hLineColor: ruleColor,
    vLineColor: ruleColor,
    paddingLeft: paddingX,
    paddingRight: paddingX,
    paddingTop: paddingY,
    paddingBottom: paddingY,
    ...overrides,
  };
}

/** Every line drawn — ReportLab's `GRID`. */
export const gridLayout: CustomTableLayout = layout({
  hLineWidth: rule,
  vLineWidth: rule,
});

/** Outline only — ReportLab's `BOX`. */
export const boxLayout: CustomTableLayout = layout({
  hLineWidth: (i, node) => (i === 0 || i === node.table.body.length ? RULE_WIDTH : 0),
  vLineWidth: (i, node) => (i === 0 || i === node.table.widths!.length ? RULE_WIDTH : 0),
});

export const noBordersLayout: CustomTableLayout = layout({
  hLineWidth: noRule,
  vLineWidth: noRule,
});

/**
 * Line-items layout. The legacy code expressed these boundaries with negative row
 * indices scaled by the number of active tax slabs; stated as a rule it is simply:
 * a full grid down to the end of the line items, a box around the summary block,
 * a grid on the final Total row, and vertical rules all the way down.
 *
 * @param summaryStartRow index of the subtotal row — the first summary row.
 */
export function lineItemsLayout(summaryStartRow: number): CustomTableLayout {
  return layout({
    vLineWidth: rule,
    hLineWidth: (i, node) => {
      const lastRow = node.table.body.length;
      if (i <= summaryStartRow) return RULE_WIDTH; // grid over header + line items
      if (i === lastRow - 1 || i === lastRow) return RULE_WIDTH; // Total row grid + box base
      return 0;
    },
  });
}

/**
 * Declaration block: an outline, a rule under the tax-amount-in-words row, and a
 * vertical rule splitting the remaining rows. The words row spans both columns, so
 * pdfmake suppresses the vertical rule inside it automatically.
 */
export const declarationLayout: CustomTableLayout = layout({
  vLineWidth: rule,
  hLineWidth: (i, node) =>
    i === 0 || i === 1 || i === node.table.body.length ? RULE_WIDTH : 0,
});

/**
 * Buyer + terms block: a grid on the first row, an outline around the whole table,
 * and a box around the address/phone/GST cells — which together means every
 * vertical rule but only the outer and first horizontal ones.
 */
export const buyerTermsLayout: CustomTableLayout = layout({
  vLineWidth: rule,
  hLineWidth: (i, node) =>
    i === 0 || i === 1 || i === node.table.body.length ? RULE_WIDTH : 0,
});
