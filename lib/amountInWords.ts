import { ToWords } from "to-words";

const toWords = new ToWords({ localeCode: "en-IN" });

/**
 * `to-words` renders the Indian lakh/crore system but omits the "and" that
 * `num2words(n, lang='en_IN')` inserted before a trailing remainder under 100
 * ("Six Hundred And Five"). The legacy invoices carry that wording, so it is
 * restored here: when the number is at least 100 and its last two digits are
 * non-zero, "And" goes in front of the final chunk.
 */
function numberToWords(value: number): string {
  const n = Math.abs(Math.trunc(value));
  const base = toWords.convert(n);
  const remainder = n % 100;

  let text = base;
  if (n >= 100 && remainder > 0) {
    const remainderWords = toWords.convert(remainder);
    if (text.endsWith(remainderWords)) {
      text = `${text.slice(0, text.length - remainderWords.length)}And ${remainderWords}`;
    }
  }

  // num2words joined compound numbers with hyphens and groups with commas; the
  // legacy code stripped both before printing.
  return text.replace(/-/g, " ").replace(/,/g, "").replace(/\s+/g, " ").trim();
}

/**
 * The exact wording the legacy invoices used:
 * `INR <rupees> Rupees and <paisa> Paisa Only`.
 *
 * Splitting on the rounded paise (rather than `Math.round((n - trunc(n)) * 100)`)
 * keeps 1.999 from printing as "One Hundred Paisa".
 */
export function amountInWords(value: number): string {
  const negative = value < 0;
  const paise = Math.round(Math.abs(value) * 100);
  const rupees = Math.floor(paise / 100);
  const remainder = paise % 100;

  const words = `INR ${numberToWords(rupees)} Rupees and ${numberToWords(remainder)} Paisa Only`;
  return negative ? `Minus ${words}` : words;
}
