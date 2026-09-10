/**
 * Regenerates `lib/signatureImage.ts` from `public/sign.jpg`.
 *
 * The signature is inlined rather than fetched so that generating a PDF makes no
 * network request. Run this after replacing the scan.
 */
import { readFileSync, writeFileSync, statSync } from "node:fs";

const SOURCE = "public/sign.jpg";
const TARGET = "lib/signatureImage.ts";

const base64 = readFileSync(SOURCE).toString("base64");
const chunks = base64.match(/.{1,110}/g) ?? [base64];

const file = `/**
 * The authorised-signature scan (\`${SOURCE}\`) inlined as a data URI.
 *
 * pdfmake needs image bytes at document-build time. Inlining keeps the promise
 * that generating a PDF issues no network request at all — the uploaded sheet
 * and everything derived from it stay in the browser.
 *
 * Regenerate with: npm run build:signature
 */
export const SIGNATURE_DATA_URL =
  "data:image/jpeg;base64,${chunks.join('" +\n  "')}";
`;

writeFileSync(TARGET, file);
console.log(`${TARGET}: ${statSync(TARGET).size} bytes from ${SOURCE}`);
