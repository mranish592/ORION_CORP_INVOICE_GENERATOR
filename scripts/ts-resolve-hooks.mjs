/**
 * Bridges the gap between Next's bundler resolution and Node's ESM resolver so the
 * app's own source runs unmodified under `node --test` and `npm run render:samples`.
 *
 * Two things the bundler does that Node does not:
 *   - fills in `.ts`/`.tsx` for extensionless relative imports (`./calculateTax`);
 *   - fills in `.js` for extensionless package subpaths (`pdfmake/build/pdfmake`).
 */
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";

const RELATIVE_EXTENSIONS = [".ts", ".tsx"];
const HAS_EXTENSION = /\.[cm]?[jt]sx?$|\.json$/;

export async function resolve(specifier, context, nextResolve) {
  if (HAS_EXTENSION.test(specifier)) return nextResolve(specifier, context);

  if (specifier.startsWith(".")) {
    for (const extension of RELATIVE_EXTENSIONS) {
      const candidate = new URL(specifier + extension, context.parentURL);
      if (existsSync(fileURLToPath(candidate))) {
        return nextResolve(specifier + extension, context);
      }
    }
    return nextResolve(specifier, context);
  }

  try {
    return await nextResolve(specifier, context);
  } catch (error) {
    if (error?.code !== "ERR_MODULE_NOT_FOUND") throw error;
    return nextResolve(`${specifier}.js`, context);
  }
}
