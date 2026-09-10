import type { TDocumentDefinitions } from "pdfmake/interfaces";

/**
 * The bits of the pdfmake browser build this app uses. pdfmake's own types
 * describe a namespace, but the browser bundle exports a class *instance* whose
 * methods live on the prototype, so the runtime surface is declared here.
 */
interface PdfMakeRuntime {
  createPdf(documentDefinition: TDocumentDefinitions): {
    /** pdfmake 0.3 returns a Promise here; 0.2 took a callback. */
    getBlob(): Promise<Blob>;
  };
  addVirtualFileSystem(vfs: Record<string, string>): void;
}

let runtime: Promise<PdfMakeRuntime> | null = null;

/**
 * pdfmake reaches for browser globals as soon as it is evaluated, so it must be
 * imported lazily from a client component rather than at module scope — during
 * SSR or prerender the import would throw.
 *
 * The fonts ship separately in a virtual file system (Roboto, which unlike
 * ReportLab's Helvetica has a ₹ glyph).
 */
async function loadPdfMake(): Promise<PdfMakeRuntime> {
  if (!runtime) {
    runtime = (async () => {
      const [pdfMakeModule, vfsModule] = await Promise.all([
        import("pdfmake/build/pdfmake"),
        import("pdfmake/build/vfs_fonts"),
      ]);

      const pdfMake = ((pdfMakeModule as { default?: unknown }).default ??
        pdfMakeModule) as PdfMakeRuntime;
      const vfs = ((vfsModule as { default?: unknown }).default ?? vfsModule) as Record<
        string,
        string
      >;

      pdfMake.addVirtualFileSystem(vfs);
      return pdfMake;
    })();
  }
  return runtime;
}

/**
 * Render a document definition to a Blob in the browser. Nothing is written to
 * disk or uploaded anywhere — the legacy app's fixed output path meant concurrent
 * users overwrote each other's PDFs (§7 bug 4).
 */
export async function renderPdfBlob(doc: TDocumentDefinitions): Promise<Blob> {
  const pdfMake = await loadPdfMake();
  return pdfMake.createPdf(doc).getBlob();
}
