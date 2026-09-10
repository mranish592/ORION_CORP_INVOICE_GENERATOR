"use client";

export default function ErrorBoundary({
  error,
  reset,
}: {
  error: Error & { digest?: string };
  reset: () => void;
}) {
  return (
    <section className="panel">
      <h2 style={{ marginTop: 0, fontSize: 18 }}>Something went wrong</h2>
      <p className="hint">
        The page hit an unexpected error. Nothing was uploaded — your spreadsheet
        never left this browser.
      </p>
      <div className="notice error">
        <code>{error.message}</code>
      </div>
      <div className="actions">
        <button type="button" className="button" onClick={reset}>
          Try again
        </button>
      </div>
    </section>
  );
}
