import Link from "next/link";

export default function NotFound() {
  return (
    <section className="panel">
      <h2 style={{ marginTop: 0, fontSize: 18 }}>Page not found</h2>
      <p className="hint">
        Try the <Link href="/invoice">tax invoice</Link> or the{" "}
        <Link href="/packing">packing list</Link> generator.
      </p>
    </section>
  );
}
