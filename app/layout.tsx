import type { Metadata } from "next";
import Link from "next/link";

import "./globals.css";

export const metadata: Metadata = {
  title: "Orion Corp Invoice Generator",
  description:
    "Turn an Excel spreadsheet into a GST tax invoice or a packing list PDF, entirely in your browser.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        <div className="shell">
          <header className="masthead">
            <h1>
              <Link href="/" style={{ color: "inherit", textDecoration: "none" }}>
                Orion Corp
              </Link>
            </h1>
            <nav>
              <Link href="/invoice">Tax invoice</Link>
              <Link href="/packing">Packing list</Link>
            </nav>
          </header>
          <main>{children}</main>
        </div>
      </body>
    </html>
  );
}
