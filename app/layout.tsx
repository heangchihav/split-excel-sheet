import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Sheet Splitter - Split Sheets into Separate Workbooks",
  description:
    "Upload your master Excel file and automatically split sheets into separate workbooks inside custom named folders with prefix and sequencing options.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col bg-slate-50 text-slate-900 font-sans">
        {children}
      </body>
    </html>
  );
}
