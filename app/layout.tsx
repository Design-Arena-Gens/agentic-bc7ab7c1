import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Outlook Startup Automation Helper",
  description:
    "Generate a Windows startup script that launches Outlook and sends an email automatically."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="fr">
      <body>{children}</body>
    </html>
  );
}
