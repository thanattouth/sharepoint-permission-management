import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "SharePoint Permission Management",
  description: "Manage SharePoint library permissions with Microsoft Graph.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
