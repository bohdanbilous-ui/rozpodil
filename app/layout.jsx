import "./globals.css";

export const metadata = {
  title: "Переведення між проєктами",
  description: "Облік переведень співробітників розробки між проєктами",
};

export default function RootLayout({ children }) {
  return (
    <html lang="uk">
      <body>{children}</body>
    </html>
  );
}
