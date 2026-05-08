import './globals.css'

export const metadata = {
  title: 'Dashboard Atrasos Médicos',
  description: 'Dashboard Hospitalar',
}

export default function RootLayout({ children }) {
  return (
    <html lang="pt-br">
      <body>{children}</body>
    </html>
  )
}
