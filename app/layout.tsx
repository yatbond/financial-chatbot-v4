import type { Metadata } from 'next'
import { Inter } from 'next/font/google'
import { ClerkProvider } from '@clerk/nextjs'
import './globals.css'

const inter = Inter({ subsets: ['latin'] })

// Force dynamic rendering to avoid Clerk prerender errors during build
export const dynamic = 'force-dynamic'

export const metadata: Metadata = {
  title: 'Financial Chatbot',
  description: 'AI-powered financial data analysis for construction projects',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <ClerkProvider>
      <html lang="en">
        <body className={inter.className}>{children}</body>
      </html>
    </ClerkProvider>
  )
}
