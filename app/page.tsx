'use client'

import { useState } from 'react'
import ExcelProcessor from '../components/ExcelProcessor'

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-8">
      <div className="z-10 max-w-5xl w-full items-center justify-between font-mono text-sm">
        <ExcelProcessor />
      </div>
    </main>
  )
} 