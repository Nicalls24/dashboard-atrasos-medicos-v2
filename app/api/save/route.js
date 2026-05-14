import { NextResponse } from 'next/server'

const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'

const sbPost = (rows, date, ts) =>
  fetch(`${SB_URL}/rest/v1/hospital_dados`, {
    method: 'POST',
    headers: {
      'apikey': SB_KEY,
      'Authorization': `Bearer ${SB_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal',
    },
    body: JSON.stringify({ verif_ts: ts, dados: rows, snapshot_dates: [date] }),
  })

export async function POST(request) {
  try {
    const { rows, date, ts, chunkIndex, totalChunks } = await request.json()

    const res = await sbPost(rows, date, ts)

    if (!res.ok) {
      const txt = await res.text()
      return NextResponse.json({ ok: false, error: txt }, { status: 500 })
    }

    return NextResponse.json({ ok: true, saved: rows.length })
  } catch (e) {
    return NextResponse.json({ ok: false, error: e.message }, { status: 500 })
  }
}

export async function DELETE() {
  try {
    const res = await fetch(`${SB_URL}/rest/v1/hospital_dados?id=gt.0`, {
      method: 'DELETE',
      headers: {
        'apikey': SB_KEY,
        'Authorization': `Bearer ${SB_KEY}`,
        'Content-Type': 'application/json',
      },
    })
    return NextResponse.json({ ok: res.ok })
  } catch (e) {
    return NextResponse.json({ ok: false, error: e.message }, { status: 500 })
  }
}