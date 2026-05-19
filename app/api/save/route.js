import { NextResponse } from 'next/server'

const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'

const SB_HEADERS = {
  'apikey': SB_KEY,
  'Authorization': `Bearer ${SB_KEY}`,
  'Content-Type': 'application/json',
  'Prefer': 'return=minimal',
}

function serialToDate(v) {
  if (!v && v !== 0) return null
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v
  const n = Number(v)
  if (isNaN(n) || n < 1) return null
  const offsetDays = n > 60 ? 25568 : 25569
  const d = new Date(Math.round((n - offsetDays) * 86400 * 1000) + 43200000)
  if (isNaN(d.getTime())) return null
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
}

// Excel timedelta: número = fração do dia (ex: 0.020833 = 30min = 30/1440)
// String pode vir como "HH:MM:SS" ou "H:MM:SS"
function toMinutes(v) {
  if (v === null || v === undefined || v === '') return null

  if (typeof v === 'number') {
    // Fração de dia → minutos (1 dia = 1440 minutos)
    if (v >= 0 && v < 1) return Math.round(v * 1440)
    // Total de segundos (inteiro > 1)
    if (v >= 1) return Math.round(v / 60)
    return 0
  }

  if (typeof v === 'string') {
    const s     = v.trim()
    const neg   = s.startsWith('-')
    const clean = s.replace('-','').trim()
    const parts = clean.split(':').map(Number)
    const mins  = (parts[0]||0)*60 + (parts[1]||0)
    return neg ? -mins : mins
  }

  return null
}

function mapAgenda(r, ts) {
  return {
    data_agenda:         serialToDate(r['DATA_AGENDA']),
    dt_registro:         serialToDate(r['DT_REGISTRO'] || r['dt_registro']),
    uf:                  r['UF']                  || null,
    nm_filial:           r['NM_FILIAL']            || null,
    nm_local:            r['NM_LOCAL']             || null,
    nm_medico:           r['NM_MEDICO']            || null,
    ds_especialidade:    r['DS_ESPECIALIDADE']     || null,
    cidade:              r['CIDADE          ']     || r['CIDADE'] || null,
    base_sigo:           r['BASE_DE_DADOS_SIGO']   || null,
    hr_inicio_min:       toMinutes(r['HR_INICIO']),
    hr_fim_min:          toMinutes(r['HR_FIM']),
    hr_entrada_min:      toMinutes(r['HR_ENTRADA']),
    status:              r['STATUS']               || null,
    atraso:              r['ATRASO']               || null,
    tempo_atraso:        r['TEMPO DE ATRASO']      || null,
    motivo_cancelamento: r['MOTIVO_CANCELAMENTO']  || null,
    entrada_origem:      r['ENTRADA_ORIGEM']        || null,
    qt_consulta:         Number(r['QT_CONSULTA'])  || 0,
    qt_encaixe:          Number(r['QT_ENCAIXE'])   || 0,
    verif_ts:            ts || null,
  }
}

function mapEspera(r, ts) {
  const tempoEspera = r['TEMPO_DE_ESPERA']
  const qtPac       = r[' QT_PACIENTES_AGUARDANDO'] ?? r['QT_PACIENTES_AGUARDANDO']

  // skip rows without any espera data
  if ((tempoEspera === null || tempoEspera === undefined || tempoEspera === '') &&
      (qtPac === null || qtPac === undefined || qtPac === '')) return null

  // tempo de atraso (coluna T) — pode ser negativo
  const tempoAtraso = r['TEMPO DE ATRASO']
  let tempoAtrasoMin = null
  if (tempoAtraso !== null && tempoAtraso !== undefined && tempoAtraso !== '') {
    tempoAtrasoMin = toMinutes(tempoAtraso)
  }

  return {
    data_agenda:             serialToDate(r['DATA_AGENDA']),
    dt_registro:             serialToDate(r['DT_REGISTRO'] || r['dt_registro']),
    uf:                      r['UF']              || null,
    nm_local:                r['NM_LOCAL']        || null,
    nm_medico:               r['NM_MEDICO']       || null,
    ds_especialidade:        r['DS_ESPECIALIDADE'] || null,
    cidade:                  r['CIDADE          '] || r['CIDADE'] || null,
    hr_registro_espera_min:  toMinutes(r['HR_REGISTRO_ESPERA']),
    qt_pacientes_aguardando: qtPac !== null && qtPac !== undefined && qtPac !== '' ? Number(qtPac) : null,
    tempo_espera_min:        toMinutes(tempoEspera),
    qt_pacts:                r['QTS PACTS'] !== null && r['QTS PACTS'] !== undefined ? Number(r['QTS PACTS']) : null,
    atraso:                  r['ATRASO']           || null,
    tempo_atraso_min:        tempoAtrasoMin,
    status:                  r['STATUS']           || null,
    verif_ts:                ts || null,
  }
}

export async function POST(request) {
  try {
    const { rows, ts, table } = await request.json()

    if (!['agendas','espera'].includes(table))
      return NextResponse.json({ ok:false, error:'Invalid table' }, { status:400 })

    const mapped = table === 'agendas'
      ? rows.map(r => mapAgenda(r, ts)).filter(Boolean)
      : rows.map(r => mapEspera(r, ts)).filter(Boolean)

    if (!mapped.length) return NextResponse.json({ ok:true, saved:0 })

    const res = await fetch(`${SB_URL}/rest/v1/${table}`, {
      method: 'POST',
      headers: SB_HEADERS,
      body: JSON.stringify(mapped),
    })

    if (!res.ok) {
      const txt = await res.text()
      return NextResponse.json({ ok:false, error:txt }, { status:500 })
    }

    return NextResponse.json({ ok:true, saved: mapped.length })
  } catch(e) {
    return NextResponse.json({ ok:false, error:e.message }, { status:500 })
  }
}

export async function DELETE() {
  try {
    await Promise.all([
      fetch(`${SB_URL}/rest/v1/agendas?id=gt.0`, { method:'DELETE', headers:SB_HEADERS }),
      fetch(`${SB_URL}/rest/v1/espera?id=gt.0`,  { method:'DELETE', headers:SB_HEADERS }),
    ])
    return NextResponse.json({ ok:true })
  } catch(e) {
    return NextResponse.json({ ok:false, error:e.message }, { status:500 })
  }
}