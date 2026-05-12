'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

// ─── Palette ──────────────────────────────────────────────
const T = {
  bg:      '#080C14',
  surface: '#0E1623',
  card:    '#111A2B',
  border:  '#1C2D45',
  accent:  '#00C6FF',
  accentB: '#0072FF',
  danger:  '#FF4D6A',
  warning: '#FFB340',
  success: '#00E5A0',
  text:    '#EDF2FF',
  muted:   '#5A7799',
  sub:     '#8FADC8',
}

// ─── Status config (imagem 1) ────────────────────────────
const STATUS_CFG = [
  { key: 'OK',                bg: '#16A34A', fg: '#fff',    label: 'Motivo (OK)',            desc: 'Atendimento dentro do prazo'         },
  { key: 'ATRASO',            bg: '#CA8A04', fg: '#fff',    label: 'Atraso',                 desc: '31 min < atraso < 45 min'            },
  { key: 'ATRASO CRÍTICO',    bg: '#EA580C', fg: '#fff',    label: 'Atraso Crítico',         desc: '46 min > atraso < 1h30'              },
  { key: 'ATRASO GRAVE',      bg: '#DC2626', fg: '#fff',    label: 'Atraso Grave',           desc: 'Atraso > 1h30'                       },
  { key: 'Falta Médica',      bg: '#1D4ED8', fg: '#fff',    label: 'Médico Faltou',          desc: 'Ausência registrada'                 },
  { key: 'Remarcação Adm',    bg: '#7C3AED', fg: '#fff',    label: 'Remarcação Adm',         desc: 'Remarcação administrativa'           },
  { key: 'Remarcação Médico', bg: '#9333EA', fg: '#fff',    label: 'Remarcação Médico',      desc: 'Remarcação pelo médico'              },
  { key: 'SEM_PONTO',         bg: '#374151', fg: '#F9A825', label: 'Sem Ponto',              desc: 'HR_ENTRADA vazia — não bateu ponto'  },
]

const getCfg = (key) =>
  STATUS_CFG.find(s => s.key === key) ||
  { bg: T.border, fg: T.text, label: key, desc: '' }

// ─── Status color (para barras/badges) ───────────────────
const statusColor = (s = '') => {
  const u = s.toUpperCase()
  if (u === 'OK')                                           return '#16A34A'
  if (u === 'ATRASO' && !u.includes('CRÍTICO') && !u.includes('GRAVE')) return '#CA8A04'
  if (u.includes('CRÍTICO'))                               return '#EA580C'
  if (u.includes('GRAVE'))                                 return '#DC2626'
  if (u.includes('FALTA') || u.includes('MÉDICO FAL'))    return '#1D4ED8'
  if (u.includes('REMARCA'))                               return '#7C3AED'
  if (u === 'SEM_PONTO')                                   return '#F9A825'
  return T.muted
}

// ─── Helpers ─────────────────────────────────────────────
const parseHM = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24
  const s = String(v).trim()
  const sign = s.startsWith('-') ? -1 : 1
  if (s.includes(':')) {
    const p = s.replace('-','').split(':').map(Number)
    return sign * ((p[0]||0) + (p[1]||0)/60)
  }
  const n = parseFloat(s.replace(',','.'))
  return isNaN(n) ? 0 : n
}

const fmtH = (h) => {
  const abs = Math.abs(h)
  const hh = Math.floor(abs), mm = Math.round((abs-hh)*60)
  return `${hh}h${mm > 0 ? ` ${mm}m` : ''}`
}

// Serial Excel → 'YYYY-MM-DD' usando UTC para evitar fuso horário
const serialToDateStr = (v) => {
  if (!v && v !== 0) return ''
  let d
  if (typeof v === 'number') {
    d = new Date(Math.round((v - 25569) * 86400 * 1000))
  } else {
    const s = String(v).trim()
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
    if (m) d = new Date(Date.UTC(+m[3], +m[2]-1, +m[1]))
    else    d = new Date(s)
  }
  if (!d || isNaN(d)) return ''
  const dd = String(d.getUTCDate()).padStart(2,'0')
  const mm = String(d.getUTCMonth()+1).padStart(2,'0')
  return `${d.getUTCFullYear()}-${mm}-${dd}`
}

// Timedelta do Excel (fração de dia ou string hh:mm:ss) → horas inteiras
const tdToHour = (v) => {
  if (v === null || v === undefined || v === '') return -1
  if (typeof v === 'number') return Math.floor(v * 24)
  const s = String(v).trim()
  if (!s || s === 'NaT') return -1
  const m = s.match(/(\d+):(\d+)/)
  return m ? parseInt(m[1], 10) : -1
}

// Parseia timestamp da célula de verificação "dd/mm/yyyy hh:mm:ss"
const parseVerifTimestamp = (raw) => {
  if (!raw) return null
  const s = String(raw).trim()
  // Formato que vem da planilha: "05/12/2026 09:06:07" (dd/mm errado ou mm/dd?)
  // Conforme o usuário: "05/12/2026" está errado, correto é "12/05/2026"
  // Ou seja: a planilha grava como mm/dd/yyyy mas deve ser dd/mm/yyyy
  // Detecta: se mês > 12 -> interpretar como dd/mm, senão tentar ambos
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{2}:\d{2}:\d{2})$/)
  if (!m) return s
  const p1 = parseInt(m[1]), p2 = parseInt(m[2]), year = m[3], time = m[4]
  let day, month
  if (p1 > 12) {
    // p1 é dia com certeza
    day = p1; month = p2
  } else if (p2 > 12) {
    // p2 é dia com certeza → formato mm/dd
    day = p2; month = p1
  } else {
    // Ambíguos: assume que a planilha grava mm/dd (como visto no bug)
    day = p2; month = p1
  }
  return `${String(day).padStart(2,'0')}/${String(month).padStart(2,'0')}/${year} ${time}`
}

// buildFilter — usa datas da própria base
const buildFilter = (allDates, período) => {
  if (período === 'TODOS' || !allDates.length) return () => true
  const sorted = [...allDates].sort()
  const maxDate = sorted[sorted.length - 1]

  if (período === 'DIA') return (ds) => ds === maxDate

  if (período === 'SEMANA') {
    const c = new Date(maxDate + 'T00:00:00Z')
    c.setUTCDate(c.getUTCDate() - 6)
    const cutStr = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
    return (ds) => ds >= cutStr && ds <= maxDate
  }
  if (período === 'MÊS') {
    const mesAno = maxDate.slice(0,7)
    return (ds) => ds.slice(0,7) === mesAno
  }
  return () => true
}

// ─── Sub-components ───────────────────────────────────────
function Badge({ label, color, bg }) {
  return (
    <span style={{
      fontSize: 10.5, fontWeight: 700, padding: '3px 10px', borderRadius: 99,
      background: bg || color + '22', color, whiteSpace: 'nowrap', letterSpacing: '.04em',
    }}>{label}</span>
  )
}

function SH({ children }) {
  return (
    <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
      letterSpacing: '.12em', color: T.muted, marginBottom: 14 }}>{children}</div>
  )
}

function Card({ children, style = {} }) {
  return (
    <div style={{ background: T.card, border: `1px solid ${T.border}`,
      borderRadius: 16, padding: 22, ...style }}>{children}</div>
  )
}

function StatCard({ icon, label, value, sub, accent }) {
  return (
    <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 16,
      padding: '20px 22px', display: 'flex', flexDirection: 'column', gap: 8,
      position: 'relative', overflow: 'hidden' }}>
      <div style={{ position: 'absolute', top: -20, right: -20, width: 80, height: 80,
        borderRadius: '50%', background: accent, opacity: .07 }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{ fontSize: 20 }}>{icon}</span>
        <span style={{ fontSize: 11, color: T.muted, textTransform: 'uppercase',
          letterSpacing: '.08em', fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{ fontSize: 36, fontWeight: 800, color: T.text, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: T.sub }}>{sub}</div>}
    </div>
  )
}

function HBar({ label, value, max, color, unit='', rank, sub }) {
  const pct = max > 0 ? (value/max)*100 : 0
  return (
    <div style={{ marginBottom: 13 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4, gap: 8 }}>
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8, minWidth: 0 }}>
          {rank != null && (
            <span style={{ fontSize: 10, fontWeight: 700, flexShrink: 0, marginTop: 2,
              color: rank < 3 ? T.danger : T.muted, minWidth: 20 }}>#{rank+1}</span>
          )}
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 12.5, color: T.text, fontWeight: 500,
              overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</div>
            {sub && <div style={{ fontSize: 10, color: T.muted, marginTop: 1 }}>{sub}</div>}
          </div>
        </div>
        <span style={{ fontSize: 13, fontWeight: 700, color, flexShrink: 0 }}>{value}{unit}</span>
      </div>
      <div style={{ background: T.border, borderRadius: 99, height: 6, overflow: 'hidden' }}>
        <div style={{ height: '100%', borderRadius: 99, background: color,
          width: `${pct}%`, transition: 'width .6s' }} />
      </div>
    </div>
  )
}

// ─── Gráfico de Status (imagem 1) ────────────────────────
function StatusChart({ data, total }) {
  if (!data.length) return <div style={{ color: T.muted, fontSize: 13 }}>Sem dados</div>

  return (
    <div>
      {/* Barra empilhada visual */}
      <div style={{ display: 'flex', borderRadius: 10, overflow: 'hidden', height: 28, marginBottom: 20 }}>
        {data.map(s => {
          const pct = total > 0 ? (s.count / total) * 100 : 0
          if (pct < .3) return null
          const cfg = getCfg(s.key)
          return (
            <div key={s.key} title={`${cfg.label}: ${s.count} (${pct.toFixed(1)}%)`}
              style={{ width: `${pct}%`, background: cfg.bg, transition: 'width .6s',
                display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              {pct > 5 && (
                <span style={{ fontSize: 9, fontWeight: 800, color: cfg.fg, whiteSpace: 'nowrap' }}>
                  {pct.toFixed(0)}%
                </span>
              )}
            </div>
          )
        })}
      </div>

      {/* Cards de cada status */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 10 }}>
        {data.map(s => {
          const cfg  = getCfg(s.key)
          const pct  = total > 0 ? ((s.count / total) * 100).toFixed(1) : '0.0'
          return (
            <div key={s.key} style={{
              background: cfg.bg + '18',
              border: `1px solid ${cfg.bg}55`,
              borderRadius: 12, padding: '12px 14px',
              display: 'flex', flexDirection: 'column', gap: 6,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <div style={{ width: 10, height: 10, borderRadius: 3, background: cfg.bg, flexShrink: 0 }} />
                <span style={{ fontSize: 10, fontWeight: 700, color: cfg.fg || T.text,
                  textTransform: 'uppercase', letterSpacing: '.06em', lineHeight: 1.2 }}>
                  {cfg.label}
                </span>
              </div>
              <div style={{ fontSize: 26, fontWeight: 900, color: cfg.bg, lineHeight: 1 }}>
                {s.count.toLocaleString('pt-BR')}
              </div>
              <div style={{ fontSize: 10, color: T.muted }}>{pct}% do total</div>
              {cfg.desc && (
                <div style={{ fontSize: 9.5, color: T.sub, lineHeight: 1.4 }}>{cfg.desc}</div>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}

// ─── Período selector ────────────────────────────────────
const PERIODOS = [
  { key: 'TODOS',  label: 'Todos'  },
  { key: 'DIA',    label: 'Ontem'  },
  { key: 'SEMANA', label: 'Semana' },
  { key: 'MÊS',    label: 'Mês'    },
]

function PeriodoSelector({ value, onChange, infoLabel }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <div style={{ display: 'flex', gap: 4, background: T.card,
        border: `1px solid ${T.border}`, borderRadius: 10, padding: 4 }}>
        {PERIODOS.map(p => (
          <button key={p.key} onClick={() => onChange(p.key)} style={{
            padding: '6px 14px', borderRadius: 7, border: 'none',
            fontSize: 12, fontWeight: 700, cursor: 'pointer', transition: 'all .15s',
            background: value === p.key
              ? `linear-gradient(135deg,${T.accent},${T.accentB})` : 'transparent',
            color: value === p.key ? '#000' : T.muted,
          }}>{p.label}</button>
        ))}
      </div>
      {infoLabel && <span style={{ fontSize: 11, color: T.muted }}>{infoLabel}</span>}
    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
export default function Home() {
  const [dados,      setDados]      = useState([])
  const [uf,         setUf]         = useState('TODOS')
  const [statusFilt, setStatusFilt] = useState('TODOS')
  const [loading,    setLoading]    = useState(false)
  const [search,     setSearch]     = useState('')
  const [período,    setPeriodo]    = useState('TODOS')
  const [horaFilt,   setHoraFilt]   = useState('TODAS') // ← NOVO: filtro hora
  const [verifTs,    setVerifTs]    = useState('')       // timestamp verificação

  const handleUpload = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setLoading(true)
    const buf  = await file.arrayBuffer()
    const wb   = XLSX.read(buf, { type: 'buffer' })
    const ws   = wb.Sheets[wb.SheetNames[0]]
    // Lê timestamp da linha 2 (índice 1), coluna D (índice 3) — raw
    const rawAll = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true })
    const tsRaw  = rawAll[1]?.[3] || rawAll[1]?.[4] || ''
    setVerifTs(parseVerifTimestamp(String(tsRaw)))

    const json = XLSX.utils.sheet_to_json(ws, { range: 3, defval: '' })
    setDados(json)
    setUf('TODOS'); setStatusFilt('TODOS'); setSearch('')
    setPeriodo('TODOS'); setHoraFilt('TODAS')
    setLoading(false)
  }

  // ── Detecção de colunas
  const cols = useMemo(() => {
    if (!dados.length) return {}
    const k = Object.keys(dados[0])
    const find = (...terms) =>
      k.find(c => terms.some(t => c.toLowerCase() === t.toLowerCase())) ||
      k.find(c => terms.some(t => c.toLowerCase().includes(t.toLowerCase()))) || ''
    return {
      unidade:   find('NM_LOCAL', 'UNIDADE'),
      medico:    find('NM_MEDICO'),
      esp:       find('DS_ESPECIALIDADE', 'ESPECIALIDADE'),
      status:    find('STATUS'),
      espera:    find('TEMPO_DE_ESPERA'),
      qtPacts:   find('QT_PACIENTES_AGUARDANDO'),
      uf:        find('UF'),
      cidade:    find('CIDADE'),
      data:      find('DATA_AGENDA', 'DATA'),
      hrInicio:  find('HR_INICIO'),
      hrEntrada: find('HR_ENTRADA'),
    }
  }, [dados])

  // ── Adiciona _dateStr, _hora, _semPonto a cada linha
  const dadosRich = useMemo(() => dados.map(d => {
    const hrInicioRaw = d[cols.hrInicio]
    const hrEntRaw    = d[cols.hrEntrada]
    const semPonto    = (hrEntRaw === '' || hrEntRaw === null || hrEntRaw === undefined ||
                         String(hrEntRaw).trim() === '' || String(hrEntRaw).trim() === 'NaT')
    return {
      ...d,
      _dateStr:  serialToDateStr(d[cols.data]),
      _hora:     tdToHour(hrInicioRaw),
      _semPonto: semPonto,
      // Normaliza status: sem ponto → 'SEM_PONTO'
      _statusNorm: semPonto ? 'SEM_PONTO' : String(d[cols.status] || '').trim(),
    }
  }), [dados, cols])

  const allDates = useMemo(() =>
    [...new Set(dadosRich.map(d => d._dateStr).filter(Boolean))].sort(),
    [dadosRich])

  const periodoFn = useMemo(() => buildFilter(allDates, período), [allDates, período])

  // Horas disponíveis
  const horasDisp = useMemo(() => {
    const hs = [...new Set(dadosRich.filter(d => d._hora >= 0).map(d => d._hora))].sort((a,b) => a-b)
    return hs
  }, [dadosRich])

  // ── Label período
  const períodoLabel = useMemo(() => {
    if (!allDates.length) return ''
    const sorted  = [...allDates].sort()
    const maxDate = sorted[sorted.length - 1]
    const minDate = sorted[0]
    const fmt     = (s) => s.split('-').reverse().join('/')
    if (período === 'TODOS')  return `${fmt(minDate)} → ${fmt(maxDate)}`
    if (período === 'DIA')    return `Ontem · ${fmt(maxDate)}`
    if (período === 'SEMANA') {
      const c = new Date(maxDate + 'T00:00:00Z')
      c.setUTCDate(c.getUTCDate() - 6)
      const cs = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
      const sm = sorted.find(d => d >= cs) || cs
      return `${fmt(sm)} → ${fmt(maxDate)}`
    }
    if (período === 'MÊS') {
      const ma  = maxDate.slice(0,7)
      const sm  = sorted.find(d => d.slice(0,7) === ma) || maxDate
      const [y,m] = maxDate.split('-')
      const nome  = new Date(+y, +m-1).toLocaleString('pt-BR', { month: 'long' })
      return `${nome.charAt(0).toUpperCase()+nome.slice(1)} · ${fmt(sm)} → ${fmt(maxDate)}`
    }
    return ''
  }, [allDates, período])

  // ── Filtro encadeado: período → hora → UF → status → busca
  const filtered = useMemo(() => {
    let r = dadosRich.filter(d => periodoFn(d._dateStr))
    if (horaFilt !== 'TODAS')     r = r.filter(d => d._hora === Number(horaFilt))
    if (uf !== 'TODOS')           r = r.filter(d => String(d[cols.uf]||'').trim() === uf)
    if (statusFilt !== 'TODOS')   r = r.filter(d => d._statusNorm === statusFilt)
    if (search)                   r = r.filter(d =>
      [d[cols.unidade], d[cols.medico], d[cols.esp]]
        .some(v => String(v||'').toLowerCase().includes(search.toLowerCase())))
    return r
  }, [dadosRich, periodoFn, horaFilt, uf, statusFilt, search, cols])

  const ufs      = useMemo(() => [...new Set(filtered.map(d => String(d[cols.uf]||'').trim()).filter(Boolean))].sort(), [filtered, cols])
  const statuses = useMemo(() => [...new Set(dadosRich.filter(d => periodoFn(d._dateStr)).map(d => d._statusNorm).filter(Boolean))].sort(), [dadosRich, periodoFn])

  // ── KPIs
  const totalRegistros = filtered.length
  const totalUnidades  = new Set(filtered.map(d => d[cols.unidade]).filter(Boolean)).size
  const totalMedicos   = new Set(filtered.map(d => String(d[cols.medico]||'').trim()).filter(Boolean)).size
  const emAtraso       = filtered.filter(d => ['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm)).length
  const taxaAtraso     = totalRegistros > 0 ? ((emAtraso/totalRegistros)*100).toFixed(1) : 0
  const totalPacAguard = filtered.reduce((a,d) => a + (Number(d[cols.qtPacts])||0), 0)
  const totalEspera    = filtered.reduce((a,d) => a + parseHM(d[cols.espera]), 0)
  const mediaEspera    = totalRegistros > 0 ? totalEspera/totalRegistros : 0

  // ── Status breakdown para o gráfico (inclui SEM_PONTO)
  const statusBreakdown = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s = d._statusNorm||'OK'; m[s] = (m[s]||0)+1 })
    return Object.entries(m)
      .map(([key, count]) => ({ key, count }))
      .sort((a,b) => {
        const order = ['OK','ATRASO','ATRASO CRÍTICO','ATRASO GRAVE','Falta Médica','Remarcação Adm','Remarcação Médico','SEM_PONTO']
        return (order.indexOf(a.key)||99) - (order.indexOf(b.key)||99)
      })
  }, [filtered])

  // ── Médicos em atraso (para tabela)
  const medicosAtraso = useMemo(() => {
    const m = {}
    filtered
      .filter(d => ['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm))
      .forEach(d => {
        const nome   = String(d[cols.medico]||'').trim()
        const unid   = String(d[cols.unidade]||'').trim()
        const status = d._statusNorm
        if (!nome) return
        const k = `${nome}||${unid}`
        if (!m[k]) m[k] = { nome, unid, status, pacts: 0, count: 0 }
        m[k].pacts += Number(d[cols.qtPacts])||0
        m[k].count++
        // prioriza status mais grave
        const order = ['ATRASO GRAVE','ATRASO CRÍTICO','ATRASO']
        if (order.indexOf(status) < order.indexOf(m[k].status)) m[k].status = status
      })
    return Object.values(m).sort((a,b) => {
      const order = ['ATRASO GRAVE','ATRASO CRÍTICO','ATRASO']
      return order.indexOf(a.status) - order.indexOf(b.status) || b.pacts - a.pacts
    })
  }, [filtered, cols])

  // ── Médicos sem ponto
  const medicosSemPonto = useMemo(() => {
    const m = {}
    filtered
      .filter(d => d._semPonto)
      .forEach(d => {
        const nome = String(d[cols.medico]||'').trim()
        const unid = String(d[cols.unidade]||'').trim()
        if (!nome) return
        const k = `${nome}||${unid}`
        if (!m[k]) m[k] = { nome, unid, pacts: 0, count: 0, status: d._statusNorm }
        m[k].pacts += Number(d[cols.qtPacts])||0
        m[k].count++
      })
    return Object.values(m).sort((a,b) => b.pacts - a.pacts)
  }, [filtered, cols])

  // ── Top unidades atraso
  const topUnidades = useMemo(() => {
    const m = {}
    filtered.filter(d => ['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm))
      .forEach(d => { const u = d[cols.unidade]||'Sem Unidade'; m[u] = (m[u]||0)+1 })
    return Object.entries(m).map(([nome,total]) => ({ nome, total }))
      .sort((a,b) => b.total-a.total).slice(0,8)
  }, [filtered, cols])

  const hasData = dados.length > 0

  return (
    <div style={{ background: T.bg, minHeight: '100vh',
      fontFamily: "'DM Sans','Segoe UI',sans-serif", color: T.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;background:${T.bg}}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
        select option{background:${T.surface};color:${T.text}}
        .ubtn:hover{opacity:.9;transform:translateY(-1px)}
        .trow:hover td{background:#0e1b2c!important}
        input::placeholder{color:${T.muted}}
      `}</style>

      {/* ── Topbar ── */}
      <div style={{
        background: T.surface, borderBottom: `1px solid ${T.border}`,
        padding: '0 36px', height: 62, display: 'flex', alignItems: 'center',
        justifyContent: 'space-between', position: 'sticky', top: 0, zIndex: 100,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8, fontSize: 16,
            background: `linear-gradient(135deg,${T.accent},${T.accentB})`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>🏥</div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700 }}>Monitor Hospitalar</div>
            <div style={{ fontSize: 11, color: T.muted }}>
              Atrasos Médicos · Operacional
              {verifTs && <span style={{ marginLeft: 8, color: T.accent }}>· Atualizado: {verifTs}</span>}
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          {hasData && (
            <PeriodoSelector
              value={período} onChange={setPeriodo}
              infoLabel={`${totalRegistros.toLocaleString('pt-BR')} reg. · ${períodoLabel}`}
            />
          )}
          <label className="ubtn" style={{
            background: `linear-gradient(135deg,${T.accent},${T.accentB})`,
            color: '#000', fontWeight: 700, fontSize: 13,
            padding: '8px 18px', borderRadius: 9, cursor: 'pointer', transition: 'all .2s',
          }}>
            {loading ? 'Carregando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleUpload} />
          </label>
        </div>
      </div>

      <div style={{ padding: '24px 36px' }}>
        {!hasData && (
          <div style={{ minHeight: 'calc(100vh - 130px)', display: 'flex',
            flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 14 }}>
            <div style={{ fontSize: 52 }}>📋</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>Nenhuma planilha carregada</div>
            <div style={{ color: T.muted, fontSize: 13 }}>Clique em "+ Carregar Planilha" para começar</div>
          </div>
        )}

        {hasData && (<>

          {/* ── Filtros ── */}
          <div style={{ display: 'flex', gap: 10, marginBottom: 22, flexWrap: 'wrap', alignItems: 'center' }}>
            <input value={search} onChange={e => setSearch(e.target.value)}
              placeholder="Buscar unidade, médico, especialidade…"
              style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 9,
                color: T.text, fontSize: 13, padding: '8px 13px', outline: 'none', width: 280 }} />

            {/* Filtro HORA */}
            <select value={horaFilt} onChange={e => setHoraFilt(e.target.value)} style={{
              background: T.card, border: `1px solid ${T.accent}66`, borderRadius: 9,
              color: T.accent, fontSize: 13, fontWeight: 700, padding: '8px 13px',
              outline: 'none', cursor: 'pointer' }}>
              <option value="TODAS">⏰ Todas as horas</option>
              {horasDisp.map(h => (
                <option key={h} value={h}>{String(h).padStart(2,'0')}h</option>
              ))}
            </select>

            <select value={uf} onChange={e => setUf(e.target.value)} style={{
              background: T.card, border: `1px solid ${T.border}`, borderRadius: 9,
              color: T.text, fontSize: 13, padding: '8px 13px', outline: 'none', cursor: 'pointer' }}>
              <option value="TODOS">Todos os Estados</option>
              {ufs.map(u => <option key={u}>{u}</option>)}
            </select>

            <select value={statusFilt} onChange={e => setStatusFilt(e.target.value)} style={{
              background: T.card, border: `1px solid ${T.border}`, borderRadius: 9,
              color: T.text, fontSize: 13, padding: '8px 13px', outline: 'none', cursor: 'pointer' }}>
              <option value="TODOS">Todos os Status</option>
              {statuses.map(s => {
                const cfg = getCfg(s)
                return <option key={s} value={s}>{cfg.label}</option>
              })}
            </select>

            {(uf !== 'TODOS' || statusFilt !== 'TODOS' || search || horaFilt !== 'TODAS') && (
              <button onClick={() => { setUf('TODOS'); setStatusFilt('TODOS'); setSearch(''); setHoraFilt('TODAS') }}
                style={{ background: 'transparent', border: `1px solid ${T.border}`, borderRadius: 9,
                  color: T.muted, fontSize: 13, padding: '8px 13px', cursor: 'pointer' }}>
                ✕ Limpar
              </button>
            )}
          </div>

          {/* ── KPIs ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 14, marginBottom: 20 }}>
            <StatCard icon="🩺" label="Total Registros"
              value={totalRegistros.toLocaleString('pt-BR')}
              sub={`${totalUnidades} unidades · ${totalMedicos} médicos únicos`} accent={T.accent} />
            <StatCard icon="⚠️" label="Em Atraso"
              value={emAtraso.toLocaleString('pt-BR')}
              sub={`${taxaAtraso}% do total`} accent={T.danger} />
            <StatCard icon="⏱️" label="Espera Média"
              value={fmtH(mediaEspera)} sub="por atendimento" accent={T.warning} />
            <StatCard icon="👥" label="Pacientes Aguardando"
              value={totalPacAguard.toLocaleString('pt-BR')}
              sub="na fila agora" accent={T.success} />
          </div>

          {/* ── GRÁFICO DE STATUS ── */}
          <Card style={{ marginBottom: 18 }}>
            <SH>📊 Distribuição de Status — {horaFilt !== 'TODAS' ? `${String(horaFilt).padStart(2,'0')}h` : períodoLabel || 'todos os registros'}</SH>
            <StatusChart data={statusBreakdown} total={totalRegistros} />
          </Card>

          {/* ── Médicos em Atraso + Sem Ponto ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>

            {/* Médicos em Atraso */}
            <Card>
              <SH>🚨 Médicos em Atraso ({medicosAtraso.length})</SH>
              <div style={{ overflowY: 'auto', maxHeight: 380 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead style={{ position: 'sticky', top: 0, background: T.card }}>
                    <tr>
                      {['#','Médico','Unidade','Status','Pac.'].map(h => (
                        <th key={h} style={{ padding: '8px 10px', textAlign: 'left',
                          borderBottom: `1px solid ${T.border}`, color: T.muted,
                          fontWeight: 600, fontSize: 10, textTransform: 'uppercase',
                          letterSpacing: '.07em', whiteSpace: 'nowrap' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {medicosAtraso.slice(0, 30).map((m, i) => {
                      const cfg = getCfg(m.status)
                      return (
                        <tr key={`${m.nome}${i}`} className="trow">
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.muted, fontSize: 10 }}>{i+1}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            fontWeight: 600, color: T.text, whiteSpace: 'nowrap',
                            maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis' }}>{m.nome}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.sub, fontSize: 11, maxWidth: 140,
                            overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.unid}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}` }}>
                            <Badge label={cfg.label} color={cfg.bg} bg={cfg.bg+'22'} />
                          </td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.warning, fontWeight: 700, textAlign: 'center' }}>{m.pacts}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                {medicosAtraso.length === 0 && (
                  <div style={{ color: T.muted, fontSize: 13, padding: 12 }}>Nenhum médico em atraso.</div>
                )}
              </div>
            </Card>

            {/* Médicos Sem Ponto */}
            <Card>
              <SH>🔴 Médicos Sem Ponto ({medicosSemPonto.length}) — HR_ENTRADA vazia</SH>
              <div style={{ overflowY: 'auto', maxHeight: 380 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead style={{ position: 'sticky', top: 0, background: T.card }}>
                    <tr>
                      {['#','Médico','Unidade','Status','Pac. Impacto'].map(h => (
                        <th key={h} style={{ padding: '8px 10px', textAlign: 'left',
                          borderBottom: `1px solid ${T.border}`, color: T.muted,
                          fontWeight: 600, fontSize: 10, textTransform: 'uppercase',
                          letterSpacing: '.07em', whiteSpace: 'nowrap' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {medicosSemPonto.slice(0, 30).map((m, i) => {
                      const cfg = getCfg(m.status)
                      return (
                        <tr key={`${m.nome}${i}`} className="trow">
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.muted, fontSize: 10 }}>{i+1}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            fontWeight: 600, color: T.text, whiteSpace: 'nowrap',
                            maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis' }}>{m.nome}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.sub, fontSize: 11, maxWidth: 140,
                            overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.unid}</td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}` }}>
                            <Badge label={cfg.label} color={cfg.bg} bg={cfg.bg+'22'} />
                          </td>
                          <td style={{ padding: '8px 10px', borderBottom: `1px solid ${T.border}`,
                            color: T.danger, fontWeight: 700, textAlign: 'center' }}>{m.pacts}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                {medicosSemPonto.length === 0 && (
                  <div style={{ color: T.muted, fontSize: 13, padding: 12 }}>Nenhum médico sem ponto.</div>
                )}
              </div>
            </Card>
          </div>

          {/* ── Top unidades ── */}
          <Card style={{ marginBottom: 16 }}>
            <SH>🏥 Unidades com Mais Atrasos</SH>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 48px' }}>
              {topUnidades.map((u, i) => (
                <HBar key={u.nome} rank={i} label={u.nome} value={u.total}
                  max={topUnidades[0]?.total||1}
                  color={i===0 ? T.danger : i===1 ? '#FF7A00' : T.warning}
                  unit=" atrasos" />
              ))}
            </div>
          </Card>

          {/* ── Footer ── */}
          <div style={{ textAlign: 'center', color: T.muted, fontSize: 11, paddingBottom: 20 }}>
            Monitor Hospitalar · {new Date().toLocaleDateString('pt-BR')}
            {período !== 'TODOS' && ` · ${períodoLabel}`}
            {horaFilt !== 'TODAS' && ` · Hora ${String(horaFilt).padStart(2,'0')}h`}
            {verifTs && ` · Base: ${verifTs}`}
          </div>
        </>)}
      </div>
    </div>
  )
}
