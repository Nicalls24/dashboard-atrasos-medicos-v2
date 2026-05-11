'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

// ─── Palette ──────────────────────────────────────────────
const T = {
  bg:       '#080C14',
  surface:  '#0E1623',
  card:     '#111A2B',
  border:   '#1C2D45',
  accent:   '#00C6FF',
  accentB:  '#0072FF',
  danger:   '#FF4D6A',
  warning:  '#FFB340',
  success:  '#00E5A0',
  text:     '#EDF2FF',
  muted:    '#5A7799',
  sub:      '#8FADC8',
}

// ─── Helpers ──────────────────────────────────────────────
const parseHM = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24
  const s = String(v).trim()
  if (s.includes(':')) {
    const [h, m] = s.split(':').map(Number)
    return (h || 0) + (m || 0) / 60
  }
  const n = parseFloat(s.replace(',', '.'))
  return isNaN(n) ? 0 : n
}

const fmtH = (h) => {
  const abs = Math.abs(h)
  const hh = Math.floor(abs)
  const mm = Math.round((abs - hh) * 60)
  return `${hh}h${mm > 0 ? ` ${mm}m` : ''}`
}

const STATUS_COLOR = (s = '') => {
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO'))  return T.danger
  if (u.includes('GRAVE'))    return '#FF7A00'
  if (u.includes('ATRASO'))   return T.warning
  if (u.includes('FALTA'))    return '#B060FF'
  if (u.includes('REMARCA'))  return T.muted
  return T.success
}

const STATUS_LABEL = (s = '') => {
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO'))  return 'Crítico'
  if (u.includes('GRAVE'))    return 'Grave'
  if (u.includes('ATRASO'))   return 'Atraso'
  if (u.includes('FALTA'))    return 'Falta'
  if (u.includes('REMARCA'))  return 'Remarcação'
  return 'OK'
}

// ─── Parseia qualquer valor de data/hora para Date ────────
const parseDate = (v) => {
  if (!v) return null
  // número serial do Excel
  if (typeof v === 'number') {
    // serial Excel: dias desde 1900-01-01
    const d = new Date((v - 25569) * 86400 * 1000)
    return isNaN(d) ? null : d
  }
  const s = String(v).trim()
  if (!s) return null
  // dd/mm/yyyy ou dd/mm/yyyy hh:mm
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if (m1) return new Date(+m1[3], +m1[2] - 1, +m1[1])
  const d = new Date(s)
  return isNaN(d) ? null : d
}

// Retorna true se a data está dentro do período selecionado
const dentroDoPeríodo = (date, período) => {
  if (!date || período === 'TODOS') return true
  const agora = new Date()
  if (período === 'DIA') {
    return (
      date.getDate()     === agora.getDate() &&
      date.getMonth()    === agora.getMonth() &&
      date.getFullYear() === agora.getFullYear()
    )
  }
  if (período === 'SEMANA') {
    const inicioSemana = new Date(agora)
    inicioSemana.setDate(agora.getDate() - agora.getDay()) // domingo
    inicioSemana.setHours(0, 0, 0, 0)
    return date >= inicioSemana
  }
  if (período === 'MÊS') {
    return (
      date.getMonth()    === agora.getMonth() &&
      date.getFullYear() === agora.getFullYear()
    )
  }
  return true
}

// ─── Mini bar chart (SVG) ─────────────────────────────────
function MiniBar({ data, color }) {
  if (!data.length) return null
  const max = Math.max(...data.map(d => d.value), 1)
  const W = 100, H = 48, barW = W / data.length - 3
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: '100%', height: 52 }}>
      {data.map((d, i) => {
        const bh = Math.max(2, (d.value / max) * (H - 12))
        const x = i * (W / data.length) + 1
        return (
          <g key={i}>
            <rect x={x} y={H - bh} width={barW} height={bh} rx="2" fill={color} opacity="0.85" />
            <text x={x + barW / 2} y={H - bh - 3} textAnchor="middle" fontSize="5" fill={T.sub}>{d.label}</text>
          </g>
        )
      })}
    </svg>
  )
}

// ─── Horizontal bar ───────────────────────────────────────
function HBar({ label, value, max, color, unit = '', rank }) {
  const pct = max > 0 ? (value / max) * 100 : 0
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 5 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          {rank != null && (
            <span style={{ fontSize: 11, fontWeight: 700, color: rank < 3 ? T.danger : T.muted, minWidth: 20 }}>
              #{rank + 1}
            </span>
          )}
          <span style={{ fontSize: 13, color: T.text, fontWeight: 500,
            maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {label}
          </span>
        </div>
        <span style={{ fontSize: 13, fontWeight: 700, color }}>{value}{unit}</span>
      </div>
      <div style={{ background: T.border, borderRadius: 6, height: 6, overflow: 'hidden' }}>
        <div style={{ height: '100%', borderRadius: 6, background: color, width: `${pct}%`, transition: 'width .6s ease' }} />
      </div>
    </div>
  )
}

// ─── Stat card ────────────────────────────────────────────
function StatCard({ icon, label, value, sub, accent }) {
  return (
    <div style={{
      background: T.card, border: `1px solid ${T.border}`, borderRadius: 16,
      padding: '22px 24px', display: 'flex', flexDirection: 'column', gap: 10,
      position: 'relative', overflow: 'hidden',
    }}>
      <div style={{ position: 'absolute', top: -20, right: -20, width: 90, height: 90,
        borderRadius: '50%', background: accent, opacity: .07 }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{ fontSize: 22 }}>{icon}</span>
        <span style={{ fontSize: 12, color: T.muted, textTransform: 'uppercase',
          letterSpacing: '.08em', fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{ fontSize: 40, fontWeight: 800, color: T.text, lineHeight: 1, letterSpacing: '-1px' }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: T.sub }}>{sub}</div>}
    </div>
  )
}

// ─── Section header ───────────────────────────────────────
function SectionHeader({ children }) {
  return (
    <h2 style={{ fontSize: 14, fontWeight: 700, textTransform: 'uppercase',
      letterSpacing: '.12em', color: T.sub, marginBottom: 18, marginTop: 0 }}>
      {children}
    </h2>
  )
}

// ─── Card wrapper ─────────────────────────────────────────
function Card({ children, style = {} }) {
  return (
    <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 16, padding: 24, ...style }}>
      {children}
    </div>
  )
}

// ─── Donut (SVG) ─────────────────────────────────────────
function Donut({ segments, size = 120 }) {
  const r = 46, cx = 60, cy = 60, circ = 2 * Math.PI * r
  const total = segments.reduce((a, s) => a + s.value, 0) || 1
  let cumulative = 0
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={T.border} strokeWidth="14" />
      {segments.map((seg, i) => {
        const dash   = (seg.value / total) * circ
        const gap    = circ - dash
        const offset = circ - cumulative * circ / total
        cumulative  += seg.value
        return (
          <circle key={i} cx={cx} cy={cy} r={r} fill="none"
            stroke={seg.color} strokeWidth="14"
            strokeDasharray={`${dash} ${gap}`}
            strokeDashoffset={offset}
            strokeLinecap="round"
            transform="rotate(-90 60 60)"
          />
        )
      })}
      <text x="60" y="56" textAnchor="middle" fontSize="16" fontWeight="800" fill={T.text}>
        {total.toLocaleString('pt-BR')}
      </text>
      <text x="60" y="70" textAnchor="middle" fontSize="8" fill={T.muted}>registros</text>
    </svg>
  )
}

// ─── Badge ────────────────────────────────────────────────
function Badge({ label, color }) {
  return (
    <span style={{ fontSize: 11, fontWeight: 700, padding: '3px 9px',
      borderRadius: 20, background: color + '22', color, whiteSpace: 'nowrap' }}>
      {label}
    </span>
  )
}

// ─── Período badge (no topbar) ────────────────────────────
const PERIODOS = [
  { key: 'TODOS', label: 'Todos'  },
  { key: 'DIA',   label: 'Dia'    },
  { key: 'SEMANA',label: 'Semana' },
  { key: 'MÊS',   label: 'Mês'    },
]

function PeriodoSelector({ value, onChange }) {
  return (
    <div style={{ display: 'flex', gap: 4, background: T.card,
      border: `1px solid ${T.border}`, borderRadius: 10, padding: 4 }}>
      {PERIODOS.map(p => (
        <button key={p.key} onClick={() => onChange(p.key)} style={{
          padding: '6px 14px', borderRadius: 7, border: 'none',
          fontSize: 12, fontWeight: 700, cursor: 'pointer', transition: 'all .15s',
          background: value === p.key ? `linear-gradient(135deg,${T.accent},${T.accentB})` : 'transparent',
          color: value === p.key ? '#000' : T.muted,
        }}>{p.label}</button>
      ))}
    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
export default function Home() {
  const [dados,   setDados]   = useState([])
  const [uf,      setUf]      = useState('TODOS')
  const [status,  setStatus]  = useState('TODOS')
  const [loading, setLoading] = useState(false)
  const [search,  setSearch]  = useState('')

  // ── NOVO: filtro de período ──────────────────────────────
  const [período, setPeriodo] = useState('TODOS')

  const handleUpload = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setLoading(true)
    const buf = await file.arrayBuffer()
    const wb  = XLSX.read(buf, { type: 'buffer' })
    const ws  = wb.Sheets[wb.SheetNames[0]]
    const json = XLSX.utils.sheet_to_json(ws, { range: 3, defval: '' })
    setDados(json)
    setUf('TODOS')
    setStatus('TODOS')
    setSearch('')
    setPeriodo('TODOS')   // resetar período ao carregar nova planilha
    setLoading(false)
  }

  // ── Column detection (inalterado) ────────────────────────
  const cols = useMemo(() => {
    if (!dados.length) return {}
    const k = Object.keys(dados[0])
    const find = (...terms) => k.find(c => terms.some(t => c.toLowerCase().includes(t.toLowerCase()))) || ''
    return {
      unidade:  find('NM_LOCAL', 'UNIDADE'),
      medico:   find('NM_MEDICO', 'MEDICO'),
      esp:      find('ESPECIALIDADE'),
      status:   find('STATUS'),
      atraso:   find('TEMPO DE ATRASO', 'ATRASO'),
      espera:   find('TEMPO_DE_ESPERA'),
      qtPacts:  find('QT_PACIENTES_AGUARDANDO'),
      uf:       find('UF'),
      filial:   find('NM_FILIAL'),
      cidade:   find('CIDADE'),
      hrInicio: find('HR_INICIO'),
      hrEntrada:find('HR_ENTRADA'),
      // ── NOVO: detecta coluna de data para filtro de período
      data:     find('HR_INICIO', 'HR_ENTRADA', 'DT_', 'DATA', 'DATE'),
    }
  }, [dados])

  // ── Filtro de período: pré-filtra os dados pela coluna de data
  const dadosPorPeriodo = useMemo(() => {
    if (período === 'TODOS' || !cols.data) return dados
    return dados.filter(d => {
      const dt = parseDate(d[cols.data])
      return dentroDoPeríodo(dt, período)
    })
  }, [dados, cols.data, período])

  // ── Filtros de UF/status/search (inalterados, agora sobre dadosPorPeriodo)
  const ufs      = useMemo(() => [...new Set(dadosPorPeriodo.map(d => String(d[cols.uf] || '').trim()).filter(Boolean))].sort(), [dadosPorPeriodo, cols])
  const statuses = useMemo(() => [...new Set(dadosPorPeriodo.map(d => String(d[cols.status] || '').trim()).filter(Boolean))].sort(), [dadosPorPeriodo, cols])

  const filtered = useMemo(() => {
    let r = dadosPorPeriodo
    if (uf !== 'TODOS')     r = r.filter(d => String(d[cols.uf] || '').trim() === uf)
    if (status !== 'TODOS') r = r.filter(d => String(d[cols.status] || '').trim() === status)
    if (search)             r = r.filter(d =>
      [d[cols.unidade], d[cols.medico], d[cols.esp], d[cols.cidade]]
        .some(v => String(v || '').toLowerCase().includes(search.toLowerCase()))
    )
    return r
  }, [dadosPorPeriodo, uf, status, search, cols])

  // ── KPIs (inalterados) ────────────────────────────────────
  const totalRegistros = filtered.length
  const totalUnidades  = new Set(filtered.map(d => d[cols.unidade]).filter(Boolean)).size
  const totalMedicos   = new Set(filtered.map(d => d[cols.medico]).filter(Boolean)).size
  const emAtraso       = filtered.filter(d => String(d[cols.status] || '').toUpperCase().includes('ATRASO')).length
  const taxaAtraso     = totalRegistros > 0 ? ((emAtraso / totalRegistros) * 100).toFixed(1) : 0
  const totalEspera    = filtered.reduce((a, d) => a + parseHM(d[cols.espera]), 0)
  const mediaEspera    = totalRegistros > 0 ? totalEspera / totalRegistros : 0
  const totalPacAguard = filtered.reduce((a, d) => a + (Number(d[cols.qtPacts]) || 0), 0)

  // ── Breakdowns (inalterados) ──────────────────────────────
  const statusBreakdown = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s = String(d[cols.status] || 'OK').trim(); m[s] = (m[s] || 0) + 1 })
    return Object.entries(m).map(([label, value]) => ({ label, value, color: STATUS_COLOR(label) }))
      .sort((a, b) => b.value - a.value)
  }, [filtered, cols])

  const topUnidadesAtraso = useMemo(() => {
    const m = {}
    filtered.filter(d => String(d[cols.status] || '').toUpperCase().includes('ATRASO'))
      .forEach(d => { const u = d[cols.unidade] || 'Sem Unidade'; m[u] = (m[u] || 0) + 1 })
    return Object.entries(m).map(([nome, total]) => ({ nome, total }))
      .sort((a, b) => b.total - a.total).slice(0, 10)
  }, [filtered, cols])

  const topMedicos = useMemo(() => {
    const m = {}
    filtered.forEach(d => {
      const med = d[cols.medico] || 'Sem Médico'
      const h   = parseHM(d[cols.espera])
      if (!m[med]) m[med] = { med, unidade: d[cols.unidade] || '', total: 0, n: 0, maxH: 0 }
      m[med].total += h; m[med].n++
      if (h > m[med].maxH) m[med].maxH = h
    })
    return Object.values(m).filter(x => x.total > 0).sort((a, b) => b.total - a.total).slice(0, 10)
  }, [filtered, cols])

  const espBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const e = d[cols.esp] || 'Outro'; m[e] = (m[e] || 0) + 1 })
    return Object.entries(m).map(([nome, total]) => ({ nome, total }))
      .sort((a, b) => b.total - a.total).slice(0, 8)
  }, [filtered, cols])
  const maxEsp = espBreak[0]?.total || 1

  const ufBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const u = String(d[cols.uf] || '?').trim(); m[u] = (m[u] || 0) + 1 })
    return Object.entries(m).map(([nome, total]) => ({ nome, total }))
      .sort((a, b) => b.total - a.total).slice(0, 8)
  }, [filtered, cols])
  const maxUF = ufBreak[0]?.total || 1

  const impactoFila = useMemo(() => {
    const m = {}
    filtered.forEach(d => {
      const s = String(d[cols.status] || '').toUpperCase()
      if (!s.includes('ATRASO')) return
      const u = d[cols.unidade] || 'Sem Unidade'
      if (!m[u]) m[u] = { unidade: u, medicos: 0, pacientes: 0 }
      m[u].medicos++
      m[u].pacientes += Number(d[cols.qtPacts]) || 0
    })
    return Object.values(m).sort((a, b) => b.pacientes - a.pacientes).slice(0, 8)
  }, [filtered, cols])

  const hasData = dados.length > 0

  // ── Label do período selecionado para exibição ────────────
  const períodoLabel = PERIODOS.find(p => p.key === período)?.label || 'Todos'

  return (
    <div style={{ background: T.bg, minHeight: '100vh',
      fontFamily: "'DM Sans', 'Segoe UI', sans-serif", color: T.text, padding: '0' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; background: ${T.bg}; }
        ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 3px; }
        select option { background: ${T.surface}; color: ${T.text}; }
        .btn-upload:hover { opacity: .9; transform: translateY(-1px); }
        .row-table:hover { background: #162033 !important; }
        input::placeholder { color: ${T.muted}; }
      `}</style>

      {/* ── Topbar ── */}
      <div style={{
        background: T.surface, borderBottom: `1px solid ${T.border}`,
        padding: '0 40px', display: 'flex', alignItems: 'center',
        justifyContent: 'space-between', height: 64, position: 'sticky', top: 0, zIndex: 100,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8,
            background: `linear-gradient(135deg, ${T.accent}, ${T.accentB})`,
            display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16,
          }}>🏥</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: '-.01em' }}>Monitor Hospitalar</div>
            <div style={{ fontSize: 11, color: T.muted }}>Atrasos Médicos · Operacional</div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          {/* ── NOVO: seletor de período ── */}
          {hasData && <PeriodoSelector value={período} onChange={setPeriodo} />}

          {hasData && (
            <div style={{ fontSize: 12, color: T.muted }}>
              {totalRegistros.toLocaleString('pt-BR')} registros · {períodoLabel}
            </div>
          )}
          <label className="btn-upload" style={{
            background: `linear-gradient(135deg, ${T.accent}, ${T.accentB})`,
            color: '#000', fontWeight: 700, fontSize: 13,
            padding: '9px 20px', borderRadius: 10, cursor: 'pointer', transition: 'all .2s',
          }}>
            {loading ? 'Carregando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleUpload} />
          </label>
        </div>
      </div>

      <div style={{ padding: '32px 40px' }}>

        {/* ── Empty state (inalterado) ── */}
        {!hasData && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center',
            justifyContent: 'center', minHeight: 'calc(100vh - 140px)', gap: 16 }}>
            <div style={{ fontSize: 56 }}>📋</div>
            <div style={{ fontSize: 22, fontWeight: 700 }}>Nenhuma planilha carregada</div>
            <div style={{ color: T.muted, fontSize: 14 }}>
              Clique em "+ Carregar Planilha" para começar a análise
            </div>
          </div>
        )}

        {hasData && (<>

          {/* ── Filters (inalterados) ── */}
          <div style={{ display: 'flex', gap: 12, marginBottom: 28, flexWrap: 'wrap', alignItems: 'center' }}>
            <input value={search} onChange={e => setSearch(e.target.value)}
              placeholder="Buscar unidade, médico, especialidade…"
              style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10,
                color: T.text, fontSize: 13, padding: '9px 14px', outline: 'none', width: 300 }} />
            <select value={uf} onChange={e => setUf(e.target.value)} style={{
              background: T.card, border: `1px solid ${T.border}`, borderRadius: 10,
              color: T.text, fontSize: 13, padding: '9px 14px', outline: 'none', cursor: 'pointer' }}>
              <option value="TODOS">Todos os Estados</option>
              {ufs.map(u => <option key={u}>{u}</option>)}
            </select>
            <select value={status} onChange={e => setStatus(e.target.value)} style={{
              background: T.card, border: `1px solid ${T.border}`, borderRadius: 10,
              color: T.text, fontSize: 13, padding: '9px 14px', outline: 'none', cursor: 'pointer' }}>
              <option value="TODOS">Todos os Status</option>
              {statuses.map(s => <option key={s}>{s}</option>)}
            </select>
            {(uf !== 'TODOS' || status !== 'TODOS' || search) && (
              <button onClick={() => { setUf('TODOS'); setStatus('TODOS'); setSearch('') }}
                style={{ background: 'transparent', border: `1px solid ${T.border}`,
                  borderRadius: 10, color: T.muted, fontSize: 13, padding: '9px 14px', cursor: 'pointer' }}>
                ✕ Limpar filtros
              </button>
            )}
          </div>

          {/* ── KPI Cards (inalterados) ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 16, marginBottom: 24 }}>
            <StatCard icon="🩺" label="Total Registros" value={totalRegistros.toLocaleString('pt-BR')}
              sub={`${totalUnidades} unidades · ${totalMedicos} médicos`} accent={T.accent} />
            <StatCard icon="⚠️" label="Em Atraso" value={emAtraso.toLocaleString('pt-BR')}
              sub={`${taxaAtraso}% do total`} accent={T.danger} />
            <StatCard icon="⏱️" label="Espera Média" value={fmtH(mediaEspera)}
              sub="por atendimento" accent={T.warning} />
            <StatCard icon="👥" label="Pacientes Aguardando" value={totalPacAguard.toLocaleString('pt-BR')}
              sub="na fila agora" accent={T.success} />
          </div>

          {/* ── Row 2 (inalterada) ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr 1fr', gap: 16, marginBottom: 16 }}>
            <Card style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 16 }}>
              <SectionHeader>Distribuição</SectionHeader>
              <Donut size={130} segments={statusBreakdown.map(s => ({ value: s.value, color: s.color }))} />
              <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: 7 }}>
                {statusBreakdown.slice(0, 5).map(s => (
                  <div key={s.label} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                      <div style={{ width: 8, height: 8, borderRadius: '50%', background: s.color }} />
                      <span style={{ fontSize: 12, color: T.sub }}>{STATUS_LABEL(s.label)}</span>
                    </div>
                    <span style={{ fontSize: 12, fontWeight: 700, color: s.color }}>
                      {s.value.toLocaleString('pt-BR')}
                    </span>
                  </div>
                ))}
              </div>
            </Card>
            <Card>
              <SectionHeader>🔴 Unidades com Mais Atrasos</SectionHeader>
              {topUnidadesAtraso.slice(0, 6).map((u, i) => (
                <HBar key={u.nome} rank={i} label={u.nome} value={u.total}
                  max={topUnidadesAtraso[0]?.total || 1}
                  color={i === 0 ? T.danger : i === 1 ? '#FF7A00' : T.warning}
                  unit=" atrasos" />
              ))}
            </Card>
            <Card>
              <SectionHeader>📍 Registros por Estado (UF)</SectionHeader>
              {ufBreak.map((u, i) => (
                <HBar key={u.nome} label={u.nome} value={u.total} max={maxUF} color={T.accent} />
              ))}
            </Card>
          </div>

          {/* ── Row 3 (inalterada) ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
            <Card>
              <SectionHeader>👨‍⚕️ Médicos com Maior Tempo de Espera (acumulado)</SectionHeader>
              {topMedicos.map((m, i) => (
                <HBar key={m.med} rank={i} label={m.med}
                  value={parseFloat(m.total.toFixed(1))} max={topMedicos[0]?.total || 1}
                  color={i < 3 ? T.danger : T.warning} unit="h" />
              ))}
              {topMedicos.length === 0 && (
                <div style={{ color: T.muted, fontSize: 13 }}>Nenhum dado de espera encontrado.</div>
              )}
            </Card>
            <Card>
              <SectionHeader>🩻 Especialidades com Mais Atendimentos</SectionHeader>
              {espBreak.map((e, i) => (
                <HBar key={e.nome} label={e.nome} value={e.total} max={maxEsp}
                  color={`hsl(${200 + i * 18},70%,55%)`} />
              ))}
            </Card>
          </div>

          {/* ── Impacto na Fila (inalterado) ── */}
          <Card style={{ marginBottom: 16 }}>
            <SectionHeader>📊 Impacto na Fila de Espera — Unidades em Atraso</SectionHeader>
            {impactoFila.length === 0 ? (
              <div style={{ color: T.muted, fontSize: 13 }}>Nenhum atraso registrado com os filtros atuais.</div>
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      {['Unidade', 'Médicos em Atraso', 'Pacientes Aguardando', 'Pacientes / Médico', 'Severidade'].map(h => (
                        <th key={h} style={{ padding: '10px 14px', textAlign: 'left',
                          color: T.muted, fontWeight: 600, fontSize: 11,
                          textTransform: 'uppercase', letterSpacing: '.06em' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {impactoFila.map((r, i) => {
                      const ratio = r.medicos > 0 ? (r.pacientes / r.medicos).toFixed(1) : '-'
                      const sev   = r.medicos >= 5 ? 'Crítico' : r.medicos >= 3 ? 'Alto' : 'Moderado'
                      const sevC  = r.medicos >= 5 ? T.danger  : r.medicos >= 3 ? T.warning : T.success
                      return (
                        <tr key={r.unidade} className="row-table"
                          style={{ borderBottom: `1px solid ${T.border}`, transition: 'background .15s' }}>
                          <td style={{ padding: '12px 14px', fontWeight: 500, color: T.text }}>{r.unidade}</td>
                          <td style={{ padding: '12px 14px', color: T.danger,   fontWeight: 700 }}>{r.medicos}</td>
                          <td style={{ padding: '12px 14px', color: T.warning,  fontWeight: 700 }}>{r.pacientes}</td>
                          <td style={{ padding: '12px 14px', color: T.sub }}>{ratio}</td>
                          <td style={{ padding: '12px 14px' }}><Badge label={sev} color={sevC} /></td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </Card>

          {/* ── Footer ── */}
          <div style={{ textAlign: 'center', color: T.muted, fontSize: 11, paddingTop: 8, paddingBottom: 24 }}>
            Dashboard Monitoramento Hospitalar · Dados carregados da planilha · {new Date().toLocaleDateString('pt-BR')}
            {período !== 'TODOS' && ` · Filtro de período: ${períodoLabel}`}
          </div>

        </>)}
      </div>
    </div>
  )
}
