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

// ─── Helpers ──────────────────────────────────────────────
const parseHM = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24
  const s    = String(v).trim()
  const sign = s.startsWith('-') ? -1 : 1
  if (s.includes(':')) {
    const parts = s.replace('-','').split(':').map(Number)
    return sign * ((parts[0] || 0) + (parts[1] || 0) / 60)
  }
  const n = parseFloat(s.replace(',','.'))
  return isNaN(n) ? 0 : n
}

const fmtH = (h) => {
  const abs = Math.abs(h)
  const hh  = Math.floor(abs)
  const mm  = Math.round((abs - hh) * 60)
  return `${hh}h${mm > 0 ? ` ${mm}m` : ''}`
}

const STATUS_COLOR = (s = '') => {
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO')) return T.danger
  if (u.includes('GRAVE'))   return '#FF7A00'
  if (u.includes('ATRASO'))  return T.warning
  if (u.includes('FALTA'))   return '#B060FF'
  if (u.includes('REMARCA')) return T.muted
  return T.success
}
const STATUS_LABEL = (s = '') => {
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO')) return 'Crítico'
  if (u.includes('GRAVE'))   return 'Grave'
  if (u.includes('ATRASO'))  return 'Atraso'
  if (u.includes('FALTA'))   return 'Falta'
  if (u.includes('REMARCA')) return 'Remarcação'
  return 'OK'
}

// ─── Data: serial Excel → 'YYYY-MM-DD' ───────────────────
// Usa métodos UTC para evitar deslocamento de fuso horário (Brasil UTC-3)
// sem UTC: serial 46153 (11/05) aparece como 10/05 às 21h no horário local
const serialToDateStr = (v) => {
  if (!v && v !== 0) return ''
  let d
  if (typeof v === 'number') {
    // Serial Excel: dias desde 1899-12-30, convertido para ms UTC
    d = new Date(Math.round((v - 25569) * 86400 * 1000))
  } else {
    const s = String(v).trim()
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
    if (m) {
      // dd/mm/yyyy — cria como UTC explícito para evitar fuso
      d = new Date(Date.UTC(+m[3], +m[2]-1, +m[1]))
    } else {
      d = new Date(s)
    }
  }
  if (!d || isNaN(d)) return ''
  // *** USA getUTC* para ler a data correta independente do fuso ***
  const dd = String(d.getUTCDate()).padStart(2,'0')
  const mm = String(d.getUTCMonth()+1).padStart(2,'0')
  return `${d.getUTCFullYear()}-${mm}-${dd}`
}

// ─── Filtro por período — baseado nas datas da própria base ─
const buildFilter = (allDates, período) => {
  if (período === 'TODOS' || !allDates.length) return () => true
  const sorted  = [...allDates].sort()
  const maxDate = sorted[sorted.length - 1]   // mais recente

  if (período === 'DIA') {
    return (ds) => ds === maxDate
  }
  if (período === 'SEMANA') {
    const cutoff = new Date(maxDate + 'T00:00:00Z')
    cutoff.setUTCDate(cutoff.getUTCDate() - 6)
    const cutStr = `${cutoff.getUTCFullYear()}-${String(cutoff.getUTCMonth()+1).padStart(2,'0')}-${String(cutoff.getUTCDate()).padStart(2,'0')}`
    return (ds) => ds >= cutStr && ds <= maxDate
  }
  if (período === 'MÊS') {
    const mesAno = maxDate.slice(0,7)
    return (ds) => ds.slice(0,7) === mesAno
  }
  return () => true
}

// ─── Sub-components ───────────────────────────────────────
function HBar({ label, value, max, color, unit = '', rank, sub }) {
  const pct = max > 0 ? (value / max) * 100 : 0
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4, gap: 8 }}>
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8, minWidth: 0 }}>
          {rank != null && (
            <span style={{ fontSize: 11, fontWeight: 700, flexShrink: 0, marginTop: 1,
              color: rank < 3 ? T.danger : T.muted, minWidth: 22 }}>#{rank + 1}</span>
          )}
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 13, color: T.text, fontWeight: 500,
              overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</div>
            {sub && <div style={{ fontSize: 10, color: T.muted, marginTop: 1 }}>{sub}</div>}
          </div>
        </div>
        <span style={{ fontSize: 13, fontWeight: 700, color, flexShrink: 0 }}>{value}{unit}</span>
      </div>
      <div style={{ background: T.border, borderRadius: 6, height: 6, overflow: 'hidden' }}>
        <div style={{ height: '100%', borderRadius: 6, background: color,
          width: `${pct}%`, transition: 'width .6s ease' }} />
      </div>
    </div>
  )
}

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
      <div style={{ fontSize: 40, fontWeight: 800, color: T.text,
        lineHeight: 1, letterSpacing: '-1px' }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: T.sub }}>{sub}</div>}
    </div>
  )
}

function SectionHeader({ children }) {
  return (
    <h2 style={{ fontSize: 14, fontWeight: 700, textTransform: 'uppercase',
      letterSpacing: '.12em', color: T.sub, marginBottom: 18, marginTop: 0 }}>
      {children}
    </h2>
  )
}

function Card({ children, style = {} }) {
  return (
    <div style={{ background: T.card, border: `1px solid ${T.border}`,
      borderRadius: 16, padding: 24, ...style }}>{children}</div>
  )
}

function Donut({ segments, size = 120 }) {
  const r = 46, cx = 60, cy = 60, circ = 2 * Math.PI * r
  const total = segments.reduce((a, s) => a + s.value, 0) || 1
  let cum = 0
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={T.border} strokeWidth="14" />
      {segments.map((seg, i) => {
        const dash = (seg.value / total) * circ
        const off  = circ - cum * circ / total
        cum += seg.value
        return (
          <circle key={i} cx={cx} cy={cy} r={r} fill="none"
            stroke={seg.color} strokeWidth="14"
            strokeDasharray={`${dash} ${circ - dash}`}
            strokeDashoffset={off}
            strokeLinecap="round" transform="rotate(-90 60 60)" />
        )
      })}
      <text x="60" y="56" textAnchor="middle" fontSize="16" fontWeight="800" fill={T.text}>
        {total.toLocaleString('pt-BR')}
      </text>
      <text x="60" y="70" textAnchor="middle" fontSize="8" fill={T.muted}>registros</text>
    </svg>
  )
}

function Badge({ label, color }) {
  return (
    <span style={{ fontSize: 11, fontWeight: 700, padding: '3px 9px',
      borderRadius: 20, background: color + '22', color, whiteSpace: 'nowrap' }}>{label}</span>
  )
}

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
  const [dados,   setDados]   = useState([])
  const [uf,      setUf]      = useState('TODOS')
  const [status,  setStatus]  = useState('TODOS')
  const [loading, setLoading] = useState(false)
  const [search,  setSearch]  = useState('')
  const [período, setPeriodo] = useState('TODOS')

  const handleUpload = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setLoading(true)
    const buf  = await file.arrayBuffer()
    const wb   = XLSX.read(buf, { type: 'buffer' })
    const ws   = wb.Sheets[wb.SheetNames[0]]
    const json = XLSX.utils.sheet_to_json(ws, { range: 3, defval: '' })
    setDados(json)
    setUf('TODOS'); setStatus('TODOS'); setSearch(''); setPeriodo('TODOS')
    setLoading(false)
  }

  // ── Detecção de colunas ───────────────────────────────
  const cols = useMemo(() => {
    if (!dados.length) return {}
    const k    = Object.keys(dados[0])
    const find = (...terms) =>
      k.find(c => terms.some(t => c.toLowerCase() === t.toLowerCase())) ||
      k.find(c => terms.some(t => c.toLowerCase().includes(t.toLowerCase()))) || ''
    return {
      unidade: find('NM_LOCAL', 'UNIDADE'),
      medico:  find('NM_MEDICO', 'MEDICO'),
      esp:     find('DS_ESPECIALIDADE', 'ESPECIALIDADE'),
      status:  find('STATUS'),
      espera:  find('TEMPO_DE_ESPERA'),
      qtPacts: find('QT_PACIENTES_AGUARDANDO'),
      uf:      find('UF'),
      cidade:  find('CIDADE'),
      data:    find('DATA_AGENDA', 'DATA'),
    }
  }, [dados])

  // ── Adiciona _dateStr a cada linha ───────────────────
  const dadosComData = useMemo(() =>
    dados.map(d => ({ ...d, _dateStr: serialToDateStr(d[cols.data]) })),
    [dados, cols.data])

  // ── Todas as datas únicas presentes na base ──────────
  const allDates = useMemo(() =>
    [...new Set(dadosComData.map(d => d._dateStr).filter(Boolean))].sort(),
    [dadosComData])

  // ── Função de filtro para o período ─────────────────
  const periodoFn = useMemo(() =>
    buildFilter(allDates, período), [allDates, período])

  // ── Label da data acompanha o filtro — usa datas reais dos dados filtrados
  const períodoLabel = useMemo(() => {
    if (!allDates.length) return `${dados.length.toLocaleString('pt-BR')} reg.`
    const sorted  = [...allDates].sort()
    const maxDate = sorted[sorted.length - 1]   // data mais recente da base
    const minDate = sorted[0]                    // data mais antiga da base
    const fmt     = (s) => s.split('-').reverse().join('/')

    if (período === 'TODOS') {
      // Mostra o range completo de datas da base
      return minDate === maxDate
        ? `${fmt(maxDate)}`
        : `${fmt(minDate)} → ${fmt(maxDate)}`
    }
    if (período === 'DIA') {
      // Ontem = dia mais recente da base
      return `Ontem · ${fmt(maxDate)}`
    }
    if (período === 'SEMANA') {
      // 7 dias a partir do mais recente — usa as datas que realmente existem no filtro
      const c = new Date(maxDate + 'T00:00:00Z')
      c.setUTCDate(c.getUTCDate() - 6)
      const cutStr = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
      // Pega a data mais antiga que está dentro da semana
      const semanaMin = sorted.find(d => d >= cutStr) || cutStr
      return `${fmt(semanaMin)} → ${fmt(maxDate)}`
    }
    if (período === 'MÊS') {
      const mesAno  = maxDate.slice(0,7)
      const mesMin  = sorted.find(d => d.slice(0,7) === mesAno) || maxDate
      const [y, m]  = maxDate.split('-')
      const nome    = new Date(+y, +m-1).toLocaleString('pt-BR', { month: 'long' })
      return `${nome.charAt(0).toUpperCase()+nome.slice(1)} · ${fmt(mesMin)} → ${fmt(maxDate)}`
    }
    return ''
  }, [allDates, período, dados.length])

  // ── Filtra por período ───────────────────────────────
  const dadosPorPeriodo = useMemo(() =>
    dadosComData.filter(d => periodoFn(d._dateStr)),
    [dadosComData, periodoFn])

  // ── Filtros UF/status/busca ──────────────────────────
  const ufs = useMemo(() =>
    [...new Set(dadosPorPeriodo.map(d => String(d[cols.uf]||'').trim()).filter(Boolean))].sort(),
    [dadosPorPeriodo, cols])
  const statuses = useMemo(() =>
    [...new Set(dadosPorPeriodo.map(d => String(d[cols.status]||'').trim()).filter(Boolean))].sort(),
    [dadosPorPeriodo, cols])

  const filtered = useMemo(() => {
    let r = dadosPorPeriodo
    if (uf !== 'TODOS')     r = r.filter(d => String(d[cols.uf]||'').trim() === uf)
    if (status !== 'TODOS') r = r.filter(d => String(d[cols.status]||'').trim() === status)
    if (search)             r = r.filter(d =>
      [d[cols.unidade], d[cols.medico], d[cols.esp], d[cols.cidade]]
        .some(v => String(v||'').toLowerCase().includes(search.toLowerCase())))
    return r
  }, [dadosPorPeriodo, uf, status, search, cols])

  // ── KPIs ─────────────────────────────────────────────
  const totalRegistros = filtered.length
  const totalUnidades  = new Set(filtered.map(d => d[cols.unidade]).filter(Boolean)).size
  // Médicos únicos por nome
  const totalMedicos   = new Set(filtered.map(d => String(d[cols.medico]||'').trim()).filter(Boolean)).size
  const emAtraso       = filtered.filter(d => String(d[cols.status]||'').toUpperCase().includes('ATRASO')).length
  const taxaAtraso     = totalRegistros > 0 ? ((emAtraso/totalRegistros)*100).toFixed(1) : 0
  const totalEspera    = filtered.reduce((a,d) => a + parseHM(d[cols.espera]), 0)
  const mediaEspera    = totalRegistros > 0 ? totalEspera / totalRegistros : 0
  const totalPacAguard = filtered.reduce((a,d) => a + (Number(d[cols.qtPacts])||0), 0)

  // ── Status breakdown ──────────────────────────────────
  const statusBreakdown = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s = String(d[cols.status]||'OK').trim(); m[s] = (m[s]||0)+1 })
    return Object.entries(m).map(([label,value]) => ({ label, value, color: STATUS_COLOR(label) }))
      .sort((a,b) => b.value - a.value)
  }, [filtered, cols])

  // ── Top unidades por atraso ───────────────────────────
  const topUnidadesAtraso = useMemo(() => {
    const m = {}
    filtered.filter(d => String(d[cols.status]||'').toUpperCase().includes('ATRASO'))
      .forEach(d => { const u = d[cols.unidade]||'Sem Unidade'; m[u] = (m[u]||0)+1 })
    return Object.entries(m).map(([nome,total]) => ({ nome, total }))
      .sort((a,b) => b.total-a.total).slice(0,10)
  }, [filtered, cols])

  // ── Top médicos — CORRIGIDO ───────────────────────────
  // Agrupa por nome de médico. Cada combinação nome+dia+unidade = 1 agenda.
  const topMedicos = useMemo(() => {
    const m = {}
    filtered.forEach(d => {
      const med    = String(d[cols.medico]||'').trim()
      if (!med) return
      const espH   = Math.max(0, parseHM(d[cols.espera]))
      const dateS  = d._dateStr || ''
      const unid   = String(d[cols.unidade]||'').trim()
      const agKey  = `${dateS}||${unid}`   // 1 agenda = 1 dia + 1 unidade

      if (!m[med]) m[med] = { med, totalEspera: 0, agendasSet: new Set(), pacts: 0 }
      m[med].totalEspera += espH
      m[med].agendasSet.add(agKey)
      m[med].pacts += Number(d[cols.qtPacts])||0
    })
    return Object.values(m)
      .filter(x => x.totalEspera > 0)
      .map(x => ({
        med:     x.med,
        total:   parseFloat(x.totalEspera.toFixed(1)),
        agendas: x.agendasSet.size,
        pacts:   x.pacts,
      }))
      .sort((a,b) => b.total - a.total)
      .slice(0,10)
  }, [filtered, cols])

  // ── Especialidades ─────────────────────────────────────
  const espBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const e = d[cols.esp]||'Outro'; m[e] = (m[e]||0)+1 })
    return Object.entries(m).map(([nome,total]) => ({ nome,total }))
      .sort((a,b) => b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxEsp = espBreak[0]?.total || 1

  // ── UF breakdown ──────────────────────────────────────
  const ufBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const u = String(d[cols.uf]||'?').trim(); m[u] = (m[u]||0)+1 })
    return Object.entries(m).map(([nome,total]) => ({ nome,total }))
      .sort((a,b) => b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxUF = ufBreak[0]?.total || 1

  // ── Impacto fila — CORRIGIDO: médicos únicos por nome ─
  const impactoFila = useMemo(() => {
    const m = {}
    filtered.forEach(d => {
      const s = String(d[cols.status]||'').toUpperCase()
      if (!s.includes('ATRASO')) return
      const u   = d[cols.unidade]||'Sem Unidade'
      const med = String(d[cols.medico]||'').trim()
      if (!m[u]) m[u] = { unidade: u, medicosSet: new Set(), pacientes: 0 }
      if (med) m[u].medicosSet.add(med)
      m[u].pacientes += Number(d[cols.qtPacts])||0
    })
    return Object.values(m)
      .map(x => ({ unidade: x.unidade, medicos: x.medicosSet.size, pacientes: x.pacientes }))
      .sort((a,b) => b.pacientes - a.pacientes).slice(0,8)
  }, [filtered, cols])

  const hasData = dados.length > 0

  return (
    <div style={{ background: T.bg, minHeight: '100vh',
      fontFamily: "'DM Sans','Segoe UI',sans-serif", color: T.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:6px;background:${T.bg}}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
        select option{background:${T.surface};color:${T.text}}
        .btn-upload:hover{opacity:.9;transform:translateY(-1px)}
        .row-table:hover{background:#162033!important}
        input::placeholder{color:${T.muted}}
      `}</style>

      {/* ── Topbar ── */}
      <div style={{
        background: T.surface, borderBottom: `1px solid ${T.border}`,
        padding: '0 40px', display: 'flex', alignItems: 'center',
        justifyContent: 'space-between', height: 64, position: 'sticky', top: 0, zIndex: 100,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8, fontSize: 16,
            background: `linear-gradient(135deg,${T.accent},${T.accentB})`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>🏥</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>Monitor Hospitalar</div>
            <div style={{ fontSize: 11, color: T.muted }}>Atrasos Médicos · Operacional</div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          {hasData && (
            <PeriodoSelector
              value={período}
              onChange={setPeriodo}
              infoLabel={`${totalRegistros.toLocaleString('pt-BR')} reg. · ${períodoLabel}`}
            />
          )}
          <label className="btn-upload" style={{
            background: `linear-gradient(135deg,${T.accent},${T.accentB})`,
            color: '#000', fontWeight: 700, fontSize: 13,
            padding: '9px 20px', borderRadius: 10, cursor: 'pointer', transition: 'all .2s',
          }}>
            {loading ? 'Carregando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleUpload} />
          </label>
        </div>
      </div>

      <div style={{ padding: '32px 40px' }}>
        {!hasData && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center',
            justifyContent: 'center', minHeight: 'calc(100vh - 140px)', gap: 16 }}>
            <div style={{ fontSize: 56 }}>📋</div>
            <div style={{ fontSize: 22, fontWeight: 700 }}>Nenhuma planilha carregada</div>
            <div style={{ color: T.muted, fontSize: 14 }}>Clique em "+ Carregar Planilha" para começar</div>
          </div>
        )}

        {hasData && (<>
          {/* ── Filtros ── */}
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
                  borderRadius: 10, color: T.muted, fontSize: 13,
                  padding: '9px 14px', cursor: 'pointer' }}>✕ Limpar filtros</button>
            )}
          </div>

          {/* ── KPIs ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 16, marginBottom: 24 }}>
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

          {/* ── Row 2 ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr 1fr', gap: 16, marginBottom: 16 }}>
            <Card style={{ display: 'flex', flexDirection: 'column',
              alignItems: 'center', justifyContent: 'center', gap: 16 }}>
              <SectionHeader>Distribuição</SectionHeader>
              <Donut size={130} segments={statusBreakdown.map(s => ({ value: s.value, color: s.color }))} />
              <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: 7 }}>
                {statusBreakdown.slice(0,5).map(s => (
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
              {topUnidadesAtraso.slice(0,6).map((u,i) => (
                <HBar key={u.nome} rank={i} label={u.nome} value={u.total}
                  max={topUnidadesAtraso[0]?.total||1}
                  color={i===0 ? T.danger : i===1 ? '#FF7A00' : T.warning}
                  unit=" atrasos" />
              ))}
            </Card>

            <Card>
              <SectionHeader>📍 Registros por Estado (UF)</SectionHeader>
              {ufBreak.map(u => (
                <HBar key={u.nome} label={u.nome} value={u.total} max={maxUF} color={T.accent} />
              ))}
            </Card>
          </div>

          {/* ── Row 3: Médicos CORRIGIDO + Especialidades ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
            <Card>
              <SectionHeader>👨‍⚕️ Médicos com Maior Tempo de Espera (acumulado)</SectionHeader>
              {topMedicos.map((m,i) => (
                <HBar key={m.med} rank={i} label={m.med}
                  value={m.total} max={topMedicos[0]?.total||1}
                  color={i < 3 ? T.danger : T.warning} unit="h"
                  sub={`${m.agendas} ${m.agendas===1 ? 'agenda' : 'agendas'} · ${m.pacts.toLocaleString('pt-BR')} pacientes`}
                />
              ))}
              {topMedicos.length === 0 && (
                <div style={{ color: T.muted, fontSize: 13 }}>Nenhum dado de espera encontrado.</div>
              )}
            </Card>

            <Card>
              <SectionHeader>🩻 Especialidades com Mais Atendimentos</SectionHeader>
              {espBreak.map((e,i) => (
                <HBar key={e.nome} label={e.nome} value={e.total} max={maxEsp}
                  color={`hsl(${200 + i*18},70%,55%)`} />
              ))}
            </Card>
          </div>

          {/* ── Impacto na Fila — CORRIGIDO: médicos únicos ── */}
          <Card style={{ marginBottom: 16 }}>
            <SectionHeader>📊 Impacto na Fila de Espera — Unidades em Atraso</SectionHeader>
            {impactoFila.length === 0 ? (
              <div style={{ color: T.muted, fontSize: 13 }}>Nenhum atraso registrado com os filtros atuais.</div>
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      {['Unidade','Médicos Únicos em Atraso','Pacientes Aguardando','Pacientes / Médico','Severidade'].map(h => (
                        <th key={h} style={{ padding: '10px 14px', textAlign: 'left',
                          color: T.muted, fontWeight: 600, fontSize: 11,
                          textTransform: 'uppercase', letterSpacing: '.06em' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {impactoFila.map(r => {
                      const ratio = r.medicos > 0 ? (r.pacientes/r.medicos).toFixed(1) : '-'
                      const sev   = r.medicos >= 5 ? 'Crítico' : r.medicos >= 3 ? 'Alto' : 'Moderado'
                      const sevC  = r.medicos >= 5 ? T.danger  : r.medicos >= 3 ? T.warning : T.success
                      return (
                        <tr key={r.unidade} className="row-table"
                          style={{ borderBottom: `1px solid ${T.border}`, transition: 'background .15s' }}>
                          <td style={{ padding: '12px 14px', fontWeight: 500, color: T.text }}>{r.unidade}</td>
                          <td style={{ padding: '12px 14px', color: T.danger, fontWeight: 700 }}>{r.medicos}</td>
                          <td style={{ padding: '12px 14px', color: T.warning, fontWeight: 700 }}>{r.pacientes}</td>
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
            Dashboard Monitoramento Hospitalar · {new Date().toLocaleDateString('pt-BR')}
            {período !== 'TODOS' && ` · ${períodoLabel}`}
          </div>
        </>)}
      </div>
    </div>
  )
}
