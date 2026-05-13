'use client'

import { useMemo, useState, useEffect, useCallback } from 'react'
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
  ok:      '#00E5A0',
  text:    '#EDF2FF',
  muted:   '#5A7799',
  sub:     '#8FADC8',
  purple:  '#B060FF',
  orange:  '#FF7A00',
}

// ─── Status config (igual à planilha de referência) ───────
const STATUS_CFG = {
  'OK':                { label: 'OK (Motivo)',          color: '#22C55E', bg: '#14532d' },
  'ATRASO':            { label: 'Atraso 31–45min',      color: '#F59E0B', bg: '#78350f' },
  'ATRASO CRÍTICO':    { label: 'Atraso Crítico 46min–1h30', color: '#EF4444', bg: '#7f1d1d' },
  'ATRASO GRAVE':      { label: 'Atraso Grave >1h30',   color: '#7C3AED', bg: '#3b0764' },
  'Falta Médica':      { label: 'Médico Faltou',        color: '#1D4ED8', bg: '#1e3a8a' },
  'ERRO DE ABERTURA':  { label: 'Erro de Abertura',     color: '#0891B2', bg: '#164e63' },
  'SEM PONTO':         { label: 'Sem Ponto',            color: '#6B7280', bg: '#1f2937' },
  'CONTINUA/PONTO EM ABERTO': { label: 'Continua/Ponto Aberto', color: '#D97706', bg: '#451a03' },
  'NÃO COBRAR':        { label: 'Não Cobrar',           color: '#9CA3AF', bg: '#374151' },
  'Remarcação Adm':    { label: 'Remarcação Adm',       color: '#64748B', bg: '#1e293b' },
  'Remarcação Médico': { label: 'Remarcação Médico',    color: '#64748B', bg: '#1e293b' },
  'Remarcação médico': { label: 'Remarcação Médico',    color: '#64748B', bg: '#1e293b' },
}
const getStatusCfg = (s = '') => {
  if (STATUS_CFG[s]) return STATUS_CFG[s]
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO')) return STATUS_CFG['ATRASO CRÍTICO']
  if (u.includes('GRAVE'))   return STATUS_CFG['ATRASO GRAVE']
  if (u.includes('ATRASO'))  return STATUS_CFG['ATRASO']
  if (u.includes('FALTA'))   return STATUS_CFG['Falta Médica']
  if (u.includes('REMARCA')) return STATUS_CFG['Remarcação Adm']
  return { label: s || 'OK', color: T.success, bg: '#0d2a1f' }
}

// ─── Helpers ──────────────────────────────────────────────
const tdToH   = (v) => { try { return v.total_seconds ? v.total_seconds()/3600 : (typeof v==='number' ? v*24 : 0) } catch { return 0 } }
const parseHM = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24
  const s = String(v).trim(), sign = s.startsWith('-') ? -1 : 1
  if (s.includes(':')) {
    const p = s.replace('-','').split(':').map(Number)
    return sign * ((p[0]||0) + (p[1]||0)/60)
  }
  const n = parseFloat(s.replace(',','.'))
  return isNaN(n) ? 0 : n
}
const fmtH = (h) => {
  const abs=Math.abs(h), hh=Math.floor(abs), mm=Math.round((abs-hh)*60)
  return `${hh}h${mm>0?` ${mm}m`:''}`
}
const fmtMin = (m) => {
  const mm=Math.round(Math.abs(m))
  if (mm < 60) return `${mm}min`
  return `${Math.floor(mm/60)}h${mm%60>0?` ${mm%60}m`:''}`
}

// ─── Serial Excel → 'YYYY-MM-DD' com UTC para evitar fuso ─
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
  const dd=String(d.getUTCDate()).padStart(2,'0')
  const mm=String(d.getUTCMonth()+1).padStart(2,'0')
  return `${d.getUTCFullYear()}-${mm}-${dd}`
}

// ─── DT_REGISTRO → timestamp legível (formato dd/mm/yyyy) ─
// A planilha guarda como serial Excel (número), não como string
const parseDtRegistro = (v, rows) => {
  // Tenta pegar do campo DT_REGISTRO + HR_REGISTRO_ESPERA ou usa DATA_AGENDA
  if (!v) return ''
  if (typeof v === 'number') {
    // Serial Excel — converte com UTC
    const d = new Date(Math.round((v - 25569) * 86400 * 1000))
    if (isNaN(d)) return ''
    const dd=String(d.getUTCDate()).padStart(2,'0')
    const mm=String(d.getUTCMonth()+1).padStart(2,'0')
    const yy=d.getUTCFullYear()
    return `${dd}/${mm}/${yy}`
  }
  const s = String(v).trim()
  // Se vier como 'mm/dd/yyyy hh:mm:ss' (formato americano errado), corrige
  const mAm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}:\d{2}:\d{2})/)
  if (mAm) {
    const [,p1,p2,y,time] = mAm
    // Detecta se está no formato errado: mês > 12 nunca pode ser mês
    // Se p1 <= 12 e p2 > 12, então p1=mês, p2=dia (formato americano)
    if (+p2 > 12) return `${p2.padStart(2,'0')}/${p1.padStart(2,'0')}/${y} ${time}`
    // Se p1 > 12, já é dia no primeiro campo (formato BR correto)
    return s
  }
  return s
}

// ─── Lê TEMPO_DE_ESPERA: pode vir como fração decimal ou string hh:mm ─
const parseEsperaMin = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24 * 60  // fração de dia → minutos
  const s = String(v).trim()
  if (s.includes(':')) {
    const [h, m] = s.split(':').map(Number)
    return (h||0)*60 + (m||0)
  }
  return 0
}

// ─── buildFilter: período baseado nas datas da base ───────
const todayStr = () => {
  const n = new Date()
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}`
}
const buildFilter = (allDates, período, dateFrom, dateTo) => {
  if (!allDates.length) return () => true
  const sorted  = [...allDates].sort()
  const maxDate = sorted[sorted.length - 1]
  if (período === 'HOJE') {
    const today = todayStr()
    return (ds) => ds === today
  }
  if (período === 'ONTEM') {
    // Ontem = dia anterior ao maxDate da base
    const ref = new Date(maxDate + 'T00:00:00Z')
    ref.setUTCDate(ref.getUTCDate() - 1)
    const ontem = `${ref.getUTCFullYear()}-${String(ref.getUTCMonth()+1).padStart(2,'0')}-${String(ref.getUTCDate()).padStart(2,'0')}`
    return (ds) => ds === ontem
  }
  if (período === 'SEMANA') {
    const c = new Date(maxDate + 'T00:00:00Z')
    c.setUTCDate(c.getUTCDate() - 6)
    const cut = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
    return (ds) => ds >= cut && ds <= maxDate
  }
  if (período === 'MES') {
    const mes = maxDate.slice(0,7)
    return (ds) => ds.slice(0,7) === mes
  }
  if (período === 'ANO') {
    const ano = maxDate.slice(0,4)
    return (ds) => ds.slice(0,4) === ano
  }
  if (período === 'PERIODO') {
    const from = dateFrom || sorted[0]
    const to   = dateTo   || maxDate
    return (ds) => ds >= from && ds <= to
  }
  return () => true  // sem filtro
}

// ─── Sub-components ───────────────────────────────────────
function HBar({ label, value, max, color, unit='', rank, sub }) {
  const pct = max > 0 ? (value/max)*100 : 0
  return (
    <div style={{ marginBottom: 13 }}>
      <div style={{ display:'flex', justifyContent:'space-between', marginBottom:4, gap:8 }}>
        <div style={{ display:'flex', alignItems:'flex-start', gap:8, minWidth:0 }}>
          {rank!=null && <span style={{ fontSize:11, fontWeight:700, flexShrink:0, marginTop:1,
            color: rank<3 ? T.danger : T.muted, minWidth:22 }}>#{rank+1}</span>}
          <div style={{ minWidth:0 }}>
            <div style={{ fontSize:13, color:T.text, fontWeight:500,
              overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{label}</div>
            {sub && <div style={{ fontSize:10, color:T.muted, marginTop:1 }}>{sub}</div>}
          </div>
        </div>
        <span style={{ fontSize:13, fontWeight:700, color, flexShrink:0 }}>{value}{unit}</span>
      </div>
      <div style={{ background:T.border, borderRadius:6, height:6, overflow:'hidden' }}>
        <div style={{ height:'100%', borderRadius:6, background:color,
          width:`${pct}%`, transition:'width .6s ease' }} />
      </div>
    </div>
  )
}

function StatCard({ icon, label, value, sub, accent }) {
  return (
    <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:16,
      padding:'20px 22px', display:'flex', flexDirection:'column', gap:8,
      position:'relative', overflow:'hidden' }}>
      <div style={{ position:'absolute', top:-20, right:-20, width:80, height:80,
        borderRadius:'50%', background:accent, opacity:.07 }} />
      <div style={{ display:'flex', alignItems:'center', gap:10 }}>
        <span style={{ fontSize:20 }}>{icon}</span>
        <span style={{ fontSize:11, color:T.muted, textTransform:'uppercase',
          letterSpacing:'.08em', fontWeight:600 }}>{label}</span>
      </div>
      <div style={{ fontSize:36, fontWeight:800, color:T.text,
        lineHeight:1, letterSpacing:'-1px' }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:T.sub }}>{sub}</div>}
    </div>
  )
}

function SH({ children }) {
  return <h2 style={{ fontSize:13, fontWeight:700, textTransform:'uppercase',
    letterSpacing:'.12em', color:T.sub, marginBottom:16, marginTop:0 }}>{children}</h2>
}

function Card({ children, style={} }) {
  return <div style={{ background:T.card, border:`1px solid ${T.border}`,
    borderRadius:16, padding:22, ...style }}>{children}</div>
}

function Badge({ label, color }) {
  return <span style={{ fontSize:11, fontWeight:700, padding:'3px 9px',
    borderRadius:20, background:color+'22', color, whiteSpace:'nowrap' }}>{label}</span>
}

function Donut({ segments, size=120 }) {
  const r=46, cx=60, cy=60, circ=2*Math.PI*r
  const total = segments.reduce((a,s)=>a+s.value,0)||1
  let cum=0
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={T.border} strokeWidth="14" />
      {segments.map((seg,i)=>{
        const dash=(seg.value/total)*circ, off=circ-cum*circ/total
        cum+=seg.value
        return <circle key={i} cx={cx} cy={cy} r={r} fill="none"
          stroke={seg.color} strokeWidth="14"
          strokeDasharray={`${dash} ${circ-dash}`}
          strokeDashoffset={off} strokeLinecap="round" transform="rotate(-90 60 60)" />
      })}
      <text x="60" y="56" textAnchor="middle" fontSize="16" fontWeight="800" fill={T.text}>
        {total.toLocaleString('pt-BR')}</text>
      <text x="60" y="70" textAnchor="middle" fontSize="8" fill={T.muted}>registros</text>
    </svg>
  )
}

// ─── Gráfico moderno de STATUS (tipo planilha referência) ─
function StatusDashboard({ breakdown, total }) {
  const order = ['OK','ATRASO','ATRASO CRÍTICO','ATRASO GRAVE','Falta Médica',
                 'Remarcação Adm','Remarcação Médico','Remarcação médico','SEM PONTO']
  const items = order
    .map(k => ({ k, cfg: getStatusCfg(k), val: breakdown[k]||0 }))
    .concat(
      Object.keys(breakdown)
        .filter(k => !order.includes(k))
        .map(k => ({ k, cfg: getStatusCfg(k), val: breakdown[k] }))
    )
    .filter(x => x.val > 0)

  return (
    <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(180px,1fr))', gap:12 }}>
      {items.map(({k, cfg, val}) => {
        const pct = total > 0 ? ((val/total)*100).toFixed(1) : '0'
        const barW = total > 0 ? (val/total)*100 : 0
        return (
          <div key={k} style={{
            background: cfg.bg, border:`1px solid ${cfg.color}40`,
            borderRadius:14, padding:'16px 18px',
            position:'relative', overflow:'hidden',
          }}>
            {/* Barra de progresso na base */}
            <div style={{ position:'absolute', bottom:0, left:0, height:3,
              width:`${barW}%`, background:cfg.color, borderRadius:'0 0 0 14px',
              transition:'width .6s ease' }} />
            <div style={{ fontSize:10, fontWeight:700, color:cfg.color,
              textTransform:'uppercase', letterSpacing:'.08em', marginBottom:10 }}>
              {cfg.label}
            </div>
            <div style={{ fontSize:36, fontWeight:900, color:cfg.color,
              lineHeight:1, letterSpacing:'-1px' }}>
              {val.toLocaleString('pt-BR')}
            </div>
            <div style={{ fontSize:11, color:cfg.color+'99', marginTop:6, fontWeight:600 }}>
              {pct}% do total
            </div>
          </div>
        )
      })}
    </div>
  )
}

// ─── Gráfico Sem Ponto / Ausências ────────────────────────
function AusenciasCard({ rows, cols }) {
  const items = useMemo(() => {
    const m = {}
    rows.forEach(d => {
      const entrada = d[cols.hrEntrada]
      const isVazio = !entrada || entrada === '' || entrada === 'NaT'
      if (!isVazio) return

      const med    = String(d[cols.medico]||'').trim() || 'Sem Nome'
      const status = String(d[cols.status]||'').trim()
      const motivo = String(d[cols.motivoCancelamento]||'').trim()
      const unid   = String(d[cols.unidade]||'').trim()
      const pacts  = Number(d[cols.qtPacts])||0

      // Classificação
      let tipo
      if (motivo) tipo = motivo
      else        tipo = 'Sem Ponto Registrado'

      const key = `${med}||${unid}`
      if (!m[key]) m[key] = { med, unid, status, tipo, pacts: 0, agendas: 0 }
      m[key].pacts   += pacts
      m[key].agendas += 1
    })
    return Object.values(m).sort((a,b) => b.pacts - a.pacts).slice(0, 15)
  }, [rows, cols])

  if (!items.length) return (
    <div style={{ color:T.muted, fontSize:13 }}>Nenhuma ausência identificada.</div>
  )

  return (
    <div style={{ overflowX:'auto' }}>
      <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12.5 }}>
        <thead>
          <tr style={{ borderBottom:`1px solid ${T.border}` }}>
            {['#','Médico','Unidade','Status','Motivo/Cancelamento','Agendas','Pac. Impactados'].map(h=>(
              <th key={h} style={{ padding:'9px 12px', textAlign:'left', color:T.muted,
                fontWeight:600, fontSize:10.5, textTransform:'uppercase',
                letterSpacing:'.06em', whiteSpace:'nowrap' }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {items.map((r, i) => {
            const cfg = getStatusCfg(r.status)
            return (
              <tr key={i} className="row-table"
                style={{ borderBottom:`1px solid ${T.border}`, transition:'background .15s' }}>
                <td style={{ padding:'10px 12px', color:T.muted, fontSize:11 }}>{i+1}</td>
                <td style={{ padding:'10px 12px', fontWeight:600, color:T.text, whiteSpace:'nowrap' }}>{r.med}</td>
                <td style={{ padding:'10px 12px', color:T.sub, maxWidth:200,
                  overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.unid}</td>
                <td style={{ padding:'10px 12px' }}>
                  <span style={{ fontSize:10, fontWeight:700, padding:'2px 8px', borderRadius:20,
                    background:cfg.color+'22', color:cfg.color }}>{r.status}</span>
                </td>
                <td style={{ padding:'10px 12px', color:T.sub, maxWidth:220,
                  overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                  {r.tipo || '—'}
                </td>
                <td style={{ padding:'10px 12px', color:T.accent, fontWeight:700, textAlign:'center' }}>{r.agendas}</td>
                <td style={{ padding:'10px 12px', color:T.warning, fontWeight:700, textAlign:'center' }}>{r.pacts}</td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

// ─── Gráfico Motivo do Impacto na Espera ─────────────────
function ImpactoEsperaCard({ rows, cols }) {
  const items = useMemo(() => {
    const m = {}
    rows.forEach(d => {
      const unid   = String(d[cols.unidade]||'').trim() || 'Sem Unidade'
      const status = String(d[cols.status]||'').trim()
      const espMin = parseEsperaMin(d[cols.espera])
      const pacts  = Number(d[cols.qtPacts])||0
      if (!m[unid]) m[unid] = { unid, totalEspMin:0, pacts:0, statusMap:{} }
      m[unid].totalEspMin += espMin
      m[unid].pacts       += pacts
      m[unid].statusMap[status] = (m[unid].statusMap[status]||0) + 1
    })
    return Object.values(m)
      .filter(x => x.totalEspMin > 0 || x.pacts > 0)
      .map(x => {
        // Motivo predominante
        const topStatus = Object.entries(x.statusMap)
          .sort((a,b)=>b[1]-a[1])[0]?.[0] || 'OK'
        const allOK = Object.keys(x.statusMap).every(s => s === 'OK')
        const motivo = allOK ? 'Fluxo Médico Normal' : (getStatusCfg(topStatus).label)
        const motivoColor = allOK ? T.success : getStatusCfg(topStatus).color
        return { ...x, motivo, motivoColor, topStatus }
      })
      .sort((a,b) => b.totalEspMin - a.totalEspMin)
      .slice(0, 10)
  }, [rows, cols])

  if (!items.length) return (
    <div style={{ color:T.muted, fontSize:13 }}>Sem dados de espera nos filtros atuais.</div>
  )

  const maxMin = items[0]?.totalEspMin || 1

  return (
    <div>
      {items.map((r, i) => (
        <div key={r.unid} style={{ marginBottom:16 }}>
          <div style={{ display:'flex', justifyContent:'space-between',
            alignItems:'flex-start', marginBottom:6, gap:8 }}>
            <div style={{ display:'flex', alignItems:'flex-start', gap:8, minWidth:0 }}>
              <span style={{ fontSize:11, fontWeight:700, color: i<3?T.danger:T.muted,
                minWidth:22, flexShrink:0, marginTop:1 }}>#{i+1}</span>
              <div style={{ minWidth:0 }}>
                <div style={{ fontSize:13, color:T.text, fontWeight:500,
                  overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.unid}</div>
                <div style={{ display:'flex', alignItems:'center', gap:6, marginTop:3 }}>
                  <div style={{ width:7, height:7, borderRadius:'50%', background:r.motivoColor, flexShrink:0 }} />
                  <span style={{ fontSize:11, color:r.motivoColor, fontWeight:600 }}>{r.motivo}</span>
                  <span style={{ fontSize:10, color:T.muted }}>· {r.pacts} pacientes</span>
                </div>
              </div>
            </div>
            <span style={{ fontSize:13, fontWeight:700, color:i<3?T.danger:T.warning,
              flexShrink:0 }}>{fmtMin(r.totalEspMin)}</span>
          </div>
          <div style={{ background:T.border, borderRadius:6, height:6, overflow:'hidden' }}>
            <div style={{ height:'100%', borderRadius:6,
              background: i<3 ? T.danger : T.warning,
              width:`${(r.totalEspMin/maxMin)*100}%`, transition:'width .6s ease' }} />
          </div>
        </div>
      ))}
    </div>
  )
}

// ─── Período selector ─────────────────────────────────────
const PERIODOS = [
  { key:'HOJE',    label:'Hoje'    },
  { key:'ONTEM',   label:'Ontem'   },
  { key:'SEMANA',  label:'Semana'  },
  { key:'MES',     label:'Mês'     },
  { key:'ANO',     label:'Ano'     },
  { key:'PERIODO', label:'Período' },
]

function PeriodoSelector({ value, onChange, infoLabel, dateFrom, dateTo, onDateFrom, onDateTo, allDates }) {
  const minDate = allDates[0] || ''
  const maxDate = allDates[allDates.length-1] || ''
  return (
    <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
      <div style={{ display:'flex', gap:3, background:T.card,
        border:`1px solid ${T.border}`, borderRadius:10, padding:4 }}>
        {PERIODOS.map(p=>(
          <button key={p.key} onClick={()=>onChange(p.key)} style={{
            padding:'6px 11px', borderRadius:7, border:'none',
            fontSize:12, fontWeight:700, cursor:'pointer', transition:'all .15s',
            background: value===p.key
              ? `linear-gradient(135deg,${T.accent},${T.accentB})` : 'transparent',
            color: value===p.key ? '#000' : T.muted,
            whiteSpace:'nowrap',
          }}>{p.label}</button>
        ))}
      </div>
      {value === 'PERIODO' && (
        <div style={{ display:'flex', alignItems:'center', gap:6 }}>
          <input type="date" value={dateFrom} min={minDate} max={maxDate}
            onChange={e=>onDateFrom(e.target.value)}
            style={{ background:T.card, border:`1px solid ${T.accent}55`, borderRadius:8,
              color:T.text, fontSize:12, padding:'6px 10px', outline:'none',
              colorScheme:'dark', cursor:'pointer' }} />
          <span style={{ color:T.muted, fontSize:12 }}>→</span>
          <input type="date" value={dateTo} min={minDate} max={maxDate}
            onChange={e=>onDateTo(e.target.value)}
            style={{ background:T.card, border:`1px solid ${T.accent}55`, borderRadius:8,
              color:T.text, fontSize:12, padding:'6px 10px', outline:'none',
              colorScheme:'dark', cursor:'pointer' }} />
        </div>
      )}
      {infoLabel && <span style={{ fontSize:11, color:T.muted }}>{infoLabel}</span>}
    </div>
  )
}


// ─── Supabase config ─────────────────────────────────────
const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'sb_publishable_x32NVeFMKLK9kLJfdunngg_GfxpTo1P'
const CHUNK_SIZE = 4000   // linhas por chunk

const sbFetch = (path, opts = {}) => {
  const isWrite = opts.method && ['POST','PATCH','PUT','DELETE'].includes(opts.method)
  return fetch(`${SB_URL}/rest/v1/${path}`, {
    headers: {
      'apikey':        SB_KEY,
      'Authorization': `Bearer ${SB_KEY}`,
      'Content-Type':  'application/json',
      ...(isWrite ? { 'Prefer': 'return=minimal' } : {}),
      ...opts.headers,
    },
    ...opts,
  })
}

// Extrai a data (YYYY-MM-DD) de uma linha
const rowDateStr = (r) => {
  const v = r['DATA_AGENDA'] ?? r[Object.keys(r).find(k => k.includes('DATA')) || ''] ?? ''
  return serialToDateStr(v)
}

// ─── Main ─────────────────────────────────────────────────
export default function Home() {
  const [dados,       setDados]       = useState([])
  const [uf,          setUf]          = useState('TODOS')
  const [status,      setStatus]      = useState('TODOS')
  const [loading,     setLoading]     = useState(false)
  const [storing,     setStoring]     = useState(false)
  const [storeStatus, setStoreStatus] = useState('')   // mensagem de progresso
  const [search,      setSearch]      = useState('')
  const [período,     setPeriodo]     = useState('TODOS')
  const [horaFilt,    setHoraFilt]    = useState('TODAS')
  const [dateFrom,    setDateFrom]    = useState('')
  const [dateTo,      setDateTo]      = useState('')
  const [timestamp,   setTimestamp]   = useState('')
  const [storageInfo, setStorageInfo] = useState({ dias: 0, total: 0 })
  const [initLoading, setInitLoading] = useState(true)

  // ── Carrega TODOS os chunks do Supabase ao abrir ───────
  useEffect(() => {
    const loadFromSupabase = async () => {
      try {
        setStoreStatus('Conectando ao Supabase…')

        // 1. Busca metadados (timestamp)
        const metaRes = await sbFetch('mh_meta?select=ts&id=eq.1')
        if (metaRes.ok) {
          const meta = await metaRes.json()
          if (meta[0]?.ts) setTimestamp(meta[0].ts)
        } else {
          const txt = await metaRes.text()
          setStoreStatus(`Erro meta ${metaRes.status}: ${txt.slice(0,120)}`)
          setInitLoading(false)
          return
        }

        // 2. Busca todos os chunks com paginação via limit/offset
        setStoreStatus('Buscando dados…')
        let allChunks = [], offset = 0
        const pageSize = 500
        while (true) {
          const r = await sbFetch(
            `mh_chunks?select=data_agenda,chunk_idx,rows_json&order=data_agenda.asc,chunk_idx.asc&limit=${pageSize}&offset=${offset}`
          )
          if (!r.ok) {
            const txt = await r.text()
            setStoreStatus(`Erro chunks ${r.status}: ${txt.slice(0,120)}`)
            setInitLoading(false)
            return
          }
          const batch = await r.json()
          if (!Array.isArray(batch) || batch.length === 0) break
          allChunks = allChunks.concat(batch)
          setStoreStatus(`Carregando… ${allChunks.length} chunks`)
          if (batch.length < pageSize) break
          offset += pageSize
        }

        if (allChunks.length > 0) {
          setStoreStatus(`Processando ${allChunks.length} chunks…`)
          let allRows = []
          for (const c of allChunks) {
            try { allRows = allRows.concat(JSON.parse(c.rows_json)) } catch(e) {}
          }
          const diasSet = new Set(allChunks.map(c => c.data_agenda))
          setDados(allRows)
          setStorageInfo({ dias: diasSet.size, total: allRows.length })
          setPeriodo('HOJE')
          setStoreStatus(`☁ ${diasSet.size} dias · ${allRows.length.toLocaleString('pt-BR')} registros`)
          setTimeout(() => setStoreStatus(''), 4000)
        } else {
          setStoreStatus('Supabase vazio — carregue uma planilha')
          setTimeout(() => setStoreStatus(''), 4000)
        }
      } catch (e) {
        console.error('Supabase load error:', e)
        setStoreStatus(`Erro: ${e.message}`)
      }
      setInitLoading(false)
    }
    loadFromSupabase()
  }, [])

  // ── Salva no Supabase (upsert por data) ───────────────
  const saveToSupabase = useCallback(async (newRows, ts) => {
    setStoring(true)
    try {
      // Datas dos novos dados
      const newDates = [...new Set(newRows.map(rowDateStr).filter(Boolean))]
      setStoreStatus(`Removendo ${newDates.length} dia(s) antigo(s)…`)

      // Apaga chunks das datas que vieram na nova planilha (upsert por data)
      for (const date of newDates) {
        await sbFetch(
          `mh_chunks?data_agenda=eq.${date}`,
          { method: 'DELETE' }
        )
      }

      // Agrupa novas linhas por data e insere em chunks
      const byDate = {}
      newRows.forEach(r => {
        const ds = rowDateStr(r)
        if (!ds) return
        if (!byDate[ds]) byDate[ds] = []
        byDate[ds].push(r)
      })

      let inserted = 0
      const datesSorted = Object.keys(byDate).sort()
      for (const date of datesSorted) {
        const rows = byDate[date]
        const totalChunks = Math.ceil(rows.length / CHUNK_SIZE)
        setStoreStatus(`Salvando ${date} (${rows.length} linhas)…`)
        for (let ci = 0; ci < totalChunks; ci++) {
          const slice = rows.slice(ci * CHUNK_SIZE, (ci + 1) * CHUNK_SIZE)
          await sbFetch('mh_chunks', {
            method: 'POST',
            headers: { 'Prefer': 'resolution=merge-duplicates,return=minimal' },
            body: JSON.stringify({
              data_agenda:  date,
              chunk_idx:    ci,
              rows_json:    JSON.stringify(slice),
              total_linhas: slice.length,
              uploaded_at:  new Date().toISOString(),
            }),
          })
          inserted += slice.length
        }
      }

      // Atualiza metadados (timestamp)
      await sbFetch('mh_meta?id=eq.1', {
        method: 'PATCH',
        headers: { 'Prefer': 'return=minimal' },
        body: JSON.stringify({ ts: ts, updated_at: new Date().toISOString() }),
      })

      // Recarrega tudo do Supabase com paginação
      setStoreStatus('Recarregando dados completos…')
      let reloadChunks = [], rOffset = 0
      const rPageSize = 500
      while (true) {
        const r = await sbFetch(
          `mh_chunks?select=data_agenda,chunk_idx,rows_json&order=data_agenda.asc,chunk_idx.asc&limit=${rPageSize}&offset=${rOffset}`
        )
        if (!r.ok) break
        const batch = await r.json()
        if (!Array.isArray(batch) || batch.length === 0) break
        reloadChunks = reloadChunks.concat(batch)
        if (batch.length < rPageSize) break
        rOffset += rPageSize
      }
      if (reloadChunks.length > 0) {
        let allRows = []
        for (const c of reloadChunks) {
          try { allRows = allRows.concat(JSON.parse(c.rows_json)) } catch(e) {}
        }
        const diasSet = new Set(reloadChunks.map(c => c.data_agenda))
        setDados(allRows)
        setStorageInfo({ dias: diasSet.size, total: allRows.length })
      }

      setInitLoading(false)
      setStoreStatus('✓ Salvo no Supabase')
      setTimeout(() => setStoreStatus(''), 3000)
    } catch (e) {
      console.error('Supabase save error:', e)
      setStoreStatus('⚠ Erro ao salvar no Supabase — dados visíveis localmente')
    }
    setStoring(false)
  }, [])

  // ── Limpar todos os dados do Supabase ─────────────────
  const clearStorage = async () => {
    if (!confirm('Apagar TODOS os dados do Supabase? Isso não pode ser desfeito.')) return
    setStoring(true)
    setStoreStatus('Apagando…')
    try {
      await sbFetch('mh_chunks?id=gt.0', { method: 'DELETE' })
      await sbFetch('mh_meta?id=eq.1', {
        method: 'PATCH',
        body: JSON.stringify({ ts: '' }),
      })
    } catch(e) { console.error(e) }
    setDados([])
    setStorageInfo({ dias: 0, total: 0 })
    setTimestamp('')
    setStoreStatus('')
    setStoring(false)
  }

  const handleUpload = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setLoading(true)

    const now = new Date()
    const pad = n => String(n).padStart(2,'0')
    const ts = `${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ` +
               `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`
    setTimestamp(ts)

    const buf  = await file.arrayBuffer()
    const wb   = XLSX.read(buf, { type:'buffer' })
    const ws   = wb.Sheets[wb.SheetNames[0]]
    const json = XLSX.utils.sheet_to_json(ws, { range:3, defval:'' })

    setUf('TODOS'); setStatus('TODOS'); setSearch('')
    setPeriodo('HOJE'); setHoraFilt('TODAS'); setDateFrom(''); setDateTo('')
    // Mostra os dados imediatamente na tela enquanto salva em background
    setDados(json)
    setInitLoading(false)
    setLoading(false)

    // Salva no Supabase em background (não bloqueia a tela)
    saveToSupabase(json, ts)
  }

  // ── Detecção de colunas ──────────────────────────────
  const cols = useMemo(() => {
    if (!dados.length) return {}
    const k = Object.keys(dados[0])
    const find = (...terms) =>
      k.find(c => terms.some(t => c === t)) ||
      k.find(c => terms.some(t => c.toLowerCase().includes(t.toLowerCase()))) || ''
    return {
      unidade:            find('NM_LOCAL','UNIDADE'),
      medico:             find('NM_MEDICO','MEDICO'),
      esp:                find('DS_ESPECIALIDADE','ESPECIALIDADE'),
      status:             find('STATUS'),
      espera:             find('TEMPO_DE_ESPERA'),
      qtPacts:            find('QT_PACIENTES_AGUARDANDO','QTS PACTS'),
      uf:                 find('UF'),
      cidade:             find('CIDADE'),
      data:               find('DATA_AGENDA','DATA'),
      hrInicio:           find('HR_INICIO'),
      hrEntrada:          find('HR_ENTRADA'),
      motivoCancelamento: find('MOTIVO_CANCELAMENTO'),
      dtRegistro:         find('DT_REGISTRO'),
      hrRegistroEspera:   find('HR_REGISTRO_ESPERA'),
    }
  }, [dados])

  // ── Pré-processa: dateStr + hora ─────────────────────
  const dadosComData = useMemo(() => {
    if (!dados.length) return []
    return dados.map(d => {
      // Hora de HR_INICIO
      const hrV = d[cols.hrInicio]
      let hora = 99
      if (typeof hrV === 'number') hora = Math.floor(hrV * 24)
      else if (hrV && typeof hrV === 'string') {
        const m = hrV.match(/(\d{1,2}):/)
        if (m) hora = parseInt(m[1])
      }
      return { ...d, _dateStr: serialToDateStr(d[cols.data]), _hora: hora }
    })
  }, [dados, cols])

  // ── Datas únicas + filtro período ───────────────────
  const allDates = useMemo(() =>
    [...new Set(dadosComData.map(d=>d._dateStr).filter(Boolean))].sort(),
    [dadosComData])

  const periodoFn = useMemo(() => buildFilter(allDates, período, dateFrom, dateTo), [allDates, período, dateFrom, dateTo])

  const dadosPorPeriodo = useMemo(() =>
    dadosComData.filter(d => periodoFn(d._dateStr)),
    [dadosComData, periodoFn])

  // ── Horas disponíveis ────────────────────────────────
  const horasDisp = useMemo(() => {
    const hs = [...new Set(dadosPorPeriodo.map(d=>d._hora).filter(h=>h>=0&&h<=23))].sort((a,b)=>a-b)
    return hs
  }, [dadosPorPeriodo])

  // ── Filtros UF / status / busca / hora ───────────────
  const ufs = useMemo(() =>
    [...new Set(dadosPorPeriodo.map(d=>String(d[cols.uf]||'').trim()).filter(Boolean))].sort(),
    [dadosPorPeriodo, cols])
  const statuses = useMemo(() =>
    [...new Set(dadosPorPeriodo.map(d=>String(d[cols.status]||'').trim()).filter(Boolean))].sort(),
    [dadosPorPeriodo, cols])

  const filtered = useMemo(() => {
    let r = dadosPorPeriodo
    if (uf !== 'TODOS')       r = r.filter(d => String(d[cols.uf]||'').trim() === uf)
    if (status !== 'TODOS')   r = r.filter(d => String(d[cols.status]||'').trim() === status)
    if (horaFilt !== 'TODAS') r = r.filter(d => d._hora === Number(horaFilt))
    if (search)               r = r.filter(d =>
      [d[cols.unidade], d[cols.medico], d[cols.esp], d[cols.cidade]]
        .some(v => String(v||'').toLowerCase().includes(search.toLowerCase())))
    return r
  }, [dadosPorPeriodo, uf, status, horaFilt, search, cols])

  // ── KPIs ─────────────────────────────────────────────
  const totalRegistros = filtered.length
  const totalUnidades  = new Set(filtered.map(d=>d[cols.unidade]).filter(Boolean)).size
  const totalMedicos   = new Set(filtered.map(d=>String(d[cols.medico]||'').trim()).filter(Boolean)).size
  const emAtraso       = filtered.filter(d=>String(d[cols.status]||'').toUpperCase().includes('ATRASO')).length
  const taxaAtraso     = totalRegistros>0 ? ((emAtraso/totalRegistros)*100).toFixed(1) : 0
  const totalEsperaMin = filtered.reduce((a,d)=>a+parseEsperaMin(d[cols.espera]),0)
  const mediaEsperaMin = totalRegistros>0 ? totalEsperaMin/totalRegistros : 0
  const totalPacAguard = filtered.reduce((a,d)=>a+(Number(d[cols.qtPacts])||0),0)
  const semPontoCount  = filtered.filter(d => {
    const e = d[cols.hrEntrada]
    return !e || e === '' || e === 'NaT'
  }).length

  // ── Status breakdown ─────────────────────────────────
  const statusBreakdown = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s=String(d[cols.status]||'OK').trim(); m[s]=(m[s]||0)+1 })
    return m
  }, [filtered, cols])

  // ── Top unidades por atraso ───────────────────────────
  const topUnidadesAtraso = useMemo(() => {
    const m = {}
    filtered.filter(d=>String(d[cols.status]||'').toUpperCase().includes('ATRASO'))
      .forEach(d=>{ const u=d[cols.unidade]||'Sem Unidade'; m[u]=(m[u]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])

  // ── Top médicos por espera ────────────────────────────
  const topMedicos = useMemo(() => {
    const m = {}
    filtered.forEach(d => {
      const med   = String(d[cols.medico]||'').trim(); if(!med) return
      const espM  = parseEsperaMin(d[cols.espera])
      const dateS = d._dateStr||''; const unid=String(d[cols.unidade]||'').trim()
      if(!m[med]) m[med]={ med, totalEspMin:0, agendasSet:new Set(), pacts:0 }
      m[med].totalEspMin += espM
      m[med].agendasSet.add(`${dateS}||${unid}`)
      m[med].pacts += Number(d[cols.qtPacts])||0
    })
    return Object.values(m).filter(x=>x.totalEspMin>0)
      .map(x=>({ med:x.med, totalMin:parseFloat(x.totalEspMin.toFixed(0)),
        agendas:x.agendasSet.size, pacts:x.pacts }))
      .sort((a,b)=>b.totalMin-a.totalMin).slice(0,10)
  }, [filtered, cols])

  // ── Especialidades ─────────────────────────────────────
  const espBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d=>{ const e=d[cols.esp]||'Outro'; m[e]=(m[e]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxEsp = espBreak[0]?.total||1

  // ── UF breakdown ──────────────────────────────────────
  const ufBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d=>{ const u=String(d[cols.uf]||'?').trim(); m[u]=(m[u]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxUF = ufBreak[0]?.total||1

  // ── Label período ─────────────────────────────────────
  const períodoLabel = useMemo(() => {
    if (!allDates.length) return ''
    const sorted=[...allDates].sort(), max=sorted[sorted.length-1]
    const fmt=s=>s.split('-').reverse().join('/')
    if (período==='HOJE') return `Hoje · ${fmt(todayStr())}`
    if (período==='ONTEM') {
      const ref = new Date(max+'T00:00:00Z')
      ref.setUTCDate(ref.getUTCDate()-1)
      const ontemStr = `${ref.getUTCFullYear()}-${String(ref.getUTCMonth()+1).padStart(2,'0')}-${String(ref.getUTCDate()).padStart(2,'0')}`
      return `Ontem · ${fmt(ontemStr)}`
    }
    if (período==='SEMANA') {
      const c=new Date(max+'T00:00:00Z'); c.setUTCDate(c.getUTCDate()-6)
      const cut=`${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
      const semMin=sorted.find(d=>d>=cut)||cut
      return `${fmt(semMin)} → ${fmt(max)}`
    }
    if (período==='MES') {
      const mes=max.slice(0,7), mesMin=sorted.find(d=>d.slice(0,7)===mes)||max
      const [y,m]=max.split('-')
      const nome=new Date(+y,+m-1).toLocaleString('pt-BR',{month:'long'})
      return `${nome.charAt(0).toUpperCase()+nome.slice(1)} · ${fmt(mesMin)} → ${fmt(max)}`
    }
    if (período==='ANO') {
      const ano=max.slice(0,4), anoMin=sorted.find(d=>d.slice(0,4)===ano)||max
      return `Ano ${ano} · ${fmt(anoMin)} → ${fmt(max)}`
    }
    if (período==='PERIODO' && dateFrom && dateTo) {
      return `${fmt(dateFrom)} → ${fmt(dateTo)}`
    }
    return ''
  }, [allDates, período, dateFrom, dateTo])

  const hasData = dados.length > 0
  const isInit  = initLoading

  return (
    <div style={{ background:T.bg, minHeight:'100vh',
      fontFamily:"'DM Sans','Segoe UI',sans-serif", color:T.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:6px;background:${T.bg}}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
        select option{background:${T.surface};color:${T.text}}
        .btn-upload:hover{opacity:.9;transform:translateY(-1px)}
        .row-table:hover{background:#162033!important}
        input::placeholder{color:${T.muted}}
        select{appearance:auto}
      `}</style>

      {/* ── Topbar ── */}
      <div style={{
        background:T.surface, borderBottom:`1px solid ${T.border}`,
        padding:'0 36px', display:'flex', alignItems:'center',
        justifyContent:'space-between', height:64,
        position:'sticky', top:0, zIndex:100,
      }}>
        <div style={{ display:'flex', alignItems:'center', gap:14 }}>
          <div style={{
            width:32, height:32, borderRadius:8, fontSize:16,
            background:`linear-gradient(135deg,${T.accent},${T.accentB})`,
            display:'flex', alignItems:'center', justifyContent:'center',
          }}>🏥</div>
          <div>
            <div style={{ fontSize:15, fontWeight:700 }}>Monitor Hospitalar</div>
            <div style={{ fontSize:11, color:T.muted }}>Atrasos Médicos · Operacional</div>
          </div>
        </div>

        <div style={{ display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
          {hasData && (
            <>
              <PeriodoSelector value={período} onChange={p=>{setPeriodo(p);setHoraFilt('TODAS');setDateFrom('');setDateTo('')}}
                infoLabel={`${totalRegistros.toLocaleString('pt-BR')} reg.${períodoLabel ? ' · '+períodoLabel : ''}`}
                dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
                allDates={allDates} />

              {/* Filtro de hora */}
              <select value={horaFilt} onChange={e=>setHoraFilt(e.target.value)} style={{
                background:T.card, border:`1px solid ${T.border}`, borderRadius:9,
                color:T.text, fontSize:12, padding:'7px 11px', outline:'none', cursor:'pointer' }}>
                <option value="TODAS">Todas as horas</option>
                {horasDisp.map(h=>(
                  <option key={h} value={h}>{String(h).padStart(2,'0')}h</option>
                ))}
              </select>
            </>
          )}

          {timestamp && (
            <div style={{ fontSize:10, color:T.muted, textAlign:'right', lineHeight:1.4 }}>
              <div>Verificação ponto</div>
              <div style={{ color:T.sub, fontWeight:600 }}>{timestamp}</div>
            </div>
          )}

          {/* Supabase status */}
          <div style={{ fontSize:10, textAlign:'right', lineHeight:1.6, minWidth:120 }}>
            {storageInfo.dias > 0 && !storeStatus && (
              <>
                <div style={{ color:T.success, fontWeight:700 }}>
                  ☁ Supabase · {storageInfo.dias} {storageInfo.dias===1?'dia':'dias'}
                </div>
                <div style={{ color:T.muted }}>{storageInfo.total.toLocaleString('pt-BR')} registros</div>
              </>
            )}
            {storeStatus && (
              <div style={{ color:T.warning, fontWeight:600 }}>{storeStatus}</div>
            )}
          </div>
          {storageInfo.dias > 0 && !storing && (
            <button onClick={clearStorage} title="Apagar todos os dados do Supabase"
              style={{ background:'transparent', border:`1px solid ${T.danger}44`,
                borderRadius:9, color:T.danger, fontSize:11, padding:'7px 11px',
                cursor:'pointer', whiteSpace:'nowrap' }}>🗑 Limpar</button>
          )}
          <label className="btn-upload" style={{
            background: storing
              ? T.border
              : `linear-gradient(135deg,${T.accent},${T.accentB})`,
            color:'#000', fontWeight:700, fontSize:13,
            padding:'9px 18px', borderRadius:10, cursor: storing ? 'default' : 'pointer',
            transition:'all .2s', whiteSpace:'nowrap',
          }}>
            {loading ? 'Lendo…' : storing ? 'Salvando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{ display:'none' }}
              onChange={handleUpload} disabled={loading || storing} />
          </label>
        </div>
      </div>

      <div style={{ padding:'28px 36px' }}>

        {/* ── Empty / Loading ── */}
        {isInit && (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center',
            justifyContent:'center', minHeight:'calc(100vh - 140px)', gap:16 }}>
            <div style={{ fontSize:40 }}>⏳</div>
            <div style={{ fontSize:18, color:T.muted }}>Conectando ao Supabase…</div>
            {storeStatus && (
              <div style={{ fontSize:13, color:T.warning, maxWidth:600, textAlign:'center',
                background:T.card, padding:'12px 20px', borderRadius:10,
                border:`1px solid ${T.warning}44` }}>{storeStatus}</div>
            )}
          </div>
        )}
        {!isInit && !hasData && (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center',
            justifyContent:'center', minHeight:'calc(100vh - 140px)', gap:16 }}>
            {storeStatus ? (
              <>
                <div style={{ fontSize:40 }}>⚠️</div>
                <div style={{ fontSize:18, color:T.warning, fontWeight:700 }}>Problema na conexão</div>
                <div style={{ fontSize:13, color:T.warning, maxWidth:600, textAlign:'center',
                  background:T.card, padding:'12px 20px', borderRadius:10,
                  border:`1px solid ${T.warning}44` }}>{storeStatus}</div>
              </>
            ) : (
              <>
                <div style={{ fontSize:56 }}>📋</div>
                <div style={{ fontSize:22, fontWeight:700 }}>Nenhuma planilha carregada</div>
                <div style={{ color:T.muted, fontSize:14, textAlign:'center' }}>
                  Clique em "+ Carregar Planilha" para começar.<br />
                  Os dados ficam salvos — carregue o acumulado de cada dia e o dashboard acumula tudo.
                </div>
              </>
            )}
          </div>
        )}

        {!isInit && hasData && (<>

          {/* ── Filtros ── */}
          <div style={{ display:'flex', gap:10, marginBottom:24, flexWrap:'wrap', alignItems:'center' }}>
            <input value={search} onChange={e=>setSearch(e.target.value)}
              placeholder="Buscar unidade, médico, especialidade…"
              style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:10,
                color:T.text, fontSize:13, padding:'9px 14px', outline:'none', width:280 }} />
            <select value={uf} onChange={e=>setUf(e.target.value)} style={{
              background:T.card, border:`1px solid ${T.border}`, borderRadius:10,
              color:T.text, fontSize:13, padding:'9px 13px', outline:'none', cursor:'pointer' }}>
              <option value="TODOS">Todos os Estados</option>
              {ufs.map(u=><option key={u}>{u}</option>)}
            </select>
            <select value={status} onChange={e=>setStatus(e.target.value)} style={{
              background:T.card, border:`1px solid ${T.border}`, borderRadius:10,
              color:T.text, fontSize:13, padding:'9px 13px', outline:'none', cursor:'pointer' }}>
              <option value="TODOS">Todos os Status</option>
              {statuses.map(s=><option key={s}>{s}</option>)}
            </select>
            {(uf!=='TODOS'||status!=='TODOS'||search||horaFilt!=='TODAS') && (
              <button onClick={()=>{setUf('TODOS');setStatus('TODOS');setSearch('');setHoraFilt('TODAS')}}
                style={{ background:'transparent', border:`1px solid ${T.border}`,
                  borderRadius:10, color:T.muted, fontSize:13,
                  padding:'9px 13px', cursor:'pointer' }}>✕ Limpar</button>
            )}
          </div>

          {/* ── KPIs ── */}
          <div style={{ display:'grid', gridTemplateColumns:'repeat(5,1fr)', gap:14, marginBottom:22 }}>
            <StatCard icon="🩺" label="Total Registros"
              value={totalRegistros.toLocaleString('pt-BR')}
              sub={`${totalUnidades} unidades · ${totalMedicos} médicos`} accent={T.accent} />
            <StatCard icon="⚠️" label="Em Atraso"
              value={emAtraso.toLocaleString('pt-BR')}
              sub={`${taxaAtraso}% do total`} accent={T.danger} />
            <StatCard icon="⏱️" label="Espera Média"
              value={fmtMin(mediaEsperaMin)} sub="por atendimento" accent={T.warning} />
            <StatCard icon="👥" label="Pacientes Aguardando"
              value={totalPacAguard.toLocaleString('pt-BR')}
              sub="na fila agora" accent={T.success} />
            <StatCard icon="🚫" label="Sem Ponto"
              value={semPontoCount.toLocaleString('pt-BR')}
              sub="HR_ENTRADA vazia" accent={T.purple} />
          </div>

          {/* ── GRÁFICO STATUS (NOVO — tipo planilha de referência) ── */}
          <Card style={{ marginBottom:18 }}>
            <SH>📊 Distribuição de Status — Visão Operacional</SH>
            <StatusDashboard breakdown={statusBreakdown} total={totalRegistros} />
          </Card>

          {/* ── Row: Unidades + UFs ── */}
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16, marginBottom:16 }}>
            <Card>
              <SH>🔴 Unidades com Mais Atrasos</SH>
              {topUnidadesAtraso.map((u,i)=>(
                <HBar key={u.nome} rank={i} label={u.nome} value={u.total}
                  max={topUnidadesAtraso[0]?.total||1}
                  color={i===0?T.danger:i===1?'#FF7A00':T.warning}
                  unit=" atrasos" />
              ))}
              {!topUnidadesAtraso.length && <div style={{color:T.muted,fontSize:13}}>Nenhum atraso.</div>}
            </Card>
            <Card>
              <SH>📍 Registros por Estado (UF)</SH>
              {ufBreak.map(u=>(
                <HBar key={u.nome} label={u.nome} value={u.total} max={maxUF} color={T.accent} />
              ))}
            </Card>
          </div>

          {/* ── Médicos com Maior Espera ── */}
          <Card style={{ marginBottom:16 }}>
            <SH>👨‍⚕️ Médicos com Maior Tempo de Espera (acumulado)</SH>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
              <div>
                {topMedicos.slice(0,5).map((m,i)=>(
                  <HBar key={m.med} rank={i} label={m.med}
                    value={fmtMin(m.totalMin)} max={topMedicos[0]?.totalMin||1}
                    color={i<3?T.danger:T.warning}
                    sub={`${m.agendas} ${m.agendas===1?'agenda':'agendas'} · ${m.pacts.toLocaleString('pt-BR')} pacientes`}
                  />
                ))}
              </div>
              <div>
                {topMedicos.slice(5,10).map((m,i)=>(
                  <HBar key={m.med} rank={i+5} label={m.med}
                    value={fmtMin(m.totalMin)} max={topMedicos[0]?.totalMin||1}
                    color={T.warning}
                    sub={`${m.agendas} ${m.agendas===1?'agenda':'agendas'} · ${m.pacts.toLocaleString('pt-BR')} pacientes`}
                  />
                ))}
              </div>
            </div>
            {!topMedicos.length && <div style={{color:T.muted,fontSize:13}}>Nenhum dado de espera.</div>}
          </Card>

          {/* ── AUSÊNCIAS / SEM PONTO (NOVO) ── */}
          <Card style={{ marginBottom:16 }}>
            <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:16 }}>
              <SH style={{ marginBottom:0 }}>🚫 Médicos Ausentes — Sem Ponto Registrado</SH>
              <span style={{ fontSize:11, color:T.muted, marginLeft:'auto' }}>
                HR_ENTRADA vazia · {semPontoCount} ocorrências
              </span>
            </div>
            <AusenciasCard rows={filtered} cols={cols} />
          </Card>

          {/* ── MOTIVO DO IMPACTO NA ESPERA (NOVO) ── */}
          <Card style={{ marginBottom:16 }}>
            <SH>⏳ Motivo do Impacto na Espera — Unidades com Maior Tempo</SH>
            <ImpactoEsperaCard rows={filtered} cols={cols} />
          </Card>

          {/* ── Especialidades ── */}
          <Card style={{ marginBottom:16 }}>
            <SH>🩻 Especialidades com Mais Atendimentos</SH>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
              <div>{espBreak.slice(0,4).map((e,i)=>(
                <HBar key={e.nome} label={e.nome} value={e.total} max={maxEsp}
                  color={`hsl(${200+i*18},70%,55%)`} />
              ))}</div>
              <div>{espBreak.slice(4).map((e,i)=>(
                <HBar key={e.nome} label={e.nome} value={e.total} max={maxEsp}
                  color={`hsl(${272+i*18},70%,55%)`} />
              ))}</div>
            </div>
          </Card>

          {/* ── Footer ── */}
          <div style={{ textAlign:'center', color:T.muted, fontSize:11, paddingTop:8, paddingBottom:20 }}>
            Dashboard Monitoramento Hospitalar · {new Date().toLocaleDateString('pt-BR')}
            {períodoLabel && ` · ${períodoLabel}`}
            {horaFilt!=='TODAS' && ` · Hora ${horaFilt}h`}
          </div>
        </>)}
      </div>
    </div>
  )
}
