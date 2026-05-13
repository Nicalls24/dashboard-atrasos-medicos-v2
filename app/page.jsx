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

// ─── Status config ─────────────────────────────────────────
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
const parseEsperaMin = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24 * 60
  const s = String(v).trim()
  if (s.includes(':')) {
    const [h, m] = s.split(':').map(Number)
    return (h||0)*60 + (m||0)
  }
  return 0
}
const fmtMin = (m) => {
  const mm=Math.round(Math.abs(m))
  if (mm < 60) return `${mm}min`
  return `${Math.floor(mm/60)}h${mm%60>0?` ${mm%60}m`:''}`
}
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
  return () => true
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

// ─── Supabase config ──────────────────────────────────────
const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'
const PAGE_SIZE = 1000  // rows per page from hospital_dados

const sbFetch = (path, opts = {}) => {
  const { headers: extraHeaders, ...restOpts } = opts
  return fetch(`${SB_URL}/rest/v1/${path}`, {
    ...restOpts,
    headers: {
      'apikey':        SB_KEY,
      'Authorization': `Bearer ${SB_KEY}`,
      'Content-Type':  'application/json',
      'Range-Unit':    'items',
      ...extraHeaders,
    },
  })
}

// Extrai a data (YYYY-MM-DD) de uma linha
const rowDateStr = (r) => {
  const v = r['DATA_AGENDA'] ?? r[Object.keys(r).find(k => k.includes('DATA')) || ''] ?? ''
  return serialToDateStr(v)
}

// ─── Main ──────────────────────────────────────────────────
export default function Home() {
  const [dados,       setDados]       = useState([])
  const [uf,          setUf]          = useState('TODOS')
  const [status,      setStatus]      = useState('TODOS')
  const [loading,     setLoading]     = useState(false)
  const [storing,     setStoring]     = useState(false)
  const [storeStatus, setStoreStatus] = useState('')
  const [search,      setSearch]      = useState('')
  const [período,     setPeriodo]     = useState('TODOS')
  const [horaFilt,    setHoraFilt]    = useState('TODAS')
  const [dateFrom,    setDateFrom]    = useState('')
  const [dateTo,      setDateTo]      = useState('')
  const [timestamp,   setTimestamp]   = useState('')
  const [storageInfo, setStorageInfo] = useState({ dias: 0, total: 0 })
  const [initLoading, setInitLoading] = useState(true)

  // ── Carrega dados de hospital_dados ao iniciar ─────────
  useEffect(() => {
    const loadFromSupabase = async () => {
      try {
        setStoreStatus('Carregando dados…')

        // Carrega todas as linhas de hospital_dados com paginação
        // Cada linha tem campo "dados" (jsonb) = array de registros da planilha
        let allRows = []
        let offset = 0
        let lastVerifTs = ''

        while (true) {
          const res = await sbFetch(
            `hospital_dados?select=dados,verif_ts,snapshot_dates&order=id.asc&limit=${PAGE_SIZE}&offset=${offset}`
          )
          if (!res.ok) {
            const txt = await res.text()
            setStoreStatus(`Erro ${res.status}: ${txt.slice(0,120)}`)
            setInitLoading(false)
            return
          }
          const batch = await res.json()
          if (!Array.isArray(batch) || batch.length === 0) break

          for (const item of batch) {
            try {
              const d = item.dados
              if (!d) continue
              const parsed = typeof d === 'string' ? JSON.parse(d) : d
              if (Array.isArray(parsed))                   allRows = allRows.concat(parsed)
              else if (parsed && typeof parsed === 'object') allRows.push(parsed)
            } catch(e) {}
            if (item.verif_ts) lastVerifTs = item.verif_ts
          }

          setStoreStatus(`Carregando… ${allRows.length} registros`)
          if (batch.length < PAGE_SIZE) break
          offset += PAGE_SIZE
        }

        if (allRows.length > 0) {
          if (lastVerifTs) setTimestamp(lastVerifTs)
          const diasSet = new Set(allRows.map(rowDateStr).filter(Boolean))
          setDados(allRows)
          setStorageInfo({ dias: diasSet.size, total: allRows.length })
          setPeriodo('MES')
          setStoreStatus(`☁ ${diasSet.size} dias · ${allRows.length.toLocaleString('pt-BR')} registros`)
          setTimeout(() => setStoreStatus(''), 4000)
        } else {
          // Tabela vazia — sem erro, só sem dados ainda
          setStoreStatus('')
        }
      } catch (e) {
        console.error('load error:', e)
        setStoreStatus(`Erro de conexão: ${e.message}`)
      }
      setInitLoading(false)
    }

    loadFromSupabase()
  }, [])

  // ── Salva no Supabase (só hospital_dados) ────────────
  const saveToSupabase = useCallback(async (newRows, ts) => {
    setStoring(true)
    try {
      const BATCH = 500
      const newDates = [...new Set(newRows.map(rowDateStr).filter(Boolean))]

      // Apaga tudo e reinsere (mais simples e confiável)
      setStoreStatus('Limpando dados antigos…')
      await sbFetch('hospital_dados?id=gt.0', { method: 'DELETE' })

      // Insere em lotes — usa os nomes EXATOS das colunas do Supabase:
      // id (serial), uploaded_at (timestamptz), verif_ts (text),
      // dados (jsonb), snapshot_dates (text[])
      const totalChunks = Math.ceil(newRows.length / BATCH)
      for (let ci = 0; ci < totalChunks; ci++) {
        const slice = newRows.slice(ci * BATCH, (ci + 1) * BATCH)
        setStoreStatus(`Salvando lote ${ci + 1}/${totalChunks}…`)
        const res = await sbFetch('hospital_dados', {
          method: 'POST',
          headers: { 'Prefer': 'return=minimal' },
          body: JSON.stringify({
            verif_ts:       ts,
            dados:          slice,      // jsonb — array direto (sem stringify)
            snapshot_dates: newDates,   // text[] — nome real da coluna
          }),
        })
        if (!res.ok) {
          const txt = await res.text()
          setStoreStatus(`⚠ Erro lote ${ci+1}: ${txt.slice(0,120)}`)
          setStoring(false)
          return
        }
      }

      const diasSet = new Set(newDates)
      setStorageInfo({ dias: diasSet.size, total: newRows.length })
      setStoreStatus('✓ Salvo no Supabase')
      setTimeout(() => setStoreStatus(''), 3000)
    } catch (e) {
      console.error('save error:', e)
      setStoreStatus(`⚠ Erro: ${e.message}`)
    }
    setStoring(false)
  }, [])

  // ── Limpar dados ──────────────────────────────────────
  const clearStorage = async () => {
    if (!confirm('Apagar TODOS os dados do Supabase? Isso não pode ser desfeito.')) return
    setStoring(true)
    setStoreStatus('Apagando…')
    try {
      await sbFetch('hospital_dados?id=gt.0', { method: 'DELETE' })
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
    setPeriodo('MES'); setHoraFilt('TODAS'); setDateFrom(''); setDateTo('')
    setDados(json)
    setInitLoading(false)
    setLoading(false)
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

  // ── Pré-processa: dateStr + hora ──────────────────────
  const dadosComData = useMemo(() => {
    if (!dados.length) return []
    return dados.map(d => {
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

  const allDates = useMemo(() =>
    [...new Set(dadosComData.map(d=>d._dateStr).filter(Boolean))].sort(),
    [dadosComData])

  const periodoFn = useMemo(() => buildFilter(allDates, período, dateFrom, dateTo), [allDates, período, dateFrom, dateTo])

  const dadosPorPeriodo = useMemo(() =>
    dadosComData.filter(d => periodoFn(d._dateStr)),
    [dadosComData, periodoFn])

  const horasDisp = useMemo(() => {
    const hs = [...new Set(dadosPorPeriodo.map(d=>d._hora).filter(h=>h>=0&&h<=23))].sort((a,b)=>a-b)
    return hs
  }, [dadosPorPeriodo])

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

  // ── KPIs ──────────────────────────────────────────────
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

  const statusBreakdown = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s=String(d[cols.status]||'OK').trim(); m[s]=(m[s]||0)+1 })
    return m
  }, [filtered, cols])

  const topUnidadesAtraso = useMemo(() => {
    const m = {}
    filtered.filter(d=>String(d[cols.status]||'').toUpperCase().includes('ATRASO'))
      .forEach(d=>{ const u=d[cols.unidade]||'Sem Unidade'; m[u]=(m[u]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])

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

  const espBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d=>{ const e=d[cols.esp]||'Outro'; m[e]=(m[e]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxEsp = espBreak[0]?.total||1

  const ufBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d=>{ const u=String(d[cols.uf]||'?').trim(); m[u]=(m[u]||0)+1 })
    return Object.entries(m).map(([nome,total])=>({nome,total}))
      .sort((a,b)=>b.total-a.total).slice(0,8)
  }, [filtered, cols])
  const maxUF = ufBreak[0]?.total||1

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

  // Quantos dados tem no período selecionado (para decidir se mostra "vazio")
  const hasData    = dados.length > 0
  const isInit     = initLoading
  const hasFiltred = filtered.length > 0

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
              <PeriodoSelector value={período}
                onChange={p=>{setPeriodo(p);setHoraFilt('TODAS');setDateFrom('');setDateTo('')}}
                infoLabel={`${totalRegistros.toLocaleString('pt-BR')} reg.${períodoLabel ? ' · '+períodoLabel : ''}`}
                dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
                allDates={allDates} />
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
              <div style={{ color: storeStatus.startsWith('☁') || storeStatus.startsWith('✓') ? T.success : T.warning, fontWeight:600 }}>{storeStatus}</div>
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

        {/* ── Loading inicial ── */}
        {isInit && (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center',
            justifyContent:'center', minHeight:'calc(100vh - 140px)', gap:16 }}>
            <div style={{ fontSize:40 }}>⏳</div>
            <div style={{ fontSize:18, color:T.muted }}>Carregando dados do Supabase…</div>
            {storeStatus && (
              <div style={{ fontSize:13, color:T.warning, maxWidth:600, textAlign:'center',
                background:T.card, padding:'12px 20px', borderRadius:10,
                border:`1px solid ${T.warning}44` }}>{storeStatus}</div>
            )}
          </div>
        )}

        {/* ── Sem dados ── */}
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

        {/* ── Dashboard ── */}
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

          {/* ── Aviso se período sem dados ── */}
          {!hasFiltred && (
            <div style={{ background:T.card, border:`1px solid ${T.warning}44`, borderRadius:14,
              padding:'20px 24px', marginBottom:20, display:'flex', alignItems:'center', gap:14 }}>
              <span style={{ fontSize:28 }}>📅</span>
              <div>
                <div style={{ fontWeight:700, color:T.warning, marginBottom:4 }}>
                  Sem registros no período selecionado
                </div>
                <div style={{ fontSize:13, color:T.muted }}>
                  Os dados disponíveis vão de{' '}
                  <strong style={{color:T.sub}}>{allDates[0]?.split('-').reverse().join('/')}</strong>
                  {' '}até{' '}
                  <strong style={{color:T.sub}}>{allDates[allDates.length-1]?.split('-').reverse().join('/')}</strong>.
                  Selecione "Mês" ou "Ano" para ver todos os dados.
                </div>
              </div>
              <button onClick={()=>setPeriodo('MES')} style={{
                marginLeft:'auto', background:`linear-gradient(135deg,${T.accent},${T.accentB})`,
                color:'#000', fontWeight:700, fontSize:12, padding:'8px 16px',
                borderRadius:8, border:'none', cursor:'pointer', whiteSpace:'nowrap' }}>
                Ver Mês Completo
              </button>
            </div>
          )}

          {hasFiltred && (<>
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

            {/* ── Status ── */}
            <Card style={{ marginBottom:18 }}>
              <SH>📊 Distribuição de Status — Visão Operacional</SH>
              <StatusDashboard breakdown={statusBreakdown} total={totalRegistros} />
            </Card>

            {/* ── Unidades + UFs ── */}
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

            {/* ── Ausências ── */}
            <Card style={{ marginBottom:16 }}>
              <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:16 }}>
                <SH style={{ marginBottom:0 }}>🚫 Médicos Ausentes — Sem Ponto Registrado</SH>
                <span style={{ fontSize:11, color:T.muted, marginLeft:'auto' }}>
                  HR_ENTRADA vazia · {semPontoCount} ocorrências
                </span>
              </div>
              <AusenciasCard rows={filtered} cols={cols} />
            </Card>

            {/* ── Impacto na Espera ── */}
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
          </>)}

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
