'use client'

import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'

// ── THEME ─────────────────────────────────────────────────────────────────────
const C = {
  // backgrounds
  bg:       '#060C18',
  bgMid:    '#080F1E',
  bgCard:   '#0B1424',
  bgCard2:  '#0D1829',
  // accents
  teal:     '#00C9A7',
  tealDim:  '#009980',
  tealGlow: 'rgba(0,201,167,0.15)',
  blue:     '#3B82F6',
  blueDim:  '#2563EB',
  violet:   '#8B5CF6',
  amber:    '#F59E0B',
  rose:     '#F43F5E',
  emerald:  '#10B981',
  cyan:     '#06B6D4',
  // text
  text:     '#F1F5F9',
  sub:      '#94A3B8',
  muted:    '#475569',
  // borders
  border:   'rgba(0,201,167,0.12)',
  border2:  'rgba(0,201,167,0.25)',
  borderGray: 'rgba(148,163,184,0.08)',
}

const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'
const SBH = { 'apikey': SB_KEY, 'Authorization': `Bearer ${SB_KEY}`, 'Content-Type': 'application/json' }

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmtMin = m => {
  if (m === null || m === undefined) return '—'
  const abs = Math.round(Math.abs(m)), s = m < 0 ? '-' : ''
  if (abs < 60) return `${s}${abs}min`
  return `${s}${Math.floor(abs/60)}h${abs%60>0?` ${abs%60}m`:''}`
}
const fmtDate = s => {
  if (!s) return '—'
  const [y,mo,d] = String(s).split('-')
  return `${d}/${mo}/${y}`
}
const todayStr = () => {
  const n = new Date()
  return `${n.getUTCFullYear()}-${String(n.getUTCMonth()+1).padStart(2,'0')}-${String(n.getUTCDate()).padStart(2,'0')}`
}
const buildFilter = (allDates, periodo, df, dt) => {
  if (!allDates.length) return () => true
  const s = [...allDates].sort(), max = s[s.length-1]
  if (periodo==='HOJE')   return d => d === todayStr()
  if (periodo==='ONTEM')  { const r=new Date(max+'T00:00:00Z'); r.setUTCDate(r.getUTCDate()-1); const x=r.toISOString().slice(0,10); return d=>d===x }
  if (periodo==='SEMANA') { const r=new Date(max+'T00:00:00Z'); r.setUTCDate(r.getUTCDate()-6); const c=r.toISOString().slice(0,10); return d=>d>=c&&d<=max }
  if (periodo==='MES')    return d => d.slice(0,7)===max.slice(0,7)
  if (periodo==='ANO')    return d => d.slice(0,4)===max.slice(0,4)
  if (periodo==='PERIODO'){ const from=df||s[0],to=dt||max; return d=>d>=from&&d<=to }
  return () => true
}

const STATUS_CFG = {
  'OK':             { label:'OK',             color:'#10B981', glow:'rgba(16,185,129,0.2)' },
  'ATRASO':         { label:'Atraso 31–45min',color:'#F59E0B', glow:'rgba(245,158,11,0.2)' },
  'ATRASO CRÍTICO': { label:'Atraso Crítico', color:'#F97316', glow:'rgba(249,115,22,0.2)' },
  'ATRASO GRAVE':   { label:'Atraso Grave',   color:'#F43F5E', glow:'rgba(244,63,94,0.2)' },
  'Falta Médica':   { label:'Médico Faltou',  color:'#3B82F6', glow:'rgba(59,130,246,0.2)' },
  'SEM PONTO':      { label:'Sem Ponto',       color:'#64748B', glow:'rgba(100,116,139,0.2)' },
}
const getStatusCfg = s => {
  if (!s) return STATUS_CFG['OK']
  if (STATUS_CFG[s]) return STATUS_CFG[s]
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO')||u.includes('CRITICO')) return STATUS_CFG['ATRASO CRÍTICO']
  if (u.includes('GRAVE'))  return STATUS_CFG['ATRASO GRAVE']
  if (u.includes('ATRASO')) return STATUS_CFG['ATRASO']
  if (u.includes('FALTA'))  return STATUS_CFG['Falta Médica']
  return STATUS_CFG['OK']
}

const PERIODOS = [
  {key:'HOJE',label:'Hoje'},{key:'ONTEM',label:'Ontem'},{key:'SEMANA',label:'Semana'},
  {key:'MES',label:'Mês'},{key:'ANO',label:'Ano'},{key:'PERIODO',label:'Período'},
]

// ── ATOMS ─────────────────────────────────────────────────────────────────────
const glass = (color='rgba(0,201,167,0.08)') => ({
  background: color,
  backdropFilter: 'blur(16px)',
  WebkitBackdropFilter: 'blur(16px)',
})

function Panel({ children, accent=C.teal, style={} }) {
  return (
    <div style={{
      ...glass(),
      border: `1px solid ${accent}22`,
      borderRadius: 16,
      padding: 20,
      position: 'relative',
      overflow: 'hidden',
      ...style,
    }}>
      <div style={{ position:'absolute', top:0, left:0, right:0, height:'1px', background:`linear-gradient(90deg,transparent,${accent}60,transparent)` }} />
      {children}
    </div>
  )
}

function KpiCard({ icon, label, value, sub, color }) {
  return (
    <div style={{
      ...glass(`${color}08`),
      border: `1px solid ${color}20`,
      borderRadius: 14,
      padding: '18px 20px',
      position: 'relative',
      overflow: 'hidden',
      cursor: 'default',
    }}>
      <div style={{ position:'absolute', top:-30, right:-30, width:80, height:80, borderRadius:'50%', background:color, opacity:.07, filter:'blur(16px)' }} />
      <div style={{ position:'absolute', bottom:0, left:0, right:0, height:'2px', background:`linear-gradient(90deg,${color}80,transparent)` }} />
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12 }}>
        <div style={{ width:32, height:32, borderRadius:9, background:`${color}20`, border:`1px solid ${color}40`, display:'flex', alignItems:'center', justifyContent:'center', fontSize:16 }}>{icon}</div>
        <span style={{ fontSize:10, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.12em' }}>{label}</span>
      </div>
      <div style={{ fontSize:38, fontWeight:900, color:C.text, lineHeight:1, letterSpacing:'-1.5px', fontFamily:"'Syne',sans-serif" }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:C.sub, marginTop:8 }}>{sub}</div>}
    </div>
  )
}

function Bar({ label, value, max, color, unit='', rank, sub }) {
  const raw = typeof value === 'number' ? value : 0
  const pct = max > 0 ? Math.min(raw/max*100, 100) : 0
  return (
    <div style={{ marginBottom:12 }}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:5, gap:8 }}>
        <div style={{ display:'flex', alignItems:'flex-start', gap:6, minWidth:0 }}>
          {rank!=null && <span style={{ fontSize:10, fontWeight:800, color:rank===0?C.amber:C.muted, minWidth:18, flexShrink:0, marginTop:2 }}>#{rank+1}</span>}
          <div style={{ minWidth:0 }}>
            <div style={{ fontSize:12, color:C.text, fontWeight:500, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{label}</div>
            {sub && <div style={{ fontSize:10, color:C.muted, marginTop:1 }}>{sub}</div>}
          </div>
        </div>
        <span style={{ fontSize:12, fontWeight:700, color, flexShrink:0 }}>{value}{unit}</span>
      </div>
      <div style={{ background:'rgba(255,255,255,0.05)', borderRadius:4, height:5, overflow:'hidden' }}>
        <div style={{ height:'100%', borderRadius:4, background:color, width:`${pct}%`, transition:'width .7s ease' }} />
      </div>
    </div>
  )
}

function SecHead({ children, sub, right }) {
  return (
    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:16 }}>
      <div>
        <div style={{ fontSize:11, fontWeight:700, textTransform:'uppercase', letterSpacing:'.13em', color:C.sub }}>{children}</div>
        {sub && <div style={{ fontSize:10, color:C.muted, marginTop:2 }}>{sub}</div>}
      </div>
      {right}
    </div>
  )
}

function PeriodoBar({ value, onChange, allDates, dateFrom, dateTo, onDateFrom, onDateTo, total }) {
  return (
    <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
      <div style={{ display:'flex', background:'rgba(255,255,255,0.03)', border:`1px solid ${C.border}`, borderRadius:10, padding:3, gap:2 }}>
        {PERIODOS.map(p => (
          <button key={p.key} onClick={()=>onChange(p.key)} style={{
            padding:'5px 12px', borderRadius:7, border:'none', cursor:'pointer', fontSize:11, fontWeight:700, transition:'all .2s',
            background: value===p.key ? C.teal : 'transparent',
            color: value===p.key ? '#001a15' : C.muted,
          }}>{p.label}</button>
        ))}
      </div>
      {value==='PERIODO' && (
        <>
          <input type="date" value={dateFrom} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateFrom(e.target.value)}
            style={{ background:'rgba(255,255,255,0.04)', border:`1px solid ${C.border2}`, borderRadius:8, color:C.text, fontSize:11, padding:'5px 9px', outline:'none', colorScheme:'dark' }} />
          <span style={{ color:C.muted, fontSize:12 }}>→</span>
          <input type="date" value={dateTo} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateTo(e.target.value)}
            style={{ background:'rgba(255,255,255,0.04)', border:`1px solid ${C.border2}`, borderRadius:8, color:C.text, fontSize:11, padding:'5px 9px', outline:'none', colorScheme:'dark' }} />
        </>
      )}
      {total !== undefined && <span style={{ fontSize:11, color:C.muted }}>{total.toLocaleString('pt-BR')} registros</span>}
    </div>
  )
}

function SearchBar({ search, onSearch, uf, onUf, ufs, extra, showClear, onClear }) {
  return (
    <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center', marginBottom:18 }}>
      <div style={{ position:'relative', flex:1, minWidth:220 }}>
        <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', color:C.muted, fontSize:14 }}>🔍</span>
        <input value={search} onChange={e=>onSearch(e.target.value)} placeholder="Buscar unidade, médico, especialidade…"
          style={{ width:'100%', background:'rgba(255,255,255,0.04)', border:`1px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:12, padding:'8px 12px 8px 34px', outline:'none' }} />
      </div>
      <select value={uf} onChange={e=>onUf(e.target.value)}
        style={{ background:'rgba(6,12,24,0.95)', border:`1px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
        <option value="TODOS">Todos os Estados</option>
        {ufs.map(u=><option key={u}>{u}</option>)}
      </select>
      {extra}
      {showClear && (
        <button onClick={onClear} style={{ background:'rgba(244,63,94,0.08)', border:`1px solid rgba(244,63,94,0.25)`, borderRadius:10, color:C.rose, fontSize:12, padding:'8px 12px', cursor:'pointer' }}>✕ Limpar</button>
      )}
    </div>
  )
}

function MiniChart({ data, color, height=56 }) {
  if (!data.length) return null
  const max = Math.max(...data.map(d=>d.v), 1)
  return (
    <div style={{ display:'flex', alignItems:'flex-end', gap:2, height }}>
      {data.map((d,i) => (
        <div key={i} title={`${fmtDate(d.k)}: ${d.v}`} style={{
          flex:1, borderRadius:'2px 2px 0 0',
          background:`linear-gradient(0deg,${color},${color}55)`,
          height:`${Math.max((d.v/max)*100, 4)}%`,
          opacity: .4+.6*(d.v/max),
          transition:'height .5s ease',
        }} />
      ))}
    </div>
  )
}

// ── SIDEBAR ───────────────────────────────────────────────────────────────────
function Sidebar({ tab, setTab, storageInfo, storeMsg, timestamp, storing, loadingDB, onUpload, onClear }) {
  const total = storageInfo.agendas + storageInfo.espera
  return (
    <div style={{
      width: 260,
      flexShrink: 0,
      background: C.bgCard,
      borderRight: `1px solid ${C.border}`,
      display: 'flex',
      flexDirection: 'column',
      position: 'fixed',
      top: 0,
      left: 0,
      bottom: 0,
      zIndex: 50,
      overflow: 'hidden',
    }}>
      {/* Top glow */}
      <div style={{ position:'absolute', top:-60, left:'50%', transform:'translateX(-50%)', width:200, height:120, background:C.teal, opacity:.07, borderRadius:'50%', filter:'blur(30px)', pointerEvents:'none' }} />

      {/* Logo */}
      <div style={{ padding:'28px 24px 24px', borderBottom:`1px solid ${C.border}`, textAlign:'center' }}>
        <div style={{ fontSize:22, fontWeight:900, color:C.text, fontFamily:"'Syne',sans-serif", lineHeight:1.2, marginBottom:10 }}>
          Monitor Clínicas
        </div>
        <div style={{ fontSize:11, fontWeight:700, color:C.teal, textTransform:'uppercase', letterSpacing:'.13em', marginBottom:5 }}>
          Núcleo Observação e Controle
        </div>
        <div style={{ fontSize:10, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em' }}>
          Clínicas - Esperas e Agenda
        </div>
      </div>

      {/* NAV */}
      <div style={{ padding:'20px 16px', flex:1 }}>
        <div style={{ fontSize:9.5, color:C.muted, textTransform:'uppercase', letterSpacing:'.12em', marginBottom:10, paddingLeft:8 }}>Painéis</div>
        {[
          { key:'agendas', label:'Agendas Médicas', count:storageInfo.agendas },
          { key:'espera',  label:'Fila de Espera',  count:storageInfo.espera  },
        ].map(item => {
          const active = tab === item.key
          return (
            <button key={item.key} onClick={()=>setTab(item.key)} style={{
              width:'100%', display:'flex', alignItems:'center', gap:10,
              padding:'10px 12px', borderRadius:11, border:'none',
              background: active ? `${C.teal}15` : 'transparent',
              borderLeft: active ? `3px solid ${C.teal}` : '3px solid transparent',
              cursor:'pointer', marginBottom:4, transition:'all .2s', textAlign:'left',
            }}>
              <div style={{ width:28, height:28, borderRadius:8, background: active?`${C.teal}20`:'rgba(255,255,255,0.04)', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0 }}>
                {item.key==='agendas' ? (
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke={active?C.teal:C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
                  </svg>
                ) : (
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke={active?C.teal:C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
                  </svg>
                )}
              </div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:13, fontWeight:700, color: active ? C.teal : C.sub }}>{item.label}</div>
              </div>
              {item.count > 0 && (
                <span style={{ fontSize:10, fontWeight:700, background: active?`${C.teal}25`:'rgba(255,255,255,0.06)', color: active?C.teal:C.muted, borderRadius:8, padding:'2px 7px' }}>
                  {item.count.toLocaleString('pt-BR')}
                </span>
              )}
            </button>
          )
        })}
      </div>

      {/* Status + upload */}
      <div style={{ padding:'16px', borderTop:`1px solid ${C.border}` }}>
        {/* Status */}
        <div style={{ background:'rgba(255,255,255,0.03)', border:`1px solid ${C.border}`, borderRadius:11, padding:'12px 14px', marginBottom:12 }}>
          <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:6 }}>
            <div style={{ width:7, height:7, borderRadius:'50%', background: storeMsg&&!storeMsg.startsWith('☁')&&!storeMsg.startsWith('✓') ? C.amber : C.teal, boxShadow:`0 0 6px ${storeMsg&&!storeMsg.startsWith('☁')&&!storeMsg.startsWith('✓') ? C.amber : C.teal}` }} />
            <span style={{ fontSize:10, fontWeight:700, color:C.sub, textTransform:'uppercase', letterSpacing:'.08em' }}>Status</span>
          </div>
          {storeMsg ? (
            <div style={{ fontSize:11, color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?C.teal:C.amber, fontWeight:600, lineHeight:1.4 }}>{storeMsg}</div>
          ) : total > 0 ? (
            <>
              <div style={{ fontSize:11, color:C.teal, fontWeight:700 }}>☁ {total.toLocaleString('pt-BR')} registros</div>
              {timestamp && <div style={{ fontSize:10, color:C.muted, marginTop:3 }}>{timestamp}</div>}
            </>
          ) : (
            <div style={{ fontSize:11, color:C.muted }}>Nenhum dado carregado</div>
          )}
        </div>

        {/* Upload */}
        <label style={{
          display:'block', width:'100%', textAlign:'center',
          background: storing ? 'rgba(255,255,255,0.05)' : `linear-gradient(135deg,${C.teal},${C.tealDim})`,
          color: storing ? C.muted : '#001a15',
          fontWeight:800, fontSize:13, padding:'11px',
          borderRadius:11, cursor: storing?'default':'pointer', transition:'all .25s',
          boxShadow: storing ? 'none' : `0 4px 24px ${C.tealGlow}`,
          fontFamily:"'Syne',sans-serif",
        }}>
          {storing ? 'Salvando…' : '+ Carregar Planilha'}
          <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={onUpload} disabled={storing||loadingDB} />
        </label>

        {total > 0 && !storing && (
          <button onClick={onClear} style={{ width:'100%', marginTop:8, background:'transparent', border:`1px solid rgba(244,63,94,0.2)`, borderRadius:11, color:C.rose, fontSize:11, padding:'7px', cursor:'pointer', opacity:.7 }}>
            🗑 Limpar banco de dados
          </button>
        )}
      </div>
    </div>
  )
}

// ── TAB AGENDAS ───────────────────────────────────────────────────────────────
function TabAgendas({ rows }) {
  const [periodo,    setPeriodo]    = useState('MES')
  const [dateFrom,   setDateFrom]   = useState('')
  const [dateTo,     setDateTo]     = useState('')
  const [ufFilt,     setUfFilt]     = useState('TODOS')
  const [statusFilt, setStatusFilt] = useState('TODOS')
  const [search,     setSearch]     = useState('')

  const allDates = useMemo(() => [...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(), [rows])
  const periodoFn = useMemo(() => buildFilter(allDates, periodo, dateFrom, dateTo), [allDates, periodo, dateFrom, dateTo])
  const ufs       = useMemo(() => [...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(), [rows])
  const statuses  = useMemo(() => [...new Set(rows.map(r=>r.status).filter(Boolean))].sort(), [rows])

  const filtered = useMemo(() => {
    let r = rows.filter(d=>periodoFn(d.data_agenda))
    if (ufFilt!=='TODOS')     r=r.filter(d=>d.uf===ufFilt)
    if (statusFilt!=='TODOS') r=r.filter(d=>d.status===statusFilt)
    if (search) { const q=search.toLowerCase(); r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))) }
    return r
  }, [rows, periodoFn, ufFilt, statusFilt, search])

  const { total, emAtraso, semPonto, topUnidades, topMedicos, statusBreak, byDate, atrasoPct, semPontoPct } = useMemo(() => {
    const total    = filtered.length
    const emAtraso = filtered.filter(d=>String(d.status||'').toUpperCase().includes('ATRASO')).length
    const semPonto = filtered.filter(d=>d.hr_entrada_min===null||d.hr_entrada_min===undefined).length
    const uMap={},mMap={},sMap={}
    filtered.forEach(d=>{
      const u=d.nm_local||'?'; uMap[u]=(uMap[u]||0)+1
      const m=d.nm_medico||''; if(m) mMap[m]=(mMap[m]||0)+1
      const s=d.status||'OK'; sMap[s]=(sMap[s]||0)+1
    })
    const topUnidades = Object.entries(uMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const topMedicos  = Object.entries(mMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const statusBreak = Object.entries(sMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)
    const dMap={}; filtered.forEach(d=>{ if(d.data_agenda) dMap[d.data_agenda]=(dMap[d.data_agenda]||0)+1 })
    const byDate = Object.entries(dMap).map(([k,v])=>({k,v})).sort((a,b)=>a.k.localeCompare(b.k)).slice(-28)
    return { total, emAtraso, semPonto, topUnidades, topMedicos, statusBreak, byDate,
      atrasoPct: total>0?((emAtraso/total)*100).toFixed(1):'0',
      semPontoPct: total>0?((semPonto/total)*100).toFixed(1):'0' }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:14 }}>
      <div style={{ fontSize:52 }}>📋</div>
      <div style={{ fontSize:20, fontWeight:800, color:C.text, fontFamily:"'Syne',sans-serif" }}>Nenhuma agenda carregada</div>
      <div style={{ fontSize:13, color:C.muted }}>Use o menu lateral para carregar uma planilha</div>
    </div>
  )

  const maxU=topUnidades[0]?.v||1, maxM=topMedicos[0]?.v||1, maxS=statusBreak[0]?.v||1

  return (
    <div>
      {/* Periodo */}
      <div style={{ marginBottom:18 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo} total={total} />
      </div>

      {/* Filters */}
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||statusFilt!=='TODOS'||!!search}
        onClear={()=>{setUfFilt('TODOS');setStatusFilt('TODOS');setSearch('')}}
        extra={<select value={statusFilt} onChange={e=>setStatusFilt(e.target.value)}
          style={{ background:'rgba(6,12,24,0.95)', border:`1px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
          <option value="TODOS">Todos os Status</option>
          {statuses.map(s=><option key={s}>{s}</option>)}
        </select>} />

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12, marginBottom:18 }}>
        <KpiCard icon="🗓" label="Total Agendas"    value={total.toLocaleString('pt-BR')}          color={C.teal}    />
        <KpiCard icon="⚠️" label="Em Atraso"        value={emAtraso.toLocaleString('pt-BR')}       sub={`${atrasoPct}% do total`}    color={C.rose}    />
        <KpiCard icon="🚫" label="Sem Ponto"        value={semPonto.toLocaleString('pt-BR')}       sub={`${semPontoPct}% do total`}  color={C.violet}  />
        <KpiCard icon="✅" label="Com Atendimento"  value={(total-semPonto).toLocaleString('pt-BR')} sub={`${(100-Number(semPontoPct)).toFixed(1)}% presentes`} color={C.emerald} />
      </div>

      {/* Charts row */}
      <div style={{ display:'grid', gridTemplateColumns:'1.6fr 1fr', gap:14, marginBottom:14 }}>
        {byDate.length>1 && (
          <Panel accent={C.teal}>
            <SecHead sub={`Últimos ${byDate.length} dias com dados`}>📈 Agendas por Dia</SecHead>
            <MiniChart data={byDate} color={C.teal} height={60} />
            <div style={{ display:'flex', justifyContent:'space-between', marginTop:8 }}>
              <span style={{ fontSize:10, color:C.muted }}>{fmtDate(byDate[0]?.k)}</span>
              <span style={{ fontSize:10, color:C.muted }}>{fmtDate(byDate[byDate.length-1]?.k)}</span>
            </div>
          </Panel>
        )}
        <Panel accent={C.violet}>
          <SecHead>📊 Por Status</SecHead>
          {statusBreak.slice(0,6).map(({k,v})=>{
            const cfg=getStatusCfg(k)
            const pct=total>0?((v/total)*100).toFixed(1):'0'
            return (
              <div key={k} style={{ display:'flex', alignItems:'center', gap:8, marginBottom:8 }}>
                <div style={{ width:8,height:8,borderRadius:'50%',background:cfg.color,flexShrink:0 }} />
                <div style={{ flex:1, fontSize:11, color:C.sub, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{cfg.label}</div>
                <span style={{ fontSize:11, fontWeight:700, color:cfg.color, minWidth:32, textAlign:'right' }}>{v.toLocaleString('pt-BR')}</span>
                <div style={{ width:50, background:'rgba(255,255,255,0.05)', borderRadius:3, height:3, overflow:'hidden', flexShrink:0 }}>
                  <div style={{ height:'100%', background:cfg.color, width:`${(v/maxS)*100}%` }} />
                </div>
                <span style={{ fontSize:10, color:C.muted, minWidth:28, textAlign:'right' }}>{pct}%</span>
              </div>
            )
          })}
        </Panel>
      </div>

      {/* Rankings */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:14 }}>
        <Panel accent={C.cyan}>
          <SecHead>🏥 Top Unidades</SecHead>
          {topUnidades.map((u,i)=><Bar key={u.n} rank={i} label={u.n} value={u.v} max={maxU} color={i===0?C.amber:i<3?C.cyan:C.teal} />)}
        </Panel>
        <Panel accent={C.violet}>
          <SecHead>👨‍⚕️ Top Médicos</SecHead>
          {topMedicos.map((m,i)=><Bar key={m.n} rank={i} label={m.n} value={m.v} max={maxM} color={i===0?C.amber:i<3?C.violet:C.blue} />)}
        </Panel>
      </div>

      {/* Table */}
      <Panel accent={C.borderGray}>
        <SecHead sub="Últimos 50 registros filtrados">📋 Detalhamento</SecHead>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11.5 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${C.border}` }}>
                {['Data','Unidade','Médico','Especialidade','UF','Status','Atraso'].map(h=>(
                  <th key={h} style={{ padding:'7px 10px', textAlign:'left', color:C.muted, fontWeight:700, fontSize:9.5, textTransform:'uppercase', letterSpacing:'.08em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0,50).map((r,i)=>{
                const cfg=getStatusCfg(r.status)
                return (
                  <tr key={i} style={{ borderBottom:`1px solid ${C.borderGray}`, transition:'background .12s' }}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(0,201,167,0.04)'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <td style={{ padding:'8px 10px', color:C.sub, whiteSpace:'nowrap' }}>{fmtDate(r.data_agenda)}</td>
                    <td style={{ padding:'8px 10px', color:C.text, fontWeight:600, maxWidth:180, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_local}</td>
                    <td style={{ padding:'8px 10px', color:C.sub, whiteSpace:'nowrap' }}>{r.nm_medico}</td>
                    <td style={{ padding:'8px 10px', color:C.muted, maxWidth:140, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.ds_especialidade}</td>
                    <td style={{ padding:'8px 10px', color:C.muted }}>{r.uf}</td>
                    <td style={{ padding:'8px 10px' }}>
                      <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:cfg.glow, color:cfg.color, border:`1px solid ${cfg.color}30` }}>{r.status||'—'}</span>
                    </td>
                    <td style={{ padding:'8px 10px', color:r.atraso==='SIM'?C.rose:C.emerald, fontWeight:700, fontSize:11 }}>{r.atraso||'—'}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </Panel>
    </div>
  )
}

// ── RÉGUAS DE ESPERA ──────────────────────────────────────────────────────────
// < 15min    → ignorar
// 15–30min   → Espera Moderada  (amber)
// 31–89min   → Espera Grave     (orange)
// >= 90min   → Espera Crítica   (rose)
const classEspera = m => {
  if (m === null || m === undefined || m < 15) return null
  if (m <= 30) return { label:'Espera Moderada', color:'#F59E0B', bg:'rgba(245,158,11,0.12)', border:'rgba(245,158,11,0.3)' }
  if (m <= 89) return { label:'Espera Grave',    color:'#F97316', bg:'rgba(249,115,22,0.12)', border:'rgba(249,115,22,0.3)' }
  return             { label:'Espera Crítica',  color:'#F43F5E', bg:'rgba(244,63,94,0.12)',  border:'rgba(244,63,94,0.3)'  }
}

// ── TAB ESPERA ────────────────────────────────────────────────────────────────
function TabEspera({ rows }) {
  const [periodo,  setPeriodo]  = useState('MES')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo,   setDateTo]   = useState('')
  const [ufFilt,   setUfFilt]   = useState('TODOS')
  const [search,   setSearch]   = useState('')

  const allDates = useMemo(() => [...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(), [rows])
  const periodoFn = useMemo(() => buildFilter(allDates, periodo, dateFrom, dateTo), [allDates, periodo, dateFrom, dateTo])
  const ufs = useMemo(() => [...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(), [rows])

  const filtered = useMemo(() => {
    let r = rows.filter(d=>periodoFn(d.data_agenda))
    if (ufFilt!=='TODOS') r=r.filter(d=>d.uf===ufFilt)
    if (search) { const q=search.toLowerCase(); r=r.filter(d=>[d.nm_local,d.nm_medico,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))) }
    return r
  }, [rows, periodoFn, ufFilt, search])

  const stats = useMemo(() => {
    // Só considera esperas >= 15min
    const comEspera = filtered.filter(d => d.tempo_espera_min !== null && d.tempo_espera_min !== undefined && d.tempo_espera_min >= 15)
    const total     = filtered.length
    const totalPac  = filtered.reduce((a,d)=>a+(d.qt_pacientes_aguardando||0), 0)

    // KPIs de classificação
    const moderada = comEspera.filter(d=>d.tempo_espera_min>=15&&d.tempo_espera_min<=30).length  // 15-30min Moderada
    const grave    = comEspera.filter(d=>d.tempo_espera_min>30&&d.tempo_espera_min<=89).length   // 31-89min Grave
    const critico  = comEspera.filter(d=>d.tempo_espera_min>=90).length                          // >=90min Critico

    // Esperas por hora do dia (HR_REGISTRO_ESPERA em minutos desde meia-noite)
    const horaMap = {}
    filtered.forEach(d => {
      const h = d.hr_registro_espera_min
      if (h === null || h === undefined) return
      const hora = Math.floor(h / 60)
      if (hora < 0 || hora > 23) return
      const key = String(hora).padStart(2,'0') + 'h'
      if (!horaMap[key]) horaMap[key] = { hora, totalPac:0, esperas:0, criticos:0, graves:0, moderadas:0 }
      horaMap[key].totalPac += d.qt_pacientes_aguardando || 0
      horaMap[key].esperas  += 1
      const mh = d.tempo_espera_min
      if (mh >= 90)       horaMap[key].criticos  += 1
      else if (mh >= 31)  horaMap[key].graves    += 1
      else if (mh >= 15)  horaMap[key].moderadas += 1
    })
    const porHora = Object.values(horaMap).sort((a,b)=>a.hora-b.hora)

    // Top unidades por espera crítica
    const unidMap = {}
    comEspera.forEach(d => {
      const u = d.nm_local||'?'
      if (!unidMap[u]) unidMap[u] = { n:u, criticos:0, total:0, somaPac:0 }
      unidMap[u].total   += 1
      unidMap[u].somaPac += d.qt_pacientes_aguardando||0
      if (d.tempo_espera_min >= 90) unidMap[u].criticos += 1
    })
    const topUnidades = Object.values(unidMap).sort((a,b)=>b.criticos-a.criticos).slice(0,8)

    // Análise falta/atraso — usando colunas atraso e status da tabela espera
    // atraso = 'SIM' + tempo_atraso > 31min → atraso real
    // atraso = 'FALTA' → falta médica
    const faltas   = filtered.filter(d => String(d.atraso||'').toUpperCase() === 'FALTA')
    const atrasosReais = filtered.filter(d => {
      if (String(d.atraso||'').toUpperCase() !== 'SIM') return false
      // tempo_atraso em minutos (já convertido no save)
      const t = d.tempo_atraso_min
      return t !== null && t !== undefined && Math.abs(t) > 31
    })

    // Por status de atraso (coluna U / status)
    const statusAtrasoMap = {}
    atrasosReais.forEach(d => {
      const s = d.status || 'Sem Status'
      statusAtrasoMap[s] = (statusAtrasoMap[s]||0)+1
    })
    const statusAtraso = Object.entries(statusAtrasoMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)

    // Tendência diária (esperas >= 15min por dia)
    const dMap = {}
    comEspera.forEach(d => {
      if (!d.data_agenda) return
      if (!dMap[d.data_agenda]) dMap[d.data_agenda] = 0
      dMap[d.data_agenda] += 1
    })
    const byDate = Object.entries(dMap).map(([k,v])=>({k,v})).sort((a,b)=>a.k.localeCompare(b.k)).slice(-28)

    return { total, totalPac, moderada, grave, critico, porHora, topUnidades, faltas, atrasosReais, statusAtraso, byDate, comEspera }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:14 }}>
      <div style={{ fontSize:52 }}>⏱️</div>
      <div style={{ fontSize:20, fontWeight:800, color:C.text, fontFamily:"'Syne',sans-serif" }}>Nenhum dado de espera</div>
      <div style={{ fontSize:13, color:C.muted }}>Use o menu lateral para carregar uma planilha</div>
    </div>
  )

  const { total, totalPac, moderada, grave, critico, porHora, topUnidades, faltas, atrasosReais, statusAtraso, byDate } = stats
  const maxEsperas = Math.max(...porHora.map(h=>h.esperas), 1)
  const maxPac     = Math.max(...porHora.map(h=>h.totalPac), 1)

  return (
    <div>
      <div style={{ marginBottom:18 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo} total={total} />
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}} />

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12, marginBottom:18 }}>
        <KpiCard icon="👥" label="Pacientes na Fila"        value={totalPac.toLocaleString('pt-BR')} color={C.teal}   />
        <KpiCard icon="🟡" label="Espera Moderada 15–30min" value={moderada.toLocaleString('pt-BR')} color={C.amber}  />
        <KpiCard icon="🟠" label="Espera Grave 31min–1h29"  value={grave.toLocaleString('pt-BR')}    color="#F97316"  />
        <KpiCard icon="🔴" label="Espera Crítica +1h30"     value={critico.toLocaleString('pt-BR')}  color={C.rose}   />
      </div>

      {/* GRÁFICO HORA A HORA */}
      <Panel accent={C.teal} style={{ marginBottom:14 }}>
        <SecHead sub="Esperas e pacientes aguardando por hora · HR_REGISTRO_ESPERA">🕐 Esperas por Hora do Dia</SecHead>

        {porHora.length === 0 && <div style={{ color:C.muted, fontSize:12 }}>Sem dados de hora de registro no período.</div>}

        {porHora.length > 0 && (<>
          {/* Barras verticais */}
          <div style={{ display:'flex', alignItems:'flex-end', gap:3, height:140, marginBottom:6, paddingTop:8 }}>
            {porHora.map(h => {
              const hPct  = (h.esperas / maxEsperas) * 100
              const color = h.criticos > 0 ? C.rose : h.graves > 0 ? '#F97316' : h.moderadas > 0 ? C.amber : C.teal
              return (
                <div key={h.hora} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', height:'100%', justifyContent:'flex-end', gap:3, cursor:'default' }}
                  title={`${String(h.hora).padStart(2,'0')}h — ${h.esperas} esperas · ${h.totalPac} pac. · ${h.criticos} crít.`}>
                  {h.esperas > 0 && (
                    <div style={{ fontSize:8, color, fontWeight:700, lineHeight:1 }}>{h.esperas}</div>
                  )}
                  <div style={{
                    width:'100%', borderRadius:'3px 3px 0 0',
                    background: h.criticos > 0
                      ? 'linear-gradient(0deg,#F43F5E,#F97316)'
                      : h.graves > 0
                      ? 'linear-gradient(0deg,#F97316,#FBBF24)'
                      : h.moderadas > 0
                      ? 'linear-gradient(0deg,#F59E0B,#FCD34D)'
                      : 'rgba(255,255,255,0.1)',
                    height: h.esperas > 0 ? `${Math.max(hPct, 4)}%` : '2px',
                    minHeight: h.esperas > 0 ? 4 : 2,
                    transition:'height .6s ease',
                  }} />
                </div>
              )
            })}
          </div>
          {/* Labels hora */}
          <div style={{ display:'flex', gap:3, marginBottom:16 }}>
            {porHora.map(h => (
              <div key={h.hora} style={{ flex:1, textAlign:'center', fontSize:8.5, color:C.muted, fontWeight:500 }}>
                {String(h.hora).padStart(2,'0')}h
              </div>
            ))}
          </div>
          {/* Legenda */}
          <div style={{ display:'flex', gap:16, marginBottom:16, flexWrap:'wrap' }}>
            {[
              { color:C.amber,   label:'Moderada 15–30min' },
              { color:'#F97316', label:'Grave 31min–1h29'  },
              { color:C.rose,    label:'Crítica +1h30'     },
            ].map(l => (
              <div key={l.label} style={{ display:'flex', alignItems:'center', gap:5 }}>
                <div style={{ width:10, height:10, borderRadius:2, background:l.color }} />
                <span style={{ fontSize:10, color:C.muted }}>{l.label}</span>
              </div>
            ))}
          </div>

          {/* Tabela hora a hora */}
          <div style={{ overflowX:'auto' }}>
            <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11.5 }}>
              <thead>
                <tr style={{ borderBottom:`1px solid ${C.border}` }}>
                  {['Hora','Total Esperas','Pac. Aguardando','🟡 Moderada','🟠 Grave','🔴 Crítica'].map(h=>(
                    <th key={h} style={{ padding:'8px 14px', textAlign:h==='Hora'?'left':'center', color:C.muted, fontWeight:700, fontSize:9.5, textTransform:'uppercase', letterSpacing:'.08em', whiteSpace:'nowrap' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {porHora.map(h => {
                  const hasCrit = h.criticos > 0
                  return (
                    <tr key={h.hora}
                      style={{ borderBottom:`1px solid ${C.borderGray}`, background: hasCrit ? 'rgba(244,63,94,0.04)' : 'transparent', transition:'background .12s' }}
                      onMouseEnter={e=>e.currentTarget.style.background='rgba(0,201,167,0.06)'}
                      onMouseLeave={e=>e.currentTarget.style.background=hasCrit?'rgba(244,63,94,0.04)':'transparent'}>
                      <td style={{ padding:'10px 14px', fontWeight:800, color:C.text, fontFamily:'monospace', fontSize:14 }}>
                        {String(h.hora).padStart(2,'0')}:00
                      </td>
                      <td style={{ padding:'10px 14px', textAlign:'center', color:C.teal, fontWeight:700, fontSize:13 }}>{h.esperas}</td>
                      <td style={{ padding:'10px 14px', textAlign:'center', color:C.cyan, fontWeight:700, fontSize:13 }}>{h.totalPac.toLocaleString('pt-BR')}</td>
                      <td style={{ padding:'10px 14px', textAlign:'center' }}>
                        {h.moderadas > 0
                          ? <span style={{ fontSize:12, fontWeight:700, padding:'3px 12px', borderRadius:20, background:'rgba(245,158,11,0.15)', color:'#F59E0B', border:'1px solid rgba(245,158,11,0.3)' }}>{h.moderadas}</span>
                          : <span style={{ color:C.muted, fontSize:12 }}>—</span>}
                      </td>
                      <td style={{ padding:'10px 14px', textAlign:'center' }}>
                        {h.graves > 0
                          ? <span style={{ fontSize:12, fontWeight:700, padding:'3px 12px', borderRadius:20, background:'rgba(249,115,22,0.15)', color:'#F97316', border:'1px solid rgba(249,115,22,0.3)' }}>{h.graves}</span>
                          : <span style={{ color:C.muted, fontSize:12 }}>—</span>}
                      </td>
                      <td style={{ padding:'10px 14px', textAlign:'center' }}>
                        {h.criticos > 0
                          ? <span style={{ fontSize:12, fontWeight:700, padding:'3px 12px', borderRadius:20, background:'rgba(244,63,94,0.15)', color:C.rose, border:'1px solid rgba(244,63,94,0.3)' }}>{h.criticos}</span>
                          : <span style={{ color:C.muted, fontSize:12 }}>—</span>}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </>)}
      </Panel>

      {/* MÉDICOS — complemento */}
      <Panel accent={C.amber} style={{ marginBottom:14 }}>
        <SecHead sub="Complemento — médicos com falta ou atraso no período">⚠️ Médicos — Falta e Atraso</SecHead>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:12, marginBottom: statusAtraso.length > 0 ? 16 : 0 }}>
          {[
            { label:'Faltas',        value:faltas.length,         color:C.rose  },
            { label:'Atrasos >31min',value:atrasosReais.length,   color:C.amber },
            { label:'Total',         value:faltas.length+atrasosReais.length, color:C.blue  },
          ].map(k => (
            <div key={k.label} style={{ background:`${k.color}0A`, border:`1px solid ${k.color}25`, borderRadius:12, padding:'14px 16px', textAlign:'center' }}>
              <div style={{ fontSize:9, fontWeight:700, color:k.color, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:8 }}>{k.label}</div>
              <div style={{ fontSize:30, fontWeight:900, color:k.color, fontFamily:"'Syne',sans-serif" }}>{k.value}</div>
            </div>
          ))}
        </div>
        {statusAtraso.length > 0 && (
          <div>
            <div style={{ fontSize:10, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:10 }}>Classificação do Atraso</div>
            {statusAtraso.map(({k,v}) => {
              const cfg = getStatusCfg(k)
              return (
                <div key={k} style={{ display:'flex', alignItems:'center', gap:10, marginBottom:8 }}>
                  <span style={{ fontSize:10, fontWeight:700, padding:'2px 10px', borderRadius:20, background:cfg.glow, color:cfg.color, border:`1px solid ${cfg.color}30`, whiteSpace:'nowrap', minWidth:130 }}>{cfg.label}</span>
                  <div style={{ flex:1, background:'rgba(255,255,255,0.05)', borderRadius:3, height:5, overflow:'hidden' }}>
                    <div style={{ height:'100%', background:cfg.color, width:`${(v/(statusAtraso[0]?.v||1))*100}%` }} />
                  </div>
                  <span style={{ fontSize:12, fontWeight:700, color:cfg.color, minWidth:28, textAlign:'right' }}>{v}</span>
                </div>
              )
            })}
          </div>
        )}
        {!statusAtraso.length && faltas.length===0 && atrasosReais.length===0 && (
          <div style={{ color:C.muted, fontSize:12 }}>Nenhuma falta ou atraso identificado no período.</div>
        )}
      </Panel>

      {/* Tabela detalhada */}
      <Panel accent={C.borderGray}>
        <SecHead sub="Esperas ≥15min · ordenado por maior tempo · top 100">📋 Detalhamento de Espera</SecHead>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${C.border}` }}>
                {['Data','Hora','Unidade','Médico','UF','Pac.','Tempo','Classificação','Médico'].map(h=>(
                  <th key={h} style={{ padding:'7px 10px', textAlign:'left', color:C.muted, fontWeight:700, fontSize:9, textTransform:'uppercase', letterSpacing:'.07em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...filtered]
                .filter(d=>d.tempo_espera_min!==null&&d.tempo_espera_min!==undefined&&d.tempo_espera_min>=15)
                .sort((a,b)=>(b.tempo_espera_min||0)-(a.tempo_espera_min||0))
                .slice(0,100)
                .map((r,i) => {
                  const cls = classEspera(r.tempo_espera_min)
                  const horaReg = r.hr_registro_espera_min!==null&&r.hr_registro_espera_min!==undefined
                    ? `${String(Math.floor(r.hr_registro_espera_min/60)).padStart(2,'0')}:${String(r.hr_registro_espera_min%60).padStart(2,'0')}`
                    : '—'
                  const isFalta   = String(r.atraso||'').toUpperCase()==='FALTA'
                  const isAtraso  = String(r.atraso||'').toUpperCase()==='SIM'
                  const tempoAtr  = r.tempo_atraso_min
                  const atrasoReal= isAtraso&&tempoAtr!==null&&tempoAtr!==undefined&&Math.abs(tempoAtr)>31
                  return (
                    <tr key={i} style={{ borderBottom:`1px solid ${C.borderGray}`, transition:'background .12s' }}
                      onMouseEnter={e=>e.currentTarget.style.background='rgba(0,201,167,0.04)'}
                      onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                      <td style={{ padding:'8px 10px', color:C.sub, whiteSpace:'nowrap' }}>{fmtDate(r.data_agenda)}</td>
                      <td style={{ padding:'8px 10px', color:C.sub, fontFamily:'monospace' }}>{horaReg}</td>
                      <td style={{ padding:'8px 10px', color:C.text, fontWeight:600, maxWidth:160, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_local}</td>
                      <td style={{ padding:'8px 10px', color:C.sub, maxWidth:130, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_medico}</td>
                      <td style={{ padding:'8px 10px', color:C.muted }}>{r.uf}</td>
                      <td style={{ padding:'8px 10px', color:C.cyan, fontWeight:700, textAlign:'center' }}>{r.qt_pacientes_aguardando??'—'}</td>
                      <td style={{ padding:'8px 10px' }}>
                        <span style={{ fontSize:12, fontWeight:900, color:cls?.color||C.muted }}>{fmtMin(r.tempo_espera_min)}</span>
                      </td>
                      <td style={{ padding:'8px 10px' }}>
                        {cls ? <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:cls.bg, color:cls.color, border:`1px solid ${cls.border}`, whiteSpace:'nowrap' }}>{cls.label}</span> : '—'}
                      </td>
                      <td style={{ padding:'8px 10px' }}>
                        {isFalta
                          ? <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:'rgba(59,130,246,0.15)', color:C.blue, border:'1px solid rgba(59,130,246,0.3)' }}>Falta</span>
                          : atrasoReal
                          ? <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:'rgba(245,158,11,0.15)', color:C.amber, border:'1px solid rgba(245,158,11,0.3)' }}>Atraso {fmtMin(Math.abs(tempoAtr))}</span>
                          : <span style={{ fontSize:9.5, color:C.muted }}>—</span>
                        }
                      </td>
                    </tr>
                  )
                })}
            </tbody>
          </table>
        </div>
      </Panel>
    </div>
  )
}


// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function Home() {
  const [tab,         setTab]         = useState('agendas')
  const [agendas,     setAgendas]     = useState([])
  const [espera,      setEspera]      = useState([])
  const [loadingDB,   setLoadingDB]   = useState(true)
  const [storing,     setStoring]     = useState(false)
  const [storeMsg,    setStoreMsg]    = useState('')
  const [timestamp,   setTimestamp]   = useState('')
  const [storageInfo, setStorageInfo] = useState({ agendas:0, espera:0 })

  const loadTable = useCallback(async table => {
    const PAGE=1000; let all=[], offset=0
    while(true) {
      const res=await fetch(`${SB_URL}/rest/v1/${table}?select=*&order=id.asc&limit=${PAGE}&offset=${offset}`,{headers:SBH})
      if(!res.ok) break
      const batch=await res.json()
      if(!Array.isArray(batch)||!batch.length) break
      all=all.concat(batch)
      if(batch.length<PAGE) break
      offset+=PAGE
    }
    return all
  }, [])

  useEffect(() => {
    const load = async () => {
      try {
        setStoreMsg('Conectando…')
        const [ag,esp]=await Promise.all([loadTable('agendas'),loadTable('espera')])
        if(ag.length||esp.length) {
          setAgendas(ag); setEspera(esp)
          setStorageInfo({agendas:ag.length,espera:esp.length})
          const ts=ag[0]?.verif_ts||esp[0]?.verif_ts||''; if(ts) setTimestamp(ts)
          setStoreMsg(`☁ ${ag.length.toLocaleString('pt-BR')} agendas · ${esp.length.toLocaleString('pt-BR')} esperas`)
          setTimeout(()=>setStoreMsg(''),4000)
        } else setStoreMsg('')
      } catch(e) { setStoreMsg(`Erro: ${e.message}`) }
      setLoadingDB(false)
    }
    load()
  }, [loadTable])

  const handleUpload = useCallback(async e => {
    const file=e.target.files[0]; if(!file) return; e.target.value=''
    setStoring(true)
    const now=new Date(), pad=n=>String(n).padStart(2,'0')
    const ts=`${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`
    setTimestamp(ts)
    try {
      setStoreMsg('Lendo planilha…')
      const buf=await file.arrayBuffer()
      const wb=XLSX.read(buf,{type:'buffer'})
      const ws=wb.Sheets['PONTOS']||wb.Sheets[wb.SheetNames[0]]
      const json=XLSX.utils.sheet_to_json(ws,{range:3,defval:''})
      setStoreMsg(`${json.length.toLocaleString('pt-BR')} linhas — limpando banco…`)
      const delRes=await fetch('/api/save',{method:'DELETE'})
      if(!delRes.ok){setStoreMsg('⚠ Erro ao limpar banco');setStoring(false);return}
      const CHUNK=500
      for(let i=0;i<json.length;i+=CHUNK){
        setStoreMsg(`Agendas… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`)
        await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:json.slice(i,i+CHUNK),ts,table:'agendas'})})
      }
      for(let i=0;i<json.length;i+=CHUNK){
        setStoreMsg(`Espera… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`)
        await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:json.slice(i,i+CHUNK),ts,table:'espera'})})
      }
      setStoreMsg('Recarregando…')
      const [ag,esp]=await Promise.all([loadTable('agendas'),loadTable('espera')])
      setAgendas(ag); setEspera(esp); setStorageInfo({agendas:ag.length,espera:esp.length})
      setStoreMsg(`✓ ${ag.length.toLocaleString('pt-BR')} agendas · ${esp.length.toLocaleString('pt-BR')} esperas`)
      setTimeout(()=>setStoreMsg(''),4000)
    } catch(e){setStoreMsg(`⚠ Erro: ${e.message}`)}
    setStoring(false)
  }, [loadTable])

  const handleClear = async () => {
    if(!confirm('Apagar TODOS os dados do banco?')) return
    setStoring(true); setStoreMsg('Apagando…')
    await fetch('/api/save',{method:'DELETE'})
    setAgendas([]); setEspera([]); setStorageInfo({agendas:0,espera:0}); setTimestamp(''); setStoreMsg(''); setStoring(false)
  }

  return (
    <div style={{ display:'flex', minHeight:'100vh', background:C.bg, fontFamily:"'DM Sans','Segoe UI',sans-serif", color:C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800;900&family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(0,201,167,0.25);border-radius:4px}
        select option{background:#0B1424;color:#F1F5F9}
        input::placeholder{color:#475569}
        select{appearance:auto}
      `}</style>

      {/* SIDEBAR */}
      <Sidebar
        tab={tab} setTab={setTab}
        storageInfo={storageInfo} storeMsg={storeMsg}
        timestamp={timestamp} storing={storing} loadingDB={loadingDB}
        onUpload={handleUpload} onClear={handleClear}
      />

      {/* MAIN */}
      <div style={{ marginLeft:260, flex:1, minHeight:'100vh', display:'flex', flexDirection:'column' }}>
        {/* Top bar */}
        <div style={{
          height:56, background:`${C.bgCard}CC`, backdropFilter:'blur(16px)',
          borderBottom:`1px solid ${C.border}`,
          display:'flex', alignItems:'center', justifyContent:'space-between',
          padding:'0 28px', position:'sticky', top:0, zIndex:40,
        }}>
          <div>
            <span style={{ fontSize:15, fontWeight:800, color:C.text, fontFamily:"'Syne',sans-serif" }}>
              {tab==='agendas' ? '🗓 Agendas Médicas' : '⏱ Fila de Espera'}
            </span>
            {storageInfo[tab]>0 && (
              <span style={{ marginLeft:10, fontSize:11, color:C.muted }}>
                {storageInfo[tab].toLocaleString('pt-BR')} registros
              </span>
            )}
          </div>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            <div style={{ width:7, height:7, borderRadius:'50%', background:C.teal, boxShadow:`0 0 8px ${C.teal}` }} />
            <span style={{ fontSize:11, color:C.sub }}>Banco conectado</span>
          </div>
        </div>

        {/* Content */}
        <div style={{ padding:'24px 28px', flex:1 }}>
          {loadingDB ? (
            <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:16 }}>
              <div style={{ fontSize:40 }}>⏳</div>
              <div style={{ fontSize:16, color:C.sub, fontFamily:"'Syne',sans-serif", fontWeight:700 }}>Conectando ao banco de dados…</div>
              {storeMsg && <div style={{ fontSize:12, color:C.amber }}>{storeMsg}</div>}
            </div>
          ) : (
            <>
              {tab==='agendas' && <TabAgendas rows={agendas} />}
              {tab==='espera'  && <TabEspera  rows={espera}  />}
            </>
          )}
        </div>
      </div>
    </div>
  )
}