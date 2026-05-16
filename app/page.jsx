'use client'

import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'

// ── THEME ─────────────────────────────────────────────────────────────────────
const C = {
  bg:      '#06080F',
  bgCard:  '#0A0D16',
  bgCard2: '#0D1020',
  amber:   '#F59E0B',
  orange:  '#F97316',
  rose:    '#F43F5E',
  teal:    '#00C9A7',
  blue:    '#3B82F6',
  violet:  '#8B5CF6',
  emerald: '#10B981',
  cyan:    '#06B6D4',
  text:    '#F1F5F9',
  sub:     '#94A3B8',
  muted:   '#475569',
  border:  'rgba(245,158,11,0.1)',
  border2: 'rgba(245,158,11,0.22)',
  borderGray: 'rgba(255,255,255,0.05)',
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
  // HOJE = data mais recente do dataset (não o sistema)
  if (periodo==='HOJE')   return d => d === max
  // ONTEM = penúltima data do dataset
  if (periodo==='ONTEM')  {
    const prev = s.length>=2 ? s[s.length-2] : null
    return d => prev ? d===prev : false
  }
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
  'SEM PONTO':      { label:'Sem Ponto',      color:'#64748B', glow:'rgba(100,116,139,0.2)' },
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

// Classifica TEMPO_DE_ESPERA (coluna AB)
const clsEspera = m => {
  if (!m || m < 15) return null
  if (m <= 30) return { label:'Espera Moderada', color:'#F59E0B', bg:'rgba(245,158,11,0.08)', border:'rgba(245,158,11,0.22)' }
  if (m <= 89) return { label:'Espera Grave',    color:'#F97316', bg:'rgba(249,115,22,0.1)',  border:'rgba(249,115,22,0.28)' }
  return             { label:'Espera Crítica',  color:'#F43F5E', bg:'rgba(244,63,94,0.1)',   border:'rgba(244,63,94,0.28)'  }
}

const PERIODOS = [
  {key:'HOJE',label:'Hoje'},{key:'ONTEM',label:'Ontem'},{key:'SEMANA',label:'Semana'},
  {key:'MES',label:'Mês'},{key:'ANO',label:'Ano'},{key:'PERIODO',label:'Período'},
]

// ── ATOMS ─────────────────────────────────────────────────────────────────────
function PeriodoBar({ value, onChange, allDates, dateFrom, dateTo, onDateFrom, onDateTo, label }) {
  return (
    <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
      <div style={{ display:'flex', background:'rgba(255,255,255,0.03)', border:`1px solid ${C.border}`, borderRadius:10, padding:3, gap:2 }}>
        {PERIODOS.map(p => (
          <button key={p.key} onClick={()=>onChange(p.key)} style={{
            padding:'5px 12px', borderRadius:7, border:'none', cursor:'pointer', fontSize:11, fontWeight:700, transition:'all .2s',
            background: value===p.key ? C.amber : 'transparent',
            color: value===p.key ? '#1a0800' : C.muted,
          }}>{p.label}</button>
        ))}
      </div>
      {value==='PERIODO' && (<>
        <input type="date" value={dateFrom} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateFrom(e.target.value)}
          style={{ background:'rgba(255,255,255,0.04)', border:`1px solid ${C.border2}`, borderRadius:8, color:C.text, fontSize:11, padding:'5px 9px', outline:'none', colorScheme:'dark' }} />
        <span style={{ color:C.muted }}>→</span>
        <input type="date" value={dateTo} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateTo(e.target.value)}
          style={{ background:'rgba(255,255,255,0.04)', border:`1px solid ${C.border2}`, borderRadius:8, color:C.text, fontSize:11, padding:'5px 9px', outline:'none', colorScheme:'dark' }} />
      </>)}
      {label && <span style={{ fontSize:11, color:C.muted }}>{label}</span>}
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
        style={{ background:'rgba(6,8,15,0.95)', border:`1px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
        <option value="TODOS">Todos os Estados</option>
        {ufs.map(u=><option key={u}>{u}</option>)}
      </select>
      {extra}
      {showClear && <button onClick={onClear} style={{ background:'rgba(244,63,94,0.08)', border:`1px solid rgba(244,63,94,0.25)`, borderRadius:10, color:C.rose, fontSize:12, padding:'8px 12px', cursor:'pointer' }}>✕ Limpar</button>}
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

  const allDates  = useMemo(() => [...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(), [rows])
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

  const agStats = useMemo(() => {
    const cnt    = filtered.length
    const atraso = filtered.filter(d=>String(d.status||'').toUpperCase().includes('ATRASO')).length
    const sponto = filtered.filter(d=>d.hr_entrada_min===null||d.hr_entrada_min===undefined).length
    const uMap={},mMap={},sMap={}
    filtered.forEach(d=>{
      const u=d.nm_local||'?'; uMap[u]=(uMap[u]||0)+1
      const m=d.nm_medico||''; if(m) mMap[m]=(mMap[m]||0)+1
      const s=d.status||'OK';  sMap[s]=(sMap[s]||0)+1
    })
    const topU  = Object.entries(uMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const topM  = Object.entries(mMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const sBrk  = Object.entries(sMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)
    const dMap={}; filtered.forEach(d=>{ if(d.data_agenda) dMap[d.data_agenda]=(dMap[d.data_agenda]||0)+1 })
    const byD   = Object.entries(dMap).map(([k,v])=>({k,v})).sort((a,b)=>a.k.localeCompare(b.k)).slice(-28)
    return { cnt, atraso, sponto, topU, topM, sBrk, byD,
      atrasoPct: cnt>0?((atraso/cnt)*100).toFixed(1):'0',
      sponPct:   cnt>0?((sponto/cnt)*100).toFixed(1):'0' }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:14 }}>
      <div style={{ fontSize:44 }}>📋</div>
      <div style={{ fontSize:18, fontWeight:700, color:C.text, fontFamily:"'Syne',sans-serif" }}>Nenhuma agenda carregada</div>
      <div style={{ fontSize:13, color:C.muted }}>Use o botão para carregar uma planilha</div>
    </div>
  )

  const { cnt, atraso, sponto, topU, topM, sBrk, byD, atrasoPct, sponPct } = agStats
  const maxU = topU[0]?.v||1, maxM = topM[0]?.v||1

  return (
    <div>
      <div style={{ marginBottom:18 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={`${cnt.toLocaleString('pt-BR')} registros`} />
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||statusFilt!=='TODOS'||!!search}
        onClear={()=>{setUfFilt('TODOS');setStatusFilt('TODOS');setSearch('')}}
        extra={<select value={statusFilt} onChange={e=>setStatusFilt(e.target.value)}
          style={{ background:'rgba(6,8,15,0.95)', border:`1px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
          <option value="TODOS">Todos os Status</option>
          {statuses.map(s=><option key={s}>{s}</option>)}
        </select>} />

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12, marginBottom:18 }}>
        {[
          { icon:'🗓', label:'Total Agendas',   value:cnt.toLocaleString('pt-BR'),          color:C.teal    },
          { icon:'⚠️', label:'Em Atraso',       value:atraso.toLocaleString('pt-BR'),       color:C.rose,    sub:`${atrasoPct}% do total`  },
          { icon:'🚫', label:'Sem Ponto',       value:sponto.toLocaleString('pt-BR'),       color:C.violet,  sub:`${sponPct}% do total`    },
          { icon:'✅', label:'Com Atendimento', value:(cnt-sponto).toLocaleString('pt-BR'), color:C.emerald, sub:`${(100-Number(sponPct)).toFixed(1)}% presentes` },
        ].map(k=>(
          <div key={k.label} style={{ background:`${k.color}08`, border:`1px solid ${k.color}20`, borderRadius:14, padding:'18px 20px', position:'relative', overflow:'hidden' }}>
            <div style={{ position:'absolute', top:0, left:0, right:0, height:2, background:`linear-gradient(90deg,${k.color},transparent)` }} />
            <div style={{ fontSize:9.5, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:10 }}>{k.icon} {k.label}</div>
            <div style={{ fontSize:32, fontWeight:800, color:k.color, letterSpacing:'-1px', lineHeight:1 }}>{k.value}</div>
            {k.sub && <div style={{ fontSize:10, color:C.muted, marginTop:7 }}>{k.sub}</div>}
          </div>
        ))}
      </div>

      {/* Status breakdown */}
      <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px', marginBottom:14 }}>
        <div style={{ fontSize:11, fontWeight:700, color:C.sub, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:14 }}>📊 Por Status</div>
        <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
          {sBrk.slice(0,6).map(({k,v})=>{
            const cfg=getStatusCfg(k)
            const pct=cnt>0?((v/cnt)*100).toFixed(1):'0'
            return (
              <div key={k} style={{ display:'flex', alignItems:'center', gap:10 }}>
                <div style={{ width:8,height:8,borderRadius:'50%',background:cfg.color,flexShrink:0 }} />
                <div style={{ flex:1, fontSize:11, color:C.sub, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{cfg.label}</div>
                <span style={{ fontSize:11, fontWeight:700, color:cfg.color, minWidth:36, textAlign:'right' }}>{v.toLocaleString('pt-BR')}</span>
                <div style={{ width:80, background:'rgba(255,255,255,0.05)', borderRadius:3, height:4, overflow:'hidden', flexShrink:0 }}>
                  <div style={{ height:'100%', background:cfg.color, width:`${(v/(sBrk[0]?.v||1))*100}%` }} />
                </div>
                <span style={{ fontSize:10, color:C.muted, minWidth:32, textAlign:'right' }}>{pct}%</span>
              </div>
            )
          })}
        </div>
      </div>

      {/* Rankings */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:14 }}>
        <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px' }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.sub, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:14 }}>🏥 Top Unidades</div>
          {topU.map((u,i)=>{
            const pct=maxU>0?(u.v/maxU)*100:0
            return (
              <div key={u.n} style={{ marginBottom:11 }}>
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4, gap:8 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, minWidth:0 }}>
                    <span style={{ fontSize:10, fontWeight:800, color:i===0?C.amber:C.muted, minWidth:18, flexShrink:0 }}>#{i+1}</span>
                    <span style={{ fontSize:12, color:C.text, fontWeight:500, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{u.n}</span>
                  </div>
                  <span style={{ fontSize:12, fontWeight:700, color:i<3?C.amber:C.teal, flexShrink:0 }}>{u.v}</span>
                </div>
                <div style={{ background:'rgba(255,255,255,0.05)', borderRadius:4, height:4, overflow:'hidden' }}>
                  <div style={{ height:'100%', borderRadius:4, background:i<3?C.amber:C.teal, width:`${pct}%`, transition:'width .7s ease' }} />
                </div>
              </div>
            )
          })}
        </div>
        <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px' }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.sub, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:14 }}>👨‍⚕️ Top Médicos</div>
          {topM.map((m,i)=>{
            const pct=maxM>0?(m.v/maxM)*100:0
            return (
              <div key={m.n} style={{ marginBottom:11 }}>
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4, gap:8 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, minWidth:0 }}>
                    <span style={{ fontSize:10, fontWeight:800, color:i===0?C.amber:C.muted, minWidth:18, flexShrink:0 }}>#{i+1}</span>
                    <span style={{ fontSize:12, color:C.text, fontWeight:500, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{m.n}</span>
                  </div>
                  <span style={{ fontSize:12, fontWeight:700, color:i<3?C.violet:C.blue, flexShrink:0 }}>{m.v}</span>
                </div>
                <div style={{ background:'rgba(255,255,255,0.05)', borderRadius:4, height:4, overflow:'hidden' }}>
                  <div style={{ height:'100%', borderRadius:4, background:i<3?C.violet:C.blue, width:`${pct}%`, transition:'width .7s ease' }} />
                </div>
              </div>
            )
          })}
        </div>
      </div>

      {/* Tabela */}
      <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px' }}>
        <div style={{ fontSize:11, fontWeight:700, color:C.sub, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:14 }}>📋 Detalhamento</div>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11.5 }}>
            <thead>
              <tr style={{ borderBottom:'1px solid rgba(255,255,255,0.06)' }}>
                {['Data','Unidade','Médico','Especialidade','UF','Status','Atraso'].map(h=>(
                  <th key={h} style={{ padding:'7px 10px', textAlign:'left', color:C.muted, fontWeight:700, fontSize:9.5, textTransform:'uppercase', letterSpacing:'.07em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0,50).map((r,i)=>{
                const cfg=getStatusCfg(r.status)
                return (
                  <tr key={i} style={{ borderBottom:'1px solid rgba(255,255,255,0.03)', transition:'background .1s' }}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(245,158,11,0.04)'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <td style={{ padding:'8px 10px', color:C.sub, whiteSpace:'nowrap' }}>{fmtDate(r.data_agenda)}</td>
                    <td style={{ padding:'8px 10px', color:C.text, fontWeight:600, maxWidth:180, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_local}</td>
                    <td style={{ padding:'8px 10px', color:C.sub, whiteSpace:'nowrap' }}>{r.nm_medico}</td>
                    <td style={{ padding:'8px 10px', color:C.muted, maxWidth:140, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.ds_especialidade}</td>
                    <td style={{ padding:'8px 10px', color:C.muted }}>{r.uf}</td>
                    <td style={{ padding:'8px 10px' }}>
                      <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:cfg.glow, color:cfg.color, border:`0.5px solid ${cfg.color}30` }}>{r.status||'—'}</span>
                    </td>
                    <td style={{ padding:'8px 10px', color:r.atraso==='SIM'?C.rose:C.emerald, fontWeight:700 }}>{r.atraso||'—'}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ── TAB ESPERA ────────────────────────────────────────────────────────────────
function TabEspera({ rows }) {
  const [periodo,     setPeriodo]     = useState('MES')
  const [dateFrom,    setDateFrom]    = useState('')
  const [dateTo,      setDateTo]      = useState('')
  const [ufFilt,      setUfFilt]      = useState('TODOS')
  const [search,      setSearch]      = useState('')
  const [horaFilt,    setHoraFilt]    = useState('TODAS')
  const [horaFiltFim, setHoraFiltFim] = useState('TODAS')
  const [unidFilt,    setUnidFilt]    = useState('')
  const [justModal,   setJustModal]   = useState(false)
  const [justificativas, setJustificativas] = useState({})
  const [justLoading, setJustLoading] = useState(false)
  const [trendView,   setTrendView]   = useState('real')

  const allDates  = useMemo(() => [...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(), [rows])
  const periodoFn = useMemo(() => buildFilter(allDates, periodo, dateFrom, dateTo), [allDates, periodo, dateFrom, dateTo])
  const ufs       = useMemo(() => [...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(), [rows])
  const dataRef   = useMemo(() => allDates.length ? allDates[allDates.length-1] : '', [allDates])

  // Carrega justificativas do Supabase
  useEffect(() => {
    if (!dataRef) return
    fetch(`${SB_URL}/rest/v1/justificativas_espera?select=hora,quantidade&data_ref=eq.${dataRef}&order=hora.asc`, { headers: SBH })
      .then(r=>r.json()).then(data=>{
        if (!Array.isArray(data)) return
        const map={}; data.forEach(r=>{ map[r.hora]=r.quantidade })
        setJustificativas(map)
      }).catch(console.error)
  }, [dataRef])

  const saveJustificativa = useCallback(async (hora, qtd) => {
    if (!dataRef) return
    setJustLoading(true)
    try {
      await fetch(`${SB_URL}/rest/v1/justificativas_espera`, {
        method:'POST',
        headers:{ ...SBH, 'Prefer':'resolution=merge-duplicates' },
        body: JSON.stringify({ data_ref:dataRef, hora:parseInt(hora), quantidade:parseInt(qtd)||0 }),
      })
      setJustificativas(prev=>({ ...prev, [hora]:parseInt(qtd)||0 }))
    } catch(e){ console.error(e) }
    setJustLoading(false)
  }, [dataRef])

  // Horas com esperas >= 15min para filtro e modal
  const horasDisp = useMemo(() => {
    const set = new Set()
    rows.filter(d=>periodoFn(d.data_agenda) && d.tempo_espera_min>=15)
      .forEach(d=>{
        const h=d.hr_registro_espera_min
        if (h===null||h===undefined) return
        const hora=Math.floor(h/60)
        if (hora>=0&&hora<=23) set.add(hora)
      })
    return [...set].sort((a,b)=>a-b)
  }, [rows, periodoFn])

  const filtered = useMemo(() => {
    let r = rows.filter(d=>periodoFn(d.data_agenda))
    if (ufFilt!=='TODOS') r=r.filter(d=>d.uf===ufFilt)
    if (search) { const q=search.toLowerCase(); r=r.filter(d=>[d.nm_local,d.nm_medico,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))) }
    // Filtro por intervalo de hora — baseado em HR_REGISTRO_ESPERA
    if (horaFilt!=='TODAS') {
      const ini=parseInt(horaFilt,10)
      const fim=horaFiltFim==='TODAS'?ini:parseInt(horaFiltFim,10)
      r=r.filter(d=>{
        const h=d.hr_registro_espera_min
        if (h===null||h===undefined) return false
        const hora=Math.floor(h/60)
        return hora>=ini && hora<=fim
      })
    }
    if (unidFilt) r=r.filter(d=>d.nm_local===unidFilt)
    return r
  }, [rows, periodoFn, ufFilt, search, horaFilt, horaFiltFim, unidFilt])

  const espStats = useMemo(() => {
    const comEsp = filtered.filter(d=>d.tempo_espera_min!==null&&d.tempo_espera_min!==undefined&&d.tempo_espera_min>=15)
    const totalReg = filtered.length
    const totalPac = filtered.reduce((a,d)=>a+(d.qt_pacientes_aguardando||0),0)
    const modCnt   = comEsp.filter(d=>d.tempo_espera_min<=30).length
    const grvCnt   = comEsp.filter(d=>d.tempo_espera_min>30&&d.tempo_espera_min<=89).length
    const critCnt  = comEsp.filter(d=>d.tempo_espera_min>=90).length
    const totalEsp = modCnt+grvCnt+critCnt

    // Feed: agrupa por hora+unidade, pior TEMPO_DE_ESPERA
    const grupoMap={}
    comEsp.forEach(d=>{
      const h=d.hr_registro_espera_min
      if (h===null||h===undefined) return
      const hora=Math.floor(h/60), mins=Math.floor(h%60)
      if (hora<0||hora>23) return
      const horaStr=String(hora).padStart(2,'0')+':'+String(mins).padStart(2,'0')
      const unidade=d.nm_local||'Sem Unidade'
      const key=`${horaStr}||${unidade}`
      if (!grupoMap[key]) grupoMap[key]={ horaStr, hora, nm_local:unidade, uf:d.uf||'', cidade:d.cidade||'', maxTempo:0, pac:0, count:0 }
      // Pacientes vem do registro com MAIOR TEMPO_DE_ESPERA (não soma)
      if (d.tempo_espera_min>grupoMap[key].maxTempo) {
        grupoMap[key].maxTempo=d.tempo_espera_min
        grupoMap[key].pac=d.qt_pacientes_aguardando||0
      }
      grupoMap[key].count+=1
    })
    const feedList=Object.values(grupoMap).sort((a,b)=>b.maxTempo-a.maxTempo).slice(0,25)

    // Top unidades
    const uMap={}
    comEsp.forEach(d=>{
      const u=d.nm_local||'?'
      if (!uMap[u]) uMap[u]={ n:u, total:0, crit:0, grv:0, mod:0 }
      uMap[u].total++
      const m=d.tempo_espera_min
      if (m>=90) uMap[u].crit++; else if (m>=31) uMap[u].grv++; else uMap[u].mod++
    })
    const topU=Object.values(uMap).sort((a,b)=>b.crit-a.crit||b.grv-a.grv).slice(0,6)

    // Médicos
    const faltasList  = filtered.filter(d=>String(d.atraso||'').toUpperCase()==='FALTA')
    const atrasosList = filtered.filter(d=>{
      if (String(d.atraso||'').toUpperCase()!=='SIM') return false
      const t=d.tempo_atraso_min; return t!==null&&t!==undefined&&Math.abs(t)>31
    })
    const sMap={}
    atrasosList.forEach(d=>{ const s=d.status||'Sem Status'; sMap[s]=(sMap[s]||0)+1 })
    const statusAt=Object.entries(sMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)

    // Por data — para o gráfico de tendência
    const dMap = {}
    filtered.forEach(d => {
      const dt = d.data_agenda; if (!dt) return
      if (!dMap[dt]) dMap[dt] = { date:dt, mod:0, grv:0, crit:0, pac:0, faltas:0, atrasos:0 }
      const m = d.tempo_espera_min
      if (m>=90)      dMap[dt].crit++
      else if (m>=31) dMap[dt].grv++
      else if (m>=15) dMap[dt].mod++
      dMap[dt].pac += d.qt_pacientes_aguardando||0
      if (String(d.atraso||'').toUpperCase()==='FALTA') dMap[dt].faltas++
      if (String(d.atraso||'').toUpperCase()==='SIM' && d.tempo_atraso_min && Math.abs(d.tempo_atraso_min)>31) dMap[dt].atrasos++
    })
    const byDate = Object.values(dMap).sort((a,b)=>a.date.localeCompare(b.date))

    // Projeção — média móvel dos últimos 3 dias projetada +5 dias
    const addDay = (dateStr, n) => {
      const d = new Date(dateStr+'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+n)
      return d.toISOString().slice(0,10)
    }
    const pts = byDate.slice(-3)
    const slopeCrit = pts.length>=2 ? (pts[pts.length-1].crit - pts[0].crit)/(pts.length-1) : 0
    const slopePac  = pts.length>=2 ? (pts[pts.length-1].pac  - pts[0].pac) /(pts.length-1) : 0
    const lastDate  = byDate[byDate.length-1]?.date || ''
    const lastCrit  = byDate[byDate.length-1]?.crit || 0
    const lastPac   = byDate[byDate.length-1]?.pac  || 0
    const projData  = lastDate ? Array.from({length:5},(_,i)=>({
      date: addDay(lastDate, i+1),
      crit: Math.max(0, Math.round(lastCrit + slopeCrit*(i+1))),
      pac:  Math.max(0, Math.round(lastPac  + slopePac *(i+1))),
      isProj: true,
    })) : []

    return { totalReg, totalPac, modCnt, grvCnt, critCnt, totalEsp, feedList, topU, faltasList, atrasosList, statusAt, byDate, projData }
  }, [filtered])

  const totalJust = useMemo(()=>Object.values(justificativas).reduce((a,v)=>a+(parseInt(v)||0),0),[justificativas])
  const metaPct   = espStats.totalEsp>0?Math.min(Math.round((totalJust/espStats.totalEsp)*100),100):0
  const metaColor = metaPct>=80?C.emerald:metaPct>=50?C.amber:C.rose

  // Desestrutura espStats antes de usar faltasList nas variáveis computadas
  const { totalReg, totalPac, modCnt, grvCnt, critCnt, totalEsp, feedList, topU, faltasList, atrasosList, statusAt, byDate, projData } = espStats

  // Médicos — computed values for donut
  const medTotalProb = faltasList.length + atrasosList.length
  const medFPct  = medTotalProb>0 ? Math.round((faltasList.length/medTotalProb)*100)  : 0
  const medAPct  = medTotalProb>0 ? Math.round((atrasosList.length/medTotalProb)*100) : 0
  const medRate  = totalReg>0 ? ((medTotalProb/totalReg)*100).toFixed(1) : '0'
  const medR=44, medCX=54, medCY=54, medCirc=2*Math.PI*medR
  const medFDash = (medFPct/100)*medCirc
  const medADash = (medAPct/100)*medCirc

  // Trend — computed values
  const trendAllData  = trendView==='real' ? byDate : [...byDate, ...projData]
  const trendMaxCrit  = Math.max(...trendAllData.map(d=>d.crit||0), 1)
  const trendMaxPac   = Math.max(...trendAllData.map(d=>d.pac||0),  1)
  const trendTotal    = trendAllData.length
  const trendRealCnt  = byDate.length
  const trendLastReal = byDate[byDate.length-1]
  const trendFirst    = byDate[0]
  const trendVarCrit  = byDate.length>=2 ? ((trendLastReal?.crit||0)-(trendFirst?.crit||0)) : 0
  const trendAvgCrit  = byDate.length>0  ? Math.round(byDate.reduce((a,d)=>a+d.crit,0)/byDate.length) : 0
  const trendMaxDay   = byDate.length>0  ? byDate.reduce((a,d)=>d.crit>a.crit?d:a, byDate[0]) : null
  const trendSlope    = projData.length>0 ? ((projData[0]?.crit||0)-(trendLastReal?.crit||0)) : 0

  const VW=800, VH=160, PL=48, PR=48, PT=10, PB=32
  const CW=VW-PL-PR, CH=VH-PT-PB
  const txPos = (i) => trendTotal<=1 ? PL+CW/2 : PL+(i/(trendTotal-1))*CW
  const tyPos = (v, mx) => PT+CH-(mx>0?(v/mx)*CH:0)
  const tMakePath = (data, key, mx, startI=0) =>
    data.map((d,i)=>`${i+startI===0?'M':'L'}${txPos(i+startI).toFixed(1)},${tyPos(d[key]||0,mx).toFixed(1)}`).join(' ')
  const tYTicks = [0, Math.round(trendMaxCrit*0.5), trendMaxCrit]

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:14 }}>
      <div style={{ fontSize:44 }}>⏱️</div>
      <div style={{ fontSize:18, fontWeight:700, color:C.text, fontFamily:"'Syne',sans-serif" }}>Nenhum dado de espera</div>
      <div style={{ fontSize:13, color:C.muted }}>Use o botão para carregar uma planilha</div>
    </div>
  )

  const horasDispFim = horaFilt==='TODAS'?[]:horasDisp.filter(h=>h>parseInt(horaFilt))

  return (
    <div>
      {/* Modal Justificativas */}
      {justModal && (
        <div style={{ position:'fixed', inset:0, zIndex:200, background:'rgba(0,0,0,0.75)', backdropFilter:'blur(4px)', display:'flex', alignItems:'center', justifyContent:'center' }}
          onClick={()=>setJustModal(false)}>
          <div style={{ background:'#0A0D16', border:`1px solid rgba(245,158,11,0.3)`, borderRadius:16, padding:'24px 28px', minWidth:360, maxWidth:460, boxShadow:'0 24px 80px rgba(0,0,0,0.6)' }}
            onClick={e=>e.stopPropagation()}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
              <div>
                <div style={{ fontSize:14, fontWeight:700, color:C.text }}>Registrar Justificativas</div>
                <div style={{ fontSize:10.5, color:C.muted, marginTop:3 }}>Meta: 80% das esperas com retorno · {dataRef}</div>
              </div>
              <button onClick={()=>setJustModal(false)} style={{ background:'transparent', border:'none', color:C.muted, fontSize:18, cursor:'pointer' }}>✕</button>
            </div>
            <div style={{ background:'rgba(255,255,255,0.03)', border:'0.5px solid rgba(255,255,255,0.07)', borderRadius:10, padding:'12px 14px', marginBottom:18 }}>
              <div style={{ display:'flex', justifyContent:'space-between', marginBottom:7 }}>
                <span style={{ fontSize:11, color:C.muted }}>Progresso atual</span>
                <span style={{ fontSize:12, fontWeight:700, color:metaColor }}>{metaPct}% de 80%</span>
              </div>
              <div style={{ background:'rgba(255,255,255,0.06)', borderRadius:4, height:8, overflow:'hidden' }}>
                <div style={{ height:'100%', width:`${metaPct}%`, background:metaColor, borderRadius:4, transition:'width .4s ease' }} />
              </div>
              <div style={{ display:'flex', justifyContent:'space-between', marginTop:6 }}>
                <span style={{ fontSize:10, color:C.muted }}>{totalJust} justificativas</span>
                <span style={{ fontSize:10, color:C.muted }}>de {totalEsp} esperas</span>
              </div>
            </div>
            <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:280, overflowY:'auto' }}>
              {horasDisp.length===0 && <div style={{ color:C.muted, fontSize:12, textAlign:'center', padding:'16px 0' }}>Sem horas disponíveis.</div>}
              {horasDisp.map(h=>(
                <div key={h} style={{ display:'flex', alignItems:'center', gap:10 }}>
                  <div style={{ width:44, fontSize:12, fontWeight:700, color:C.sub, fontFamily:'monospace', flexShrink:0 }}>{String(h).padStart(2,'0')}:00</div>
                  <div style={{ flex:1, background:'rgba(255,255,255,0.04)', borderRadius:6, height:4, overflow:'hidden' }}>
                    <div style={{ height:'100%', borderRadius:6, background:C.amber, transition:'width .3s', width:`${Math.min(((justificativas[h]||0)/Math.max(totalEsp/Math.max(horasDisp.length,1),1))*100,100)}%` }} />
                  </div>
                  <input type="number" min="0" value={justificativas[h]||''} placeholder="0"
                    onChange={e=>saveJustificativa(h,e.target.value)}
                    style={{ width:60, background:'rgba(255,255,255,0.05)', border:`0.5px solid ${justificativas[h]>0?'rgba(245,158,11,0.4)':'rgba(255,255,255,0.1)'}`, borderRadius:7, color:justificativas[h]>0?C.amber:C.text, fontSize:12, fontWeight:700, padding:'5px 8px', outline:'none', textAlign:'center' }} />
                  <span style={{ fontSize:10, color:C.muted, minWidth:24 }}>just.</span>
                </div>
              ))}
            </div>
            {justLoading && <div style={{ textAlign:'center', marginTop:12, fontSize:11, color:C.amber }}>Salvando…</div>}
          </div>
        </div>
      )}

      {/* Período */}
      <div style={{ marginBottom:16 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('');setUnidFilt('');setHoraFilt('TODAS');setHoraFiltFim('TODAS')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={`${totalReg.toLocaleString('pt-BR')} registros`} />
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}} />

      {/* BARRA DE META */}
      <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'16px 20px', marginBottom:14 }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:10 }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em' }}>Progresso vs Meta 80% — Justificativas de Retorno</div>
          <button onClick={()=>setJustModal(true)} style={{ background:'rgba(245,158,11,0.1)', border:'0.5px solid rgba(245,158,11,0.3)', borderRadius:8, color:C.amber, fontSize:11, fontWeight:700, padding:'5px 12px', cursor:'pointer' }}>+ Registrar</button>
        </div>
        <div style={{ position:'relative', marginBottom:22 }}>
          <div style={{ background:'rgba(255,255,255,0.06)', borderRadius:6, height:12, overflow:'hidden', position:'relative' }}>
            <div style={{ height:'100%', borderRadius:6, width:`${metaPct}%`, transition:'width .6s ease',
              background:metaPct>=80?'linear-gradient(90deg,#10B981,#059669)':metaPct>=50?'linear-gradient(90deg,#F59E0B,#D97706)':'linear-gradient(90deg,#F43F5E,#DC2626)' }} />
            <div style={{ position:'absolute', top:0, bottom:0, left:'80%', width:2, background:'rgba(255,255,255,0.4)' }} />
          </div>
          <div style={{ position:'absolute', top:14, left:'80%', transform:'translateX(-50%)', fontSize:9, color:C.muted, whiteSpace:'nowrap' }}>← Meta 80%</div>
        </div>
        {/* 3 cards classificação */}
        <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:10 }}>
          {[
            { label:'Espera Moderada', value:modCnt,  sub:'15 – 30 min',   color:'#F59E0B' },
            { label:'Espera Grave',    value:grvCnt,  sub:'31 min – 1h29', color:'#F97316' },
            { label:'Espera Crítica',  value:critCnt, sub:'acima de 1h30', color:'#F43F5E' },
          ].map(k=>(
            <div key={k.label} style={{ background:`${k.color}08`, border:`1px solid ${k.color}22`, borderRadius:11, padding:'12px 16px', position:'relative', overflow:'hidden' }}>
              <div style={{ position:'absolute', top:0, left:0, right:0, height:2, background:k.color }} />
              <div style={{ fontSize:9.5, fontWeight:700, color:k.color, textTransform:'uppercase', letterSpacing:'.09em', marginBottom:8 }}>{k.label}</div>
              <div style={{ fontSize:26, fontWeight:800, color:k.color, letterSpacing:'-1px', lineHeight:1 }}>{k.value.toLocaleString('pt-BR')}</div>
              <div style={{ fontSize:10, color:C.muted, marginTop:6 }}>{k.sub}</div>
            </div>
          ))}
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:16, marginTop:12, paddingTop:12, borderTop:'0.5px solid rgba(255,255,255,0.05)' }}>
          <span style={{ fontSize:11, color:C.muted }}>
            <span style={{ color:metaColor, fontWeight:700 }}>{totalJust}</span> justificativas de{' '}
            <span style={{ color:C.sub, fontWeight:600 }}>{totalEsp}</span> esperas ≥ 15min
          </span>
          <span style={{ fontSize:14, fontWeight:800, color:metaColor, marginLeft:'auto' }}>{metaPct}%</span>
          <span style={{ fontSize:10, color:C.muted }}>meta: 80%</span>
        </div>
      </div>

      {/* FEED + TOP UNIDADES */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 320px', gap:14, marginBottom:14 }}>
        {/* Feed */}
        <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px' }}>
          {unidFilt && (
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12, padding:'8px 12px', background:'rgba(245,158,11,0.08)', border:'0.5px solid rgba(245,158,11,0.25)', borderRadius:9 }}>
              <div style={{ width:6,height:6,borderRadius:'50%',background:C.amber,flexShrink:0 }} />
              <span style={{ fontSize:11, color:C.amber, flex:1, fontWeight:600, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>Filtrando: {unidFilt}</span>
              <button onClick={()=>setUnidFilt('')} style={{ background:'transparent', border:'none', color:C.muted, cursor:'pointer', fontSize:12 }}>✕</button>
            </div>
          )}
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:16 }}>
            <div>
              <div style={{ fontSize:13, fontWeight:700, color:C.text }}>Feed de Esperas por Hora</div>
              <div style={{ fontSize:10.5, color:C.muted, marginTop:3 }}>TEMPO_DE_ESPERA · hora via HR_REGISTRO_ESPERA · {unidFilt?'clique ✕ para limpar':'clique na unidade para filtrar'}</div>
            </div>
            {/* Filtro hora intervalo */}
            <div style={{ display:'flex', alignItems:'center', gap:6, flexShrink:0 }}>
              <select value={horaFilt} onChange={e=>{setHoraFilt(e.target.value);setHoraFiltFim('TODAS');setUnidFilt('')}}
                style={{ background:'rgba(255,255,255,0.05)', border:'0.5px solid rgba(245,158,11,0.2)', borderRadius:8, color:C.text, fontSize:11, padding:'5px 8px', outline:'none', cursor:'pointer' }}>
                <option value="TODAS">Todas</option>
                {horasDisp.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}
              </select>
              {horaFilt!=='TODAS' && (<>
                <span style={{ fontSize:11, color:C.muted }}>→</span>
                <select value={horaFiltFim} onChange={e=>setHoraFiltFim(e.target.value)}
                  style={{ background:'rgba(255,255,255,0.05)', border:'0.5px solid rgba(245,158,11,0.2)', borderRadius:8, color:C.text, fontSize:11, padding:'5px 8px', outline:'none', cursor:'pointer' }}>
                  <option value="TODAS">{String(horaFilt).padStart(2,'0')}:00 só</option>
                  {horasDispFim.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}
                </select>
              </>)}
            </div>
          </div>
          {feedList.length===0 ? (
            <div style={{ textAlign:'center', padding:'32px 0', color:C.muted, fontSize:12 }}>Sem esperas ≥ 15min no período/filtro selecionado.</div>
          ) : (
            <div style={{ display:'flex', flexDirection:'column', gap:5, maxHeight:440, overflowY:'auto' }}>
              {feedList.map((item,i)=>{
                const cls    = clsEspera(item.maxTempo)
                const isCrit = item.maxTempo>=90
                const isGrv  = item.maxTempo>=31&&item.maxTempo<90
                const isSel  = unidFilt===item.nm_local
                return (
                  <div key={i} onClick={()=>setUnidFilt(isSel?'':item.nm_local)} style={{
                    display:'flex', alignItems:'center', gap:12, padding:'10px 14px', borderRadius:10, cursor:'pointer',
                    background:isSel?`${cls.color}18`:isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':'rgba(255,255,255,0.02)',
                    border:isSel?`1px solid ${cls.color}55`:`0.5px solid ${i<3?cls.border:'rgba(255,255,255,0.05)'}`,
                    transition:'all .15s',
                  }}
                    onMouseEnter={e=>{ if(!isSel) e.currentTarget.style.background=cls.bg }}
                    onMouseLeave={e=>{ if(!isSel) e.currentTarget.style.background=isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':'rgba(255,255,255,0.02)' }}
                  >
                    <div style={{ width:8,height:8,borderRadius:'50%',background:cls.color,flexShrink:0,boxShadow:isCrit?`0 0 8px ${cls.color}`:'none' }} />
                    <div style={{ fontFamily:'monospace',fontSize:13,fontWeight:700,color:C.sub,flexShrink:0,minWidth:44 }}>{item.horaStr}</div>
                    <div style={{ flex:1,minWidth:0 }}>
                      <div style={{ fontSize:12,fontWeight:600,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap' }}>{item.nm_local}</div>
                      <div style={{ fontSize:10,color:C.muted,marginTop:2 }}>{[item.cidade,item.uf].filter(Boolean).join(' · ')}</div>
                    </div>
                    <div style={{ textAlign:'center',flexShrink:0,minWidth:42 }}>
                      <div style={{ fontSize:12,fontWeight:700,color:'#0EA5E9' }}>{item.pac>0?item.pac:'—'}</div>
                      <div style={{ fontSize:9,color:C.muted }}>pac.</div>
                    </div>
                    <div style={{ fontSize:15,fontWeight:900,color:cls.color,flexShrink:0,minWidth:52,textAlign:'right' }}>{fmtMin(item.maxTempo)}</div>
                    <span style={{ fontSize:9.5,fontWeight:700,padding:'3px 9px',borderRadius:20,background:cls.bg,color:cls.color,border:`0.5px solid ${cls.border}`,whiteSpace:'nowrap',flexShrink:0 }}>{cls.label}</span>
                  </div>
                )
              })}
            </div>
          )}
        </div>

        {/* Médicos — redesenho moderno */}
        <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 22px' }}>
          <div style={{ fontSize:13,fontWeight:700,color:C.text,marginBottom:2 }}>Médicos — Falta e Atraso</div>
          <div style={{ fontSize:10.5,color:C.muted,marginBottom:16 }}>Ocorrências no período · coluna ATRASO</div>
          <div style={{ display:'flex', alignItems:'center', gap:18, marginBottom:18 }}>
            <svg width="108" height="108" viewBox="0 0 108 108" style={{ flexShrink:0 }}>
              <circle cx={medCX} cy={medCY} r={medR} fill="none" stroke="rgba(255,255,255,0.06)" strokeWidth="10" />
              {faltasList.length>0 && <circle cx={medCX} cy={medCY} r={medR} fill="none" stroke="#F43F5E" strokeWidth="10" strokeLinecap="round"
                strokeDasharray={`${medFDash.toFixed(1)} ${(medCirc-medFDash).toFixed(1)}`}
                transform={`rotate(-90 ${medCX} ${medCY})`} opacity={0.9} />}
              {atrasosList.length>0 && <circle cx={medCX} cy={medCY} r={medR} fill="none" stroke="#F59E0B" strokeWidth="10" strokeLinecap="round"
                strokeDasharray={`${medADash.toFixed(1)} ${(medCirc-medADash).toFixed(1)}`}
                transform={`rotate(${(-90+medFPct*3.6).toFixed(1)} ${medCX} ${medCY})`} opacity={0.85} />}
              <text x={medCX} y={medCY-6} textAnchor="middle" fontSize="18" fontWeight="800" fill={medTotalProb>0?'#F97316':'#475569'}>{medTotalProb}</text>
              <text x={medCX} y={medCY+10} textAnchor="middle" fontSize="9" fill="#475569">ocorrências</text>
              <text x={medCX} y={medCY+22} textAnchor="middle" fontSize="8" fill="#334155">{medRate}% do total</text>
            </svg>
            <div style={{ flex:1, display:'flex', flexDirection:'column', gap:10 }}>
              {[
                { label:'Faltas',        value:faltasList.length,  color:C.rose,  pct:medFPct },
                { label:'Atrasos >31min',value:atrasosList.length, color:C.amber, pct:medAPct },
              ].map(k=>(
                <div key={k.label}>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4 }}>
                    <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                      <div style={{ width:8,height:8,borderRadius:'50%',background:k.color }} />
                      <span style={{ fontSize:11,color:C.sub }}>{k.label}</span>
                    </div>
                    <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                      <span style={{ fontSize:11,fontWeight:700,color:k.color }}>{k.value}</span>
                      <span style={{ fontSize:9.5,color:C.muted }}>{k.pct}%</span>
                    </div>
                  </div>
                  <div style={{ background:'rgba(255,255,255,0.05)',borderRadius:3,height:4,overflow:'hidden' }}>
                    <div style={{ height:'100%',background:k.color,width:`${k.pct}%`,borderRadius:3,transition:'width .6s' }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
          {statusAt.length>0 && (
            <div style={{ borderTop:'0.5px solid rgba(255,255,255,0.06)', paddingTop:14 }}>
              <div style={{ fontSize:9.5,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.09em',marginBottom:12 }}>Classificação do Atraso</div>
              {statusAt.map(({k,v})=>{
                const cfg=getStatusCfg(k)
                const pctAt=statusAt[0]?.v>0?Math.round(v/statusAt[0].v*100):0
                return (
                  <div key={k} style={{ marginBottom:10 }}>
                    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4 }}>
                      <span style={{ fontSize:10,color:C.sub }}>{cfg.label}</span>
                      <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                        <span style={{ fontSize:11,fontWeight:700,color:cfg.color }}>{v}</span>
                        <span style={{ fontSize:9,color:C.muted }}>{pctAt}%</span>
                      </div>
                    </div>
                    <div style={{ background:'rgba(255,255,255,0.05)',borderRadius:3,height:5,overflow:'hidden' }}>
                      <div style={{ height:'100%',background:cfg.color,width:`${pctAt}%`,borderRadius:3,transition:'width .6s' }} />
                    </div>
                  </div>
                )
              })}
            </div>
          )}
          {medTotalProb===0 && <div style={{ color:C.muted,fontSize:11,textAlign:'center',padding:'8px 0' }}>Nenhuma ocorrência no período.</div>}
        </div>
      </div>

      {/* TENDÊNCIA */}
      <div style={{ background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)', borderRadius:14, padding:'20px 24px', marginBottom:14 }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:16 }}>
          <div>
            <div style={{ fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.13em' }}>
              TENDÊNCIA DE ESPERAS · {trendRealCnt} DIA{trendRealCnt!==1?'S':''} · TODOS OS ESTADOS
            </div>
            <div style={{ fontSize:10.5,color:C.muted,marginTop:4 }}>Esperas críticas ao longo do tempo</div>
          </div>
          <div style={{ display:'flex', background:'rgba(255,255,255,0.04)', border:'0.5px solid rgba(255,255,255,0.08)', borderRadius:9, padding:3, gap:2 }}>
            {[{key:'real',label:'Real'},{key:'proj',label:'Projeção'}].map(v=>(
              <button key={v.key} onClick={()=>setTrendView(v.key)} style={{
                padding:'5px 16px', borderRadius:6, border:'none', cursor:'pointer', fontSize:11, fontWeight:700, transition:'all .15s',
                background:trendView===v.key?C.amber:'transparent',
                color:trendView===v.key?'#1a0800':C.muted,
              }}>{v.label}</button>
            ))}
          </div>
        </div>

        {/* KPIs */}
        <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:10, marginBottom:16 }}>
          {(trendView==='real' ? [
            { label:'Esperas Críticas (hoje)', value:trendLastReal?.crit??'—',                                                    color:C.rose   },
            { label:'Variação no período',     value:byDate.length>=2?(trendVarCrit>=0?`+${trendVarCrit}`:String(trendVarCrit)):'—', color:trendVarCrit>0?C.rose:trendVarCrit<0?C.emerald:C.muted },
            { label:'Média diária crítica',    value:trendAvgCrit||'—',                                                          color:C.amber  },
            { label:'Pior dia registrado',     value:trendMaxDay?`${trendMaxDay.crit} em ${trendMaxDay.date.slice(5).replace('-','/')}`:'—', color:C.orange },
          ] : [
            { label:'Projeção amanhã',         value:projData[0]?.crit??'—',                color:C.rose   },
            { label:'Projeção +5 dias',        value:projData[4]?.crit??'—',                color:C.orange },
            { label:'Tendência diária',        value:trendSlope>=0?`+${trendSlope}`:String(trendSlope), color:trendSlope>0?C.rose:trendSlope<0?C.emerald:C.muted },
            { label:'Média atual (base)',       value:trendAvgCrit||'—',                    color:C.amber  },
          ]).map((k,i)=>(
            <div key={k.label} style={{ background:`${k.color}08`,border:`0.5px solid ${k.color}20`,borderRadius:10,padding:'11px 14px' }}>
              <div style={{ fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.08em',marginBottom:7 }}>{k.label}</div>
              <div style={{ fontSize:22,fontWeight:800,color:k.color,letterSpacing:'-.5px' }}>{k.value}</div>
            </div>
          ))}
        </div>

        {/* Chart */}
        {trendAllData.length===0 ? (
          <div style={{ textAlign:'center', padding:'40px 0', color:C.muted, fontSize:12 }}>
            Sem dados suficientes. Carregue uma planilha com múltiplos dias para ver a tendência.
          </div>
        ) : (
          <div>
            <svg width="100%" viewBox={`0 0 ${VW} ${VH}`} style={{ display:'block', overflow:'visible' }}>


              {/* Grid */}
              {tYTicks.map((v,ti)=>(
                <g key={ti}>
                  <line x1={PL} y1={tyPos(v,trendMaxCrit)} x2={VW-PR} y2={tyPos(v,trendMaxCrit)} stroke="rgba(255,255,255,0.05)" strokeWidth="0.5" />
                  <text x={PL-6} y={tyPos(v,trendMaxCrit)+4} textAnchor="end" fontSize="9" fill="#334155">{v}</text>
                </g>
              ))}

              {/* Area */}
              {trendRealCnt>=2 && (
                <path d={`${tMakePath(byDate,'crit',trendMaxCrit)} L${txPos(trendRealCnt-1).toFixed(1)},${(PT+CH).toFixed(1)} L${txPos(0).toFixed(1)},${(PT+CH).toFixed(1)} Z`}
                  fill="rgba(244,63,94,0.15)" />
              )}

              {/* Separator */}
              {trendView==='proj' && trendRealCnt>0 && trendTotal>trendRealCnt && (
                <line x1={txPos(trendRealCnt-1)} y1={PT} x2={txPos(trendRealCnt-1)} y2={PT+CH}
                  stroke="rgba(245,158,11,0.3)" strokeWidth="1" strokeDasharray="4,3" />
              )}

              {/* Lines — real */}
              {trendRealCnt>=2 && (
                <path d={tMakePath(byDate,'grv',trendMaxCrit)} fill="none" stroke="#F97316" strokeWidth="1.5" opacity={0.7} />
              )}
              {trendRealCnt>=2 && (
                <path d={tMakePath(byDate,'crit',trendMaxCrit)} fill="none" stroke="#F43F5E" strokeWidth="2" />
              )}

              {/* Lines — projection */}
              {trendView==='proj' && projData.length>=1 && trendRealCnt>=1 && (
                <path d={`M${txPos(trendRealCnt-1).toFixed(1)},${tyPos(trendLastReal?.crit||0,trendMaxCrit).toFixed(1)} ${tMakePath(projData,'crit',trendMaxCrit,trendRealCnt)}`}
                  fill="none" stroke="#F43F5E" strokeWidth="1.5" strokeOpacity="0.45" strokeDasharray="6,4" />
              )}

              {/* Dots — real */}
              {byDate.map((d,i)=>(
                <circle key={`dot-${i}`} cx={txPos(i)} cy={tyPos(d.crit,trendMaxCrit)} r="3.5"
                  fill="#F43F5E" stroke="#06080F" strokeWidth="1.5" />
              ))}

              {/* Proj label */}
              {trendView==='proj' && trendRealCnt>0 && trendTotal>trendRealCnt && (
                <text x={txPos(trendRealCnt)+4} y={PT+12} fontSize="8" fill="rgba(245,158,11,0.6)">Projeção</text>
              )}

              {/* X labels */}
              {trendAllData.map((d,i)=>{
                const show = trendTotal<=10 || i===0 || i===trendTotal-1 || i%Math.ceil(trendTotal/6)===0
                return show ? (
                  <text key={`xl-${i}`} x={txPos(i)} y={VH-4} textAnchor="middle" fontSize="9"
                    fill={d.isProj?'rgba(245,158,11,0.45)':'#334155'}>
                    {d.date.slice(5).replace('-','/')}
                  </text>
                ) : null
              })}

              <line x1={PL} y1={PT+CH} x2={VW-PR} y2={PT+CH} stroke="rgba(255,255,255,0.08)" strokeWidth="0.5" />
            </svg>

            {/* Legend */}
            <div style={{ display:'flex', gap:16, marginTop:8, justifyContent:'center', flexWrap:'wrap' }}>
              {[
                { color:'#F43F5E', label:'Espera Crítica (real)',  dashed:false },
                { color:'#F97316', label:'Espera Grave (real)',    dashed:false },
              ].concat(trendView==='proj' ? [
                { color:'#F43F5E', label:'Projeção crítica', dashed:true },
              ] : []).map(l=>(
                <div key={l.label} style={{ display:'flex',alignItems:'center',gap:5 }}>
                  <div style={{ width:16,height:3,borderRadius:2,background:l.dashed?'transparent':l.color,
                    border:l.dashed?`1.5px dashed ${l.color}`:'none',opacity:l.dashed?.6:1 }} />
                  <span style={{ fontSize:9.5,color:C.muted }}>{l.label}</span>
                </div>
              ))}
            </div>

            {trendView==='proj' && (
              <div style={{ marginTop:12,padding:'9px 14px',background:'rgba(245,158,11,0.05)',border:'0.5px solid rgba(245,158,11,0.18)',borderRadius:8,fontSize:10.5,color:C.muted }}>
                <span style={{ color:C.amber,fontWeight:700 }}>Como a projeção é calculada — </span>
                Média dos últimos {Math.min(byDate.length,3)} dia{byDate.length!==1?'s':''} + tendência linear para os próximos 5 dias.
                {byDate.length<3 && <span style={{ color:C.amber }}> Precisão aumenta com mais dias na base.</span>}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function Home() {
  const [tab,         setTab]         = useState('espera')
  const [agendas,     setAgendas]     = useState([])
  const [espera,      setEspera]      = useState([])
  const [loadingDB,   setLoadingDB]   = useState(true)
  const [storing,     setStoring]     = useState(false)
  const [storeMsg,    setStoreMsg]    = useState('')
  const [timestamp,   setTimestamp]   = useState('')
  const [storageInfo, setStorageInfo] = useState({ agendas:0, espera:0 })

  const loadTable = useCallback(async table => {
    const PAGE=1000; let all=[], offset=0, maxTries=50
    while(maxTries-->0) {
      const from=offset, to=offset+PAGE-1
      const res=await fetch(`${SB_URL}/rest/v1/${table}?select=*&order=id.asc`, {
        headers: { ...SBH, 'Range':`${from}-${to}`, 'Range-Unit':'items', 'Prefer':'count=none' }
      })
      if (res.status===416||!res.ok) break
      const batch=await res.json()
      if(!Array.isArray(batch)||!batch.length) break
      all=all.concat(batch)
      if (res.status===200) break
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
      } catch(e){ setStoreMsg(`Erro: ${e.message}`) }
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
    <div style={{ minHeight:'100vh', background:C.bg, fontFamily:"'DM Sans','Segoe UI',sans-serif", color:C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800;900&family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(245,158,11,0.2);border-radius:4px}
        select option{background:#0A0D16;color:#F1F5F9}
        input::placeholder{color:#475569}
        select{appearance:auto}
      `}</style>
      <div style={{ background:'#070910', borderBottom:'1px solid rgba(245,158,11,0.1)', display:'flex', alignItems:'center', padding:'0 32px', height:52, position:'sticky', top:0, zIndex:100, backdropFilter:'blur(12px)' }}>
        <div style={{ position:'absolute', top:-40, left:'50%', transform:'translateX(-50%)', width:500, height:80, background:'radial-gradient(ellipse,rgba(245,158,11,0.07),transparent 70%)', pointerEvents:'none' }} />
        <div style={{ display:'flex', alignItems:'center', gap:10, marginRight:32 }}>
          <div style={{ width:28, height:28, borderRadius:8, background:'linear-gradient(135deg,#F59E0B,#F97316)', display:'flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
          </div>
          <div>
            <span style={{ fontSize:14, fontWeight:900, color:C.text, fontFamily:"'Syne',sans-serif", letterSpacing:'-.3px' }}>Monitor </span>
            <span style={{ fontSize:14, fontWeight:900, color:C.amber, fontFamily:"'Syne',sans-serif", letterSpacing:'-.3px' }}>Clínicas</span>
          </div>
        </div>
        <div style={{ display:'flex', height:'100%', gap:0 }}>
          {[
            { key:'espera',  label:'Fila de Espera',  count:storageInfo.espera  },
            { key:'agendas', label:'Agendas Médicas', count:storageInfo.agendas },
          ].map(t=>(
            <button key={t.key} onClick={()=>setTab(t.key)} style={{
              padding:'0 20px', border:'none', background:'transparent', cursor:'pointer',
              fontSize:12, fontWeight:700, height:'100%',
              color:tab===t.key?C.amber:C.muted,
              borderBottom:tab===t.key?`2px solid ${C.amber}`:'2px solid transparent',
              transition:'all .2s', display:'flex', alignItems:'center', gap:7,
            }}>
              {t.label}
              {t.count>0 && <span style={{ fontSize:9, padding:'1px 6px', borderRadius:10, background:tab===t.key?`${C.amber}20`:'rgba(255,255,255,0.06)', color:tab===t.key?C.amber:C.muted, fontWeight:700 }}>{t.count.toLocaleString('pt-BR')}</span>}
            </button>
          ))}
        </div>
        <div style={{ marginLeft:'auto', display:'flex', alignItems:'center', gap:14 }}>
          <div style={{ display:'flex', alignItems:'center', gap:5 }}>
            <div style={{ width:6, height:6, borderRadius:'50%', background:C.emerald, boxShadow:`0 0 6px ${C.emerald}` }} />
            <span style={{ fontSize:10, color:C.muted }}>Banco conectado</span>
          </div>
          {timestamp && <span style={{ fontSize:10, color:C.muted }}>{timestamp}</span>}
          {storeMsg && <span style={{ fontSize:10, color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?C.emerald:C.amber, fontWeight:600, maxWidth:220, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{storeMsg}</span>}
          {(storageInfo.agendas>0||storageInfo.espera>0)&&!storing && (
            <button onClick={handleClear} style={{ background:'transparent', border:'0.5px solid rgba(244,63,94,0.3)', borderRadius:8, color:C.rose, fontSize:11, padding:'5px 10px', cursor:'pointer' }}>🗑</button>
          )}
          <label style={{ background:storing?'rgba(255,255,255,0.06)':'linear-gradient(135deg,#F59E0B,#F97316)', color:storing?C.muted:'#1a0800', fontWeight:800, fontSize:12, padding:'7px 16px', borderRadius:9, cursor:storing?'default':'pointer', transition:'all .2s', whiteSpace:'nowrap', fontFamily:"'Syne',sans-serif" }}>
            {storing?'Salvando…':'+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={storing||loadingDB} />
          </label>
        </div>
      </div>
      <div style={{ padding:'24px 32px' }}>
        {loadingDB ? (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'calc(100vh - 120px)', gap:16 }}>
            <div style={{ fontSize:36, filter:'drop-shadow(0 0 20px #F59E0B88)' }}>⏳</div>
            <div style={{ fontSize:16, color:C.sub, fontFamily:"'Syne',sans-serif", fontWeight:700 }}>Conectando ao banco…</div>
            {storeMsg && <div style={{ fontSize:12, color:C.amber }}>{storeMsg}</div>}
          </div>
        ) : (<>
          {tab==='agendas' && <TabAgendas rows={agendas} />}
          {tab==='espera'  && <TabEspera  rows={espera}  />}
        </>)}
      </div>
    </div>
  )
}