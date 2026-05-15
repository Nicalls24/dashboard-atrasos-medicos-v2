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
// Baseado em TEMPO_DE_ESPERA (coluna AB)
// HR_REGISTRO_ESPERA → apenas para saber a hora do registro
// < 15min    → ignorar
// 15–30min   → Espera Moderada
// 31–89min   → Espera Grave
// >= 90min   → Espera Crítica
const CLS_ESPERA = {
  get: m => {
    if (!m || m < 15) return null
    if (m <= 30) return { key:'MOD',  label:'Espera Moderada', color:'#F59E0B', bg:'rgba(245,158,11,0.08)',  border:'rgba(245,158,11,0.22)' }
    if (m <= 89) return { key:'GRV',  label:'Espera Grave',    color:'#F97316', bg:'rgba(249,115,22,0.1)',   border:'rgba(249,115,22,0.28)' }
    return             { key:'CRIT', label:'Espera Crítica',  color:'#F43F5E', bg:'rgba(244,63,94,0.1)',    border:'rgba(244,63,94,0.28)'  }
  }
}

// ── TAB ESPERA ────────────────────────────────────────────────────────────────
function TabEspera({ rows }) {
  const [periodo,  setPeriodo]  = useState('MES')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo,   setDateTo]   = useState('')
  const [ufFilt,   setUfFilt]   = useState('TODOS')
  const [search,   setSearch]   = useState('')
  const [horaFilt, setHoraFilt] = useState('TODAS')

  const allDates = useMemo(() =>
    [...new Set(rows.map(r => r.data_agenda).filter(Boolean))].sort()
  , [rows])

  const periodoFn = useMemo(() =>
    buildFilter(allDates, periodo, dateFrom, dateTo)
  , [allDates, periodo, dateFrom, dateTo])

  const ufs = useMemo(() =>
    [...new Set(rows.map(r => r.uf).filter(Boolean))].sort()
  , [rows])

  // Horas disponíveis baseadas em HR_REGISTRO_ESPERA (somente para esperas >= 15min)
  const horasDisp = useMemo(() => {
    const set = new Set()
    rows.filter(d => periodoFn(d.data_agenda) && d.tempo_espera_min >= 15)
      .forEach(d => {
        const h = d.hr_registro_espera_min
        if (h === null || h === undefined) return
        const hora = Math.floor(h / 60)
        if (hora >= 0 && hora <= 23) set.add(hora)
      })
    return [...set].sort((a,b) => a-b)
  }, [rows, periodoFn])

  const filtered = useMemo(() => {
    let r = rows.filter(d => periodoFn(d.data_agenda))
    if (ufFilt !== 'TODOS') r = r.filter(d => d.uf === ufFilt)
    if (search) {
      const q = search.toLowerCase()
      r = r.filter(d => [d.nm_local, d.nm_medico, d.cidade].some(v => String(v||'').toLowerCase().includes(q)))
    }
    // Filtro por hora: aplica sobre HR_REGISTRO_ESPERA (apenas para esperas >= 15min)
    // Registros sem espera (< 15min ou null) são mantidos para os KPIs totais
    // mas o feed filtra por hora internamente
    return r
  }, [rows, periodoFn, ufFilt, search])

  const stats = useMemo(() => {
    // Apenas registros com TEMPO_DE_ESPERA >= 15min
    const comEspera = filtered.filter(d =>
      d.tempo_espera_min !== null && d.tempo_espera_min !== undefined && d.tempo_espera_min >= 15
    )

    const total     = filtered.length
    const totalPac  = filtered.reduce((a, d) => a + (d.qt_pacientes_aguardando || 0), 0)
    const moderada  = comEspera.filter(d => d.tempo_espera_min <= 30).length
    const grave     = comEspera.filter(d => d.tempo_espera_min > 30 && d.tempo_espera_min <= 89).length
    const critico   = comEspera.filter(d => d.tempo_espera_min >= 90).length
    const totalEsp  = moderada + grave + critico

    // Feed de alertas:
    // Agrupa por hora (HR_REGISTRO_ESPERA) + unidade
    // Pega o pior TEMPO_DE_ESPERA de cada grupo
    // HR_REGISTRO_ESPERA → apenas para definir a hora do evento
    const grupoMap = {}
    comEspera.forEach(d => {
      const h = d.hr_registro_espera_min
      if (h === null || h === undefined) return
      const hora    = Math.floor(h / 60)
      const minutos = Math.floor(h % 60)
      if (hora < 0 || hora > 23) return
      const horaStr = String(hora).padStart(2,'0') + ':' + String(minutos).padStart(2,'0')
      const unidade = d.nm_local || 'Sem Unidade'
      const key     = `${horaStr}||${unidade}`

      if (!grupoMap[key]) {
        grupoMap[key] = {
          horaStr,
          hora,
          minutos,
          nm_local:  unidade,
          nm_medico: d.nm_medico || '—',
          uf:        d.uf || '—',
          cidade:    d.cidade || '',
          maxTempo:  0,
          totalPac:  0,
          count:     0,
        }
      }
      if (d.tempo_espera_min > grupoMap[key].maxTempo) {
        grupoMap[key].maxTempo  = d.tempo_espera_min
        grupoMap[key].nm_medico = d.nm_medico || '—'
      }
      grupoMap[key].totalPac += d.qt_pacientes_aguardando || 0
      grupoMap[key].count    += 1
    })

    // Ordena: críticos primeiro, depois graves, depois moderados
    // Dentro de cada categoria: maior tempo primeiro
    const feed = Object.values(grupoMap)
      .sort((a, b) => b.maxTempo - a.maxTempo)
      .slice(0, 25)

    // Top unidades por esperas críticas
    const unidMap = {}
    comEspera.forEach(d => {
      const u = d.nm_local || 'Sem Unidade'
      if (!unidMap[u]) unidMap[u] = { n: u, total: 0, criticos: 0, graves: 0, moderadas: 0, pac: 0 }
      unidMap[u].total++
      unidMap[u].pac += d.qt_pacientes_aguardando || 0
      const m = d.tempo_espera_min
      if (m >= 90)      unidMap[u].criticos++
      else if (m >= 31) unidMap[u].graves++
      else              unidMap[u].moderadas++
    })
    const topUnidades = Object.values(unidMap)
      .sort((a, b) => b.criticos - a.criticos || b.graves - a.graves)
      .slice(0, 6)

    // Faltas e atrasos médicos
    const faltas       = filtered.filter(d => String(d.atraso||'').toUpperCase() === 'FALTA')
    const atrasosReais = filtered.filter(d => {
      if (String(d.atraso||'').toUpperCase() !== 'SIM') return false
      const t = d.tempo_atraso_min
      return t !== null && t !== undefined && Math.abs(t) > 31
    })
    const sMap = {}
    atrasosReais.forEach(d => { const s = d.status||'Sem Status'; sMap[s] = (sMap[s]||0)+1 })
    const statusAtraso = Object.entries(sMap).map(([k,v]) => ({k,v})).sort((a,b) => b.v-a.v)

    return {
      total, totalPac, moderada, grave, critico, totalEsp,
      feed, topUnidades, faltas, atrasosReais, statusAtraso
    }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'60vh', gap:14 }}>
      <div style={{ fontSize:44 }}>⏱️</div>
      <div style={{ fontSize:18, fontWeight:700, color:C.text, fontFamily:"'Syne',sans-serif" }}>Nenhum dado de espera</div>
      <div style={{ fontSize:13, color:C.muted }}>Use o menu lateral para carregar uma planilha</div>
    </div>
  )

  const {
    total, totalPac, moderada, grave, critico, totalEsp,
    feed, topUnidades, faltas, atrasosReais, statusAtraso
  } = stats

  const fmtMin = m => {
    if (!m) return '—'
    const abs = Math.round(Math.abs(m))
    if (abs < 60) return `${abs}min`
    return `${Math.floor(abs/60)}h${abs%60>0?` ${abs%60}m`:''}`
  }

  return (
    <div>
      {/* Periodo + Filtros */}
      <div style={{ marginBottom:16 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo}
          onDateFrom={setDateFrom} onDateTo={setDateTo} total={total} />
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}} />

      {/* ROW 1 — KPIs + Distribuição */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr 1fr 1.4fr', gap:12, marginBottom:14 }}>
        {[
          { label:'Pacientes na Fila',    value:totalPac.toLocaleString('pt-BR'), sub:`${total} registros`,  color:'#00C9A7' },
          { label:'Espera Moderada',      value:moderada,  sub:'15 – 30 min',                               color:'#F59E0B' },
          { label:'Espera Grave',         value:grave,     sub:'31 min – 1h29',                             color:'#F97316' },
          { label:'Espera Crítica',       value:critico,   sub:'acima de 1h30',                             color:'#F43F5E' },
        ].map(k => (
          <div key={k.label} style={{
            background:'rgba(255,255,255,0.025)', border:`1px solid ${k.color}20`,
            borderRadius:14, padding:'18px 18px 14px', position:'relative', overflow:'hidden',
          }}>
            <div style={{ position:'absolute', top:0, left:0, right:0, height:2, background:`linear-gradient(90deg,${k.color},transparent)` }} />
            <div style={{ fontSize:9.5, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:12 }}>{k.label}</div>
            <div style={{ fontSize:28, fontWeight:800, color:k.color, letterSpacing:'-1px', lineHeight:1 }}>
              {typeof k.value === 'number' ? k.value.toLocaleString('pt-BR') : k.value}
            </div>
            <div style={{ fontSize:10, color:C.muted, marginTop:7 }}>{k.sub}</div>
          </div>
        ))}

        {/* Distribuição empilhada */}
        <div style={{
          background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)',
          borderRadius:14, padding:'18px',
        }}>
          <div style={{ fontSize:9.5, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:12 }}>Distribuição</div>
          {totalEsp > 0 ? (<>
            <div style={{ display:'flex', height:8, borderRadius:4, overflow:'hidden', marginBottom:12, gap:2 }}>
              {moderada>0&&<div style={{ flex:moderada, background:'#F59E0B' }} />}
              {grave>0&&<div style={{ flex:grave, background:'#F97316' }} />}
              {critico>0&&<div style={{ flex:critico, background:'#F43F5E' }} />}
            </div>
            {[
              { label:'Moderada', v:moderada, color:'#F59E0B' },
              { label:'Grave',    v:grave,    color:'#F97316' },
              { label:'Crítica',  v:critico,  color:'#F43F5E' },
            ].map(x => (
              <div key={x.label} style={{ display:'flex', alignItems:'center', gap:7, marginBottom:7 }}>
                <div style={{ width:7, height:7, borderRadius:2, background:x.color, flexShrink:0 }} />
                <span style={{ fontSize:10.5, color:C.sub, flex:1 }}>{x.label}</span>
                <span style={{ fontSize:11, fontWeight:700, color:x.color }}>{x.v}</span>
                <span style={{ fontSize:9.5, color:C.muted, minWidth:28, textAlign:'right' }}>
                  {totalEsp>0?Math.round((x.v/totalEsp)*100):0}%
                </span>
              </div>
            ))}
          </>) : <div style={{ color:C.muted, fontSize:11 }}>Sem esperas no período</div>}
        </div>
      </div>

      {/* ROW 2 — Feed de Alertas + Top Unidades */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 320px', gap:14, marginBottom:14 }}>

        {/* FEED DE ALERTAS */}
        <div style={{
          background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)',
          borderRadius:14, padding:'20px 22px',
        }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:16 }}>
            <div>
              <div style={{ fontSize:13, fontWeight:700, color:C.text }}>Feed de Esperas por Hora</div>
              <div style={{ fontSize:10.5, color:C.muted, marginTop:3 }}>
                Classificado por TEMPO_DE_ESPERA · hora via HR_REGISTRO_ESPERA · ordenado por gravidade
              </div>
            </div>
            <div style={{ display:'flex', alignItems:'center', gap:12, flexShrink:0 }}>
              {/* Filtro por hora — baseado em HR_REGISTRO_ESPERA, só esperas >= 15min */}
              <select value={horaFilt} onChange={e=>setHoraFilt(e.target.value)} style={{
                background:'rgba(255,255,255,0.05)', border:`0.5px solid rgba(245,158,11,0.2)`,
                borderRadius:8, color:C.text, fontSize:11, padding:'5px 10px',
                outline:'none', cursor:'pointer',
              }}>
                <option value="TODAS">Todas as horas</option>
                {horasDisp.map(h => (
                  <option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>
                ))}
              </select>
              <div style={{ display:'flex', gap:10 }}>
                {[
                  { color:'#F59E0B', label:'Moderada' },
                  { color:'#F97316', label:'Grave'    },
                  { color:'#F43F5E', label:'Crítica'  },
                ].map(l => (
                  <div key={l.label} style={{ display:'flex', alignItems:'center', gap:4 }}>
                    <div style={{ width:6, height:6, borderRadius:'50%', background:l.color }} />
                    <span style={{ fontSize:9, color:C.muted }}>{l.label}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {(() => {
            // Aplica filtro de hora no feed (HR_REGISTRO_ESPERA)
            const feedFiltrado = horaFilt === 'TODAS'
              ? feed
              : feed.filter(item => item.hora === Number(horaFilt))
            if (feedFiltrado.length === 0) return (
              <div style={{ textAlign:'center', padding:'32px 0', color:C.muted, fontSize:12 }}>
                {horaFilt !== 'TODAS'
                  ? `Sem esperas ≥ 15min às ${String(horaFilt).padStart(2,'0')}:00.`
                  : 'Sem esperas ≥ 15min no período selecionado.'}
              </div>
            )
            return (
            <div style={{ display:'flex', flexDirection:'column', gap:5, maxHeight:420, overflowY:'auto' }}>
              {feedFiltrado.map((item, i) => {
                const cls = CLS_ESPERA.get(item.maxTempo)
                const isCrit = item.maxTempo >= 90
                const isGrv  = item.maxTempo >= 31 && item.maxTempo < 90
                return (
                  <div key={i} style={{
                    display:'flex', alignItems:'center', gap:12,
                    padding:'10px 14px', borderRadius:10,
                    background: isCrit ? 'rgba(244,63,94,0.06)' : isGrv ? 'rgba(249,115,22,0.04)' : 'rgba(255,255,255,0.02)',
                    border: `0.5px solid ${i < 3 ? cls.border : 'rgba(255,255,255,0.05)'}`,
                    transition:'background .12s', cursor:'default',
                  }}
                    onMouseEnter={e => e.currentTarget.style.background = cls.bg}
                    onMouseLeave={e => e.currentTarget.style.background = isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':'rgba(255,255,255,0.02)'}
                  >
                    {/* Dot */}
                    <div style={{
                      width:8, height:8, borderRadius:'50%', background:cls.color, flexShrink:0,
                      boxShadow: isCrit ? `0 0 8px ${cls.color}` : 'none',
                    }} />

                    {/* Hora */}
                    <div style={{ fontFamily:'monospace', fontSize:13, fontWeight:700, color:C.sub, flexShrink:0, minWidth:44 }}>
                      {item.horaStr}
                    </div>

                    {/* Unidade + médico */}
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontSize:12, fontWeight:600, color:C.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                        {item.nm_local}
                      </div>
                      <div style={{ fontSize:10, color:C.muted, marginTop:2, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                        {[item.cidade, item.uf].filter(Boolean).join(' · ')}
                      </div>
                    </div>

                    {/* Pacientes */}
                    <div style={{ textAlign:'center', flexShrink:0, minWidth:48 }}>
                      <div style={{ fontSize:12, fontWeight:700, color:'#0EA5E9' }}>
                        {item.totalPac > 0 ? item.totalPac.toLocaleString('pt-BR') : '—'}
                      </div>
                      <div style={{ fontSize:9, color:C.muted }}>pac.</div>
                    </div>

                    {/* Tempo */}
                    <div style={{ fontSize:15, fontWeight:900, color:cls.color, flexShrink:0, minWidth:52, textAlign:'right', letterSpacing:'-.3px' }}>
                      {fmtMin(item.maxTempo)}
                    </div>

                    {/* Badge classificação */}
                    <span style={{
                      fontSize:9.5, fontWeight:700, padding:'3px 9px', borderRadius:20,
                      background:cls.bg, color:cls.color, border:`0.5px solid ${cls.border}`,
                      whiteSpace:'nowrap', flexShrink:0,
                    }}>
                      {cls.label}
                    </span>
                  </div>
                )
              })}
            </div>
            )
          })()}
        </div>

        {/* TOP UNIDADES */}
        <div style={{
          background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)',
          borderRadius:14, padding:'20px 22px', display:'flex', flexDirection:'column',
        }}>
          <div style={{ fontSize:13, fontWeight:700, color:C.text, marginBottom:4 }}>Top Unidades</div>
          <div style={{ fontSize:10.5, color:C.muted, marginBottom:16 }}>Maior volume de esperas críticas</div>
          {topUnidades.length === 0 && <div style={{ color:C.muted, fontSize:11 }}>Sem dados no período.</div>}
          {topUnidades.map((u, i) => (
            <div key={u.n} style={{ marginBottom:14 }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:5 }}>
                <div style={{ display:'flex', alignItems:'center', gap:7, minWidth:0 }}>
                  <span style={{ fontSize:10, fontWeight:800, color:i<3?'#F59E0B':C.muted, minWidth:16, flexShrink:0 }}>#{i+1}</span>
                  <span style={{ fontSize:11.5, color:C.text, fontWeight:500, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{u.n}</span>
                </div>
                <div style={{ display:'flex', gap:7, flexShrink:0, fontSize:10.5 }}>
                  {u.criticos>0&&<span style={{ color:'#F43F5E', fontWeight:700 }}>🔴{u.criticos}</span>}
                  {u.graves>0&&<span style={{ color:'#F97316', fontWeight:700 }}>🟠{u.graves}</span>}
                  {u.moderadas>0&&<span style={{ color:'#F59E0B', fontWeight:700 }}>🟡{u.moderadas}</span>}
                </div>
              </div>
              <div style={{ display:'flex', height:5, borderRadius:3, overflow:'hidden', gap:1 }}>
                {u.moderadas>0&&<div style={{ flex:u.moderadas, background:'#F59E0B' }} />}
                {u.graves>0&&<div style={{ flex:u.graves, background:'#F97316' }} />}
                {u.criticos>0&&<div style={{ flex:u.criticos, background:'#F43F5E' }} />}
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* ROW 3 — Médicos complemento */}
      <div style={{
        background:'rgba(255,255,255,0.025)', border:'1px solid rgba(255,255,255,0.06)',
        borderRadius:14, padding:'20px 22px', marginBottom:14,
      }}>
        <div style={{ fontSize:13, fontWeight:700, color:C.text, marginBottom:4 }}>Médicos — Falta e Atraso</div>
        <div style={{ fontSize:10.5, color:C.muted, marginBottom:16 }}>Complemento · coluna ATRASO da base</div>
        <div style={{ display:'grid', gridTemplateColumns:'120px 120px 1fr', gap:12, alignItems:'start' }}>
          {[
            { label:'Faltas',         value:faltas.length,        color:'#F43F5E' },
            { label:'Atrasos >31min', value:atrasosReais.length,  color:'#F59E0B' },
          ].map(k => (
            <div key={k.label} style={{ background:`${k.color}08`, border:`0.5px solid ${k.color}22`, borderRadius:11, padding:'12px 14px' }}>
              <div style={{ fontSize:9, fontWeight:700, color:k.color, textTransform:'uppercase', letterSpacing:'.1em', marginBottom:7 }}>{k.label}</div>
              <div style={{ fontSize:26, fontWeight:800, color:k.color }}>{k.value.toLocaleString('pt-BR')}</div>
            </div>
          ))}
          {statusAtraso.length > 0 && (
            <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
              {statusAtraso.map(({k,v}) => {
                const cfg = getStatusCfg(k)
                return (
                  <div key={k} style={{ display:'flex', alignItems:'center', gap:10 }}>
                    <span style={{ fontSize:10, fontWeight:700, padding:'2px 9px', borderRadius:20, background:cfg.glow, color:cfg.color, border:`0.5px solid ${cfg.color}30`, whiteSpace:'nowrap', minWidth:110 }}>{cfg.label}</span>
                    <div style={{ flex:1, background:'rgba(255,255,255,0.05)', borderRadius:3, height:5, overflow:'hidden' }}>
                      <div style={{ height:'100%', background:cfg.color, width:`${(v/(statusAtraso[0]?.v||1))*100}%`, transition:'width .6s ease' }} />
                    </div>
                    <span style={{ fontSize:11, fontWeight:700, color:cfg.color, minWidth:28, textAlign:'right' }}>{v}</span>
                  </div>
                )
              })}
            </div>
          )}
          {!statusAtraso.length&&faltas.length===0&&atrasosReais.length===0&&(
            <div style={{ color:C.muted, fontSize:11 }}>Nenhuma ocorrência no período.</div>
          )}
        </div>
      </div>
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

      {/* HEADER — Estilo C: wide, tabs com underline amber */}
      <div style={{
        background:'#070910',
        borderBottom:`1px solid rgba(245,158,11,0.1)`,
        display:'flex', alignItems:'center',
        padding:'0 32px', height:52, position:'sticky', top:0, zIndex:100,
        backdropFilter:'blur(12px)',
      }}>
        {/* Glow top */}
        <div style={{ position:'absolute', top:-40, left:'50%', transform:'translateX(-50%)', width:500, height:80, background:'radial-gradient(ellipse,rgba(245,158,11,0.07),transparent 70%)', pointerEvents:'none' }} />

        {/* Logo */}
        <div style={{ display:'flex', alignItems:'center', gap:10, marginRight:32 }}>
          <div style={{ width:28, height:28, borderRadius:8, background:'linear-gradient(135deg,#F59E0B,#F97316)', display:'flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
          </div>
          <div>
            <span style={{ fontSize:14, fontWeight:900, color:C.text, fontFamily:"'Syne',sans-serif", letterSpacing:'-.3px' }}>Monitor </span>
            <span style={{ fontSize:14, fontWeight:900, color:C.amber, fontFamily:"'Syne',sans-serif", letterSpacing:'-.3px' }}>Clínicas</span>
          </div>
        </div>

        {/* Tabs com underline */}
        <div style={{ display:'flex', height:'100%', gap:0 }}>
          {[
            { key:'espera',  label:'Fila de Espera',   count:storageInfo.espera  },
            { key:'agendas', label:'Agendas Médicas',  count:storageInfo.agendas },
          ].map(t => (
            <button key={t.key} onClick={()=>setTab(t.key)} style={{
              padding:'0 20px', border:'none', background:'transparent', cursor:'pointer',
              fontSize:12, fontWeight:700, height:'100%',
              color: tab===t.key ? C.amber : C.muted,
              borderBottom: tab===t.key ? `2px solid ${C.amber}` : '2px solid transparent',
              transition:'all .2s', display:'flex', alignItems:'center', gap:7,
            }}>
              {t.label}
              {t.count>0 && (
                <span style={{ fontSize:9, padding:'1px 6px', borderRadius:10, background:tab===t.key?`${C.amber}20`:'rgba(255,255,255,0.06)', color:tab===t.key?C.amber:C.muted, fontWeight:700 }}>
                  {t.count.toLocaleString('pt-BR')}
                </span>
              )}
            </button>
          ))}
        </div>

        {/* Right */}
        <div style={{ marginLeft:'auto', display:'flex', alignItems:'center', gap:14 }}>
          {/* Status banco */}
          <div style={{ display:'flex', alignItems:'center', gap:5 }}>
            <div style={{ width:6, height:6, borderRadius:'50%', background:C.emerald, boxShadow:`0 0 6px ${C.emerald}` }} />
            <span style={{ fontSize:10, color:C.muted }}>Banco conectado</span>
          </div>

          {/* Timestamp */}
          {timestamp && <span style={{ fontSize:10, color:C.muted }}>{timestamp}</span>}

          {/* Store msg */}
          {storeMsg && (
            <span style={{ fontSize:10, color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?C.emerald:C.amber, fontWeight:600, maxWidth:220, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
              {storeMsg}
            </span>
          )}

          {/* Clear */}
          {(storageInfo.agendas>0||storageInfo.espera>0) && !storing && (
            <button onClick={handleClear} style={{ background:'transparent', border:`0.5px solid rgba(244,63,94,0.3)`, borderRadius:8, color:C.rose, fontSize:11, padding:'5px 10px', cursor:'pointer' }}>
              🗑
            </button>
          )}

          {/* Upload */}
          <label style={{
            background: storing ? 'rgba(255,255,255,0.06)' : 'linear-gradient(135deg,#F59E0B,#F97316)',
            color: storing ? C.muted : '#1a0800',
            fontWeight:800, fontSize:12, padding:'7px 16px',
            borderRadius:9, cursor:storing?'default':'pointer', transition:'all .2s',
            whiteSpace:'nowrap', fontFamily:"'Syne',sans-serif",
          }}>
            {storing ? 'Salvando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={storing||loadingDB} />
          </label>
        </div>
      </div>

      {/* CONTENT */}
      <div style={{ padding:'24px 32px' }}>
        {loadingDB ? (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', height:'calc(100vh - 120px)', gap:16 }}>
            <div style={{ fontSize:36, filter:'drop-shadow(0 0 20px #F59E0B88)' }}>⏳</div>
            <div style={{ fontSize:16, color:C.sub, fontFamily:"'Syne',sans-serif", fontWeight:700 }}>Conectando ao banco…</div>
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
  )
}