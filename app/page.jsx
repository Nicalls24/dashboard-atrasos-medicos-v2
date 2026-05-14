'use client'

import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'

// ── THEME ─────────────────────────────────────────────────────────────────────
const T = {
  bg0:     '#03050E',
  bg1:     '#060B1A',
  bg2:     '#080E22',
  glass:   'rgba(255,255,255,0.03)',
  glass2:  'rgba(255,255,255,0.06)',
  border:  'rgba(99,120,255,0.18)',
  border2: 'rgba(99,120,255,0.35)',
  blue:    '#4F7BFF',
  blueB:   '#2952E3',
  blueC:   '#1A3ACC',
  cyan:    '#00D4FF',
  cyanB:   '#00A8CC',
  violet:  '#9B5DFF',
  violetB: '#7B3FE4',
  emerald: '#00E5A0',
  amber:   '#FFB020',
  rose:    '#FF4060',
  text:    '#EEF2FF',
  sub:     '#8898CC',
  muted:   '#3D5080',
  gBlue:   'linear-gradient(135deg,#4F7BFF,#2952E3)',
  gViolet: 'linear-gradient(135deg,#9B5DFF,#7B3FE4)',
  gCyan:   'linear-gradient(135deg,#00D4FF,#007ACC)',
  gEm:     'linear-gradient(135deg,#00E5A0,#00A870)',
  gRose:   'linear-gradient(135deg,#FF4060,#CC1040)',
}

// ── SUPABASE ──────────────────────────────────────────────────────────────────
const SB_URL = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'
const SBH = { 'apikey': SB_KEY, 'Authorization': `Bearer ${SB_KEY}`, 'Content-Type': 'application/json' }

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmtMin = m => {
  if (m === null || m === undefined) return '—'
  const abs = Math.round(Math.abs(m)), sign = m < 0 ? '-' : ''
  if (abs < 60) return `${sign}${abs}min`
  return `${sign}${Math.floor(abs/60)}h${abs%60>0?` ${abs%60}m`:''}`
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
const buildFilter = (allDates, periodo, dateFrom, dateTo) => {
  if (!allDates.length) return () => true
  const sorted = [...allDates].sort(), max = sorted[sorted.length-1]
  if (periodo==='HOJE')   return ds => ds === todayStr()
  if (periodo==='ONTEM')  { const ref=new Date(max+'T00:00:00Z'); ref.setUTCDate(ref.getUTCDate()-1); const s=ref.toISOString().slice(0,10); return ds=>ds===s }
  if (periodo==='SEMANA') { const ref=new Date(max+'T00:00:00Z'); ref.setUTCDate(ref.getUTCDate()-6); const cut=ref.toISOString().slice(0,10); return ds=>ds>=cut&&ds<=max }
  if (periodo==='MES')    return ds => ds.slice(0,7)===max.slice(0,7)
  if (periodo==='ANO')    return ds => ds.slice(0,4)===max.slice(0,4)
  if (periodo==='PERIODO'){ const from=dateFrom||sorted[0],to=dateTo||max; return ds=>ds>=from&&ds<=to }
  return () => true
}
const STATUS_CFG = {
  'OK':             { label:'OK',             color:'#00E5A0', glow:'#00E5A044' },
  'ATRASO':         { label:'Atraso 31–45min',color:'#FFB020', glow:'#FFB02044' },
  'ATRASO CRÍTICO': { label:'Atraso Crítico', color:'#FF6020', glow:'#FF602044' },
  'ATRASO GRAVE':   { label:'Atraso Grave',   color:'#FF4060', glow:'#FF406044' },
  'Falta Médica':   { label:'Médico Faltou',  color:'#4F7BFF', glow:'#4F7BFF44' },
  'SEM PONTO':      { label:'Sem Ponto',      color:'#6B7280', glow:'#6B728044' },
}
const getStatusCfg = (s='') => {
  if (STATUS_CFG[s]) return STATUS_CFG[s]
  const u = s.toUpperCase()
  if (u.includes('CRÍTICO')||u.includes('CRITICO')) return STATUS_CFG['ATRASO CRÍTICO']
  if (u.includes('GRAVE'))  return STATUS_CFG['ATRASO GRAVE']
  if (u.includes('ATRASO')) return STATUS_CFG['ATRASO']
  if (u.includes('FALTA'))  return STATUS_CFG['Falta Médica']
  return { label: s||'OK', color:'#00E5A0', glow:'#00E5A044' }
}
const PERIODOS = [{key:'HOJE',label:'Hoje'},{key:'ONTEM',label:'Ontem'},{key:'SEMANA',label:'Semana'},{key:'MES',label:'Mês'},{key:'ANO',label:'Ano'},{key:'PERIODO',label:'Período'}]

// ── ATOMS ─────────────────────────────────────────────────────────────────────
function GlowCard({ children, color=T.blue, style={} }) {
  return (
    <div style={{
      background: T.glass2,
      border: `1px solid ${color}44`,
      borderRadius: 20,
      padding: 22,
      backdropFilter: 'blur(12px)',
      boxShadow: `0 0 40px ${color}18, inset 0 1px 0 rgba(255,255,255,0.06)`,
      position: 'relative',
      overflow: 'hidden',
      ...style,
    }}>
      <div style={{ position:'absolute', top:0, left:0, right:0, height:1, background:`linear-gradient(90deg,transparent,${color}88,transparent)` }} />
      {children}
    </div>
  )
}

function KpiCard({ icon, label, value, sub, color, grad }) {
  return (
    <GlowCard color={color} style={{ padding:'20px 22px' }}>
      <div style={{ position:'absolute', top:-40, right:-40, width:120, height:120, borderRadius:'50%', background:color, opacity:.08, filter:'blur(20px)' }} />
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:14 }}>
        <div style={{ width:36, height:36, borderRadius:10, background:grad||color, display:'flex', alignItems:'center', justifyContent:'center', fontSize:18, boxShadow:`0 0 16px ${color}66` }}>
          {icon}
        </div>
        <span style={{ fontSize:10, fontWeight:700, color:T.muted, textTransform:'uppercase', letterSpacing:'.12em' }}>{label}</span>
      </div>
      <div style={{ fontSize:40, fontWeight:900, color:T.text, lineHeight:1, letterSpacing:'-2px', fontFamily:"'Clash Display','Syne',sans-serif", textShadow:`0 0 30px ${color}88` }}>
        {value}
      </div>
      {sub && <div style={{ fontSize:11, color:T.sub, marginTop:8, fontWeight:500 }}>{sub}</div>}
      <div style={{ position:'absolute', bottom:0, left:0, right:0, height:2, background:`linear-gradient(90deg,transparent,${color},transparent)`, opacity:.6 }} />
    </GlowCard>
  )
}

function NeonBar({ label, value, max, color, unit='', rank, sub }) {
  const pct = max > 0 ? Math.min((typeof value==='number'?value:0)/max*100,100) : 0
  return (
    <div style={{ marginBottom:14 }}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:6, gap:8 }}>
        <div style={{ display:'flex', alignItems:'flex-start', gap:7, minWidth:0 }}>
          {rank!=null && (
            <span style={{ fontSize:10, fontWeight:900, color: rank===0?T.amber:rank<3?T.blue:T.muted, minWidth:20, flexShrink:0, marginTop:1 }}>
              #{rank+1}
            </span>
          )}
          <div style={{ minWidth:0 }}>
            <div style={{ fontSize:12.5, color:T.text, fontWeight:600, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{label}</div>
            {sub && <div style={{ fontSize:10, color:T.muted, marginTop:1 }}>{sub}</div>}
          </div>
        </div>
        <span style={{ fontSize:12, fontWeight:800, color, flexShrink:0, textShadow:`0 0 12px ${color}` }}>{value}{unit}</span>
      </div>
      <div style={{ background:'rgba(255,255,255,0.05)', borderRadius:6, height:6, overflow:'hidden', position:'relative' }}>
        <div style={{
          height:'100%', borderRadius:6,
          background:`linear-gradient(90deg,${color},${color}88)`,
          width:`${pct}%`,
          transition:'width .8s cubic-bezier(.4,0,.2,1)',
          boxShadow:`0 0 10px ${color}`,
        }} />
      </div>
    </div>
  )
}

function SecTitle({ children, sub }) {
  return (
    <div style={{ marginBottom:18 }}>
      <h2 style={{ fontSize:11.5, fontWeight:800, textTransform:'uppercase', letterSpacing:'.14em', color:T.sub, margin:0 }}>{children}</h2>
      {sub && <div style={{ fontSize:10, color:T.muted, marginTop:3 }}>{sub}</div>}
    </div>
  )
}

function PeriodoBar({ value, onChange, allDates, dateFrom, dateTo, onDateFrom, onDateTo, label }) {
  return (
    <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
      <div style={{ display:'flex', gap:2, background:'rgba(255,255,255,0.04)', border:`1px solid ${T.border}`, borderRadius:12, padding:3, backdropFilter:'blur(8px)' }}>
        {PERIODOS.map(p => (
          <button key={p.key} onClick={()=>onChange(p.key)} style={{
            padding:'6px 13px', borderRadius:9, border:'none', fontSize:11.5, fontWeight:700,
            cursor:'pointer', transition:'all .2s',
            background: value===p.key ? T.gBlue : 'transparent',
            color: value===p.key ? '#fff' : T.muted,
            boxShadow: value===p.key ? `0 0 16px ${T.blue}66` : 'none',
          }}>{p.label}</button>
        ))}
      </div>
      {value==='PERIODO' && (
        <>
          <input type="date" value={dateFrom} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateFrom(e.target.value)}
            style={{ background:'rgba(255,255,255,0.05)', border:`1px solid ${T.border2}`, borderRadius:9, color:T.text, fontSize:11, padding:'6px 10px', outline:'none', colorScheme:'dark' }} />
          <span style={{ color:T.muted }}>→</span>
          <input type="date" value={dateTo} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateTo(e.target.value)}
            style={{ background:'rgba(255,255,255,0.05)', border:`1px solid ${T.border2}`, borderRadius:9, color:T.text, fontSize:11, padding:'6px 10px', outline:'none', colorScheme:'dark' }} />
        </>
      )}
      {label && <span style={{ fontSize:10.5, color:T.muted }}>{label}</span>}
    </div>
  )
}

function Filters({ search, onSearch, uf, onUf, ufs, extra, showClear, onClear }) {
  return (
    <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center', marginBottom:20 }}>
      <input value={search} onChange={e=>onSearch(e.target.value)} placeholder="🔍  Buscar unidade, médico, especialidade…"
        style={{ background:'rgba(255,255,255,0.04)', border:`1px solid ${T.border}`, borderRadius:11, color:T.text, fontSize:12, padding:'8px 14px', outline:'none', width:280, backdropFilter:'blur(8px)' }} />
      <select value={uf} onChange={e=>onUf(e.target.value)}
        style={{ background:'rgba(10,14,35,0.9)', border:`1px solid ${T.border}`, borderRadius:11, color:T.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
        <option value="TODOS">Todos os Estados</option>
        {ufs.map(u=><option key={u}>{u}</option>)}
      </select>
      {extra}
      {showClear && (
        <button onClick={onClear} style={{ background:'rgba(255,64,96,0.1)', border:`1px solid ${T.rose}44`, borderRadius:11, color:T.rose, fontSize:12, padding:'8px 12px', cursor:'pointer' }}>✕ Limpar</button>
      )}
    </div>
  )
}

// Neon line/bar sparkline
function NeonSpark({ data, color, height=60 }) {
  if (!data.length) return null
  const max = Math.max(...data.map(d=>d.v), 1)
  const w = 100 / data.length
  return (
    <div style={{ display:'flex', alignItems:'flex-end', gap:3, height, position:'relative' }}>
      {/* Glow backdrop */}
      <div style={{ position:'absolute', bottom:0, left:0, right:0, height:'40%', background:`linear-gradient(0deg,${color}18,transparent)`, borderRadius:4, pointerEvents:'none' }} />
      {data.map((d,i) => {
        const h = Math.max((d.v/max)*100, 4)
        return (
          <div key={i} title={`${fmtDate(d.k)}: ${d.v}`} style={{
            flex:1, borderRadius:'3px 3px 0 0',
            background:`linear-gradient(0deg,${color},${color}66)`,
            height:`${h}%`,
            boxShadow:`0 0 8px ${color}88`,
            transition:'height .6s cubic-bezier(.4,0,.2,1)',
            opacity: .5 + .5*(d.v/max),
          }} />
        )
      })}
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
  const ufs      = useMemo(() => [...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(), [rows])
  const statuses = useMemo(() => [...new Set(rows.map(r=>r.status).filter(Boolean))].sort(), [rows])

  const filtered = useMemo(() => {
    let r = rows.filter(d=>periodoFn(d.data_agenda))
    if (ufFilt!=='TODOS')     r = r.filter(d=>d.uf===ufFilt)
    if (statusFilt!=='TODOS') r = r.filter(d=>d.status===statusFilt)
    if (search) { const q=search.toLowerCase(); r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))) }
    return r
  }, [rows, periodoFn, ufFilt, statusFilt, search])

  const stats = useMemo(() => {
    const total    = filtered.length
    const emAtraso = filtered.filter(d=>String(d.status||'').toUpperCase().includes('ATRASO')).length
    const semPonto = filtered.filter(d=>d.hr_entrada_min===null||d.hr_entrada_min===undefined).length
    const unidMap={}, medMap={}, statMap={}
    filtered.forEach(d=>{
      const u=d.nm_local||'Sem Unidade'; unidMap[u]=(unidMap[u]||0)+1
      const m=d.nm_medico||''; if(m) medMap[m]=(medMap[m]||0)+1
      const s=d.status||'OK'; statMap[s]=(statMap[s]||0)+1
    })
    const topUnidades = Object.entries(unidMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const topMedicos  = Object.entries(medMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8)
    const statusBreak = Object.entries(statMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)
    const dayMap={}
    filtered.forEach(d=>{ if(d.data_agenda) dayMap[d.data_agenda]=(dayMap[d.data_agenda]||0)+1 })
    const byDate = Object.entries(dayMap).map(([k,v])=>({k,v})).sort((a,b)=>a.k.localeCompare(b.k)).slice(-28)
    return {
      total, emAtraso, semPonto, topUnidades, topMedicos, statusBreak, byDate,
      atrasoPct: total>0?((emAtraso/total)*100).toFixed(1):'0',
      semPontoPct: total>0?((semPonto/total)*100).toFixed(1):'0',
    }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', minHeight:420, gap:16 }}>
      <div style={{ fontSize:56, filter:'drop-shadow(0 0 20px #4F7BFF88)' }}>📋</div>
      <div style={{ fontSize:22, fontWeight:900, color:T.text, fontFamily:"'Syne',sans-serif", letterSpacing:'-0.5px' }}>Nenhuma agenda carregada</div>
      <div style={{ fontSize:13, color:T.muted }}>Carregue uma planilha para começar</div>
    </div>
  )

  const { total, emAtraso, semPonto, topUnidades, topMedicos, statusBreak, byDate, atrasoPct, semPontoPct } = stats
  const maxUnid = topUnidades[0]?.v||1, maxMed = topMedicos[0]?.v||1

  return (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20, flexWrap:'wrap', gap:12 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={`${total.toLocaleString('pt-BR')} registros`} />
      </div>
      <Filters search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||statusFilt!=='TODOS'||!!search}
        onClear={()=>{setUfFilt('TODOS');setStatusFilt('TODOS');setSearch('')}}
        extra={
          <select value={statusFilt} onChange={e=>setStatusFilt(e.target.value)}
            style={{ background:'rgba(10,14,35,0.9)', border:`1px solid ${T.border}`, borderRadius:11, color:T.text, fontSize:12, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
            <option value="TODOS">Todos os Status</option>
            {statuses.map(s=><option key={s}>{s}</option>)}
          </select>
        } />

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:14, marginBottom:20 }}>
        <KpiCard icon="🗓" label="Total Agendas"    value={total.toLocaleString('pt-BR')}         color={T.blue}    grad={T.gBlue} />
        <KpiCard icon="⚠️" label="Em Atraso"        value={emAtraso.toLocaleString('pt-BR')}      sub={`${atrasoPct}% do total`}    color={T.rose}    grad={T.gRose} />
        <KpiCard icon="🚫" label="Sem Ponto"        value={semPonto.toLocaleString('pt-BR')}      sub={`${semPontoPct}% do total`}  color={T.violet}  grad={T.gViolet} />
        <KpiCard icon="✅" label="Com Atendimento"  value={(total-semPonto).toLocaleString('pt-BR')} sub={`${(100-Number(semPontoPct)).toFixed(1)}% presentes`} color={T.emerald} grad={T.gEm} />
      </div>

      {/* Tendência + Status */}
      <div style={{ display:'grid', gridTemplateColumns:'1.4fr 1fr', gap:14, marginBottom:14 }}>
        {byDate.length > 1 && (
          <GlowCard color={T.blue}>
            <SecTitle sub={`Últimos ${byDate.length} dias com dados`}>📈 Agendas por Dia</SecTitle>
            <NeonSpark data={byDate} color={T.blue} height={64} />
            <div style={{ display:'flex', justifyContent:'space-between', marginTop:8 }}>
              <span style={{ fontSize:10, color:T.muted }}>{fmtDate(byDate[0]?.k)} · {byDate[0]?.v} ag.</span>
              <span style={{ fontSize:10, color:T.muted }}>{fmtDate(byDate[byDate.length-1]?.k)} · {byDate[byDate.length-1]?.v} ag.</span>
            </div>
          </GlowCard>
        )}
        <GlowCard color={T.violet}>
          <SecTitle>📊 Status das Agendas</SecTitle>
          <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
            {statusBreak.slice(0,6).map(({k,v}) => {
              const cfg = getStatusCfg(k)
              const pct = total>0?((v/total)*100).toFixed(1):'0'
              return (
                <div key={k} style={{ display:'flex', alignItems:'center', gap:10 }}>
                  <div style={{ width:8, height:8, borderRadius:'50%', background:cfg.color, boxShadow:`0 0 8px ${cfg.color}`, flexShrink:0 }} />
                  <div style={{ flex:1, fontSize:11.5, color:T.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{cfg.label}</div>
                  <span style={{ fontSize:11, fontWeight:800, color:cfg.color, minWidth:36, textAlign:'right', textShadow:`0 0 8px ${cfg.color}` }}>{v.toLocaleString('pt-BR')}</span>
                  <span style={{ fontSize:10, color:T.muted, minWidth:36, textAlign:'right' }}>{pct}%</span>
                  <div style={{ width:60, background:'rgba(255,255,255,0.05)', borderRadius:4, height:4, overflow:'hidden' }}>
                    <div style={{ height:'100%', borderRadius:4, background:cfg.color, width:`${(v/(statusBreak[0]?.v||1))*100}%`, boxShadow:`0 0 6px ${cfg.color}` }} />
                  </div>
                </div>
              )
            })}
          </div>
        </GlowCard>
      </div>

      {/* Rankings */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:14 }}>
        <GlowCard color={T.cyan}>
          <SecTitle>🏥 Top Unidades por Agendas</SecTitle>
          {topUnidades.map((u,i)=>(
            <NeonBar key={u.n} rank={i} label={u.n} value={u.v} max={maxUnid}
              color={i===0?T.amber:i<3?T.cyan:T.blue} />
          ))}
        </GlowCard>
        <GlowCard color={T.violet}>
          <SecTitle>👨‍⚕️ Top Médicos por Agendas</SecTitle>
          {topMedicos.map((m,i)=>(
            <NeonBar key={m.n} rank={i} label={m.n} value={m.v} max={maxMed}
              color={i===0?T.amber:i<3?T.violet:T.blue} />
          ))}
        </GlowCard>
      </div>

      {/* Tabela */}
      <GlowCard color={T.border}>
        <SecTitle sub="Últimos 50 registros filtrados">📋 Detalhamento de Agendas</SecTitle>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11.5 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${T.border}` }}>
                {['Data','Unidade','Médico','Especialidade','UF','Status','Atraso'].map(h=>(
                  <th key={h} style={{ padding:'8px 12px', textAlign:'left', color:T.muted, fontWeight:700, fontSize:9.5, textTransform:'uppercase', letterSpacing:'.08em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0,50).map((r,i)=>{
                const cfg = getStatusCfg(r.status||'')
                return (
                  <tr key={i} style={{ borderBottom:`1px solid rgba(99,120,255,0.07)`, transition:'background .15s' }}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(79,123,255,0.06)'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <td style={{ padding:'9px 12px', color:T.sub, whiteSpace:'nowrap', fontSize:11 }}>{fmtDate(r.data_agenda)}</td>
                    <td style={{ padding:'9px 12px', color:T.text, fontWeight:600, maxWidth:190, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_local}</td>
                    <td style={{ padding:'9px 12px', color:T.sub, whiteSpace:'nowrap' }}>{r.nm_medico}</td>
                    <td style={{ padding:'9px 12px', color:T.muted, maxWidth:150, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.ds_especialidade}</td>
                    <td style={{ padding:'9px 12px', color:T.muted }}>{r.uf}</td>
                    <td style={{ padding:'9px 12px' }}>
                      <span style={{ fontSize:9.5, fontWeight:700, padding:'2px 8px', borderRadius:20, background:cfg.glow, color:cfg.color, border:`1px solid ${cfg.color}44` }}>{r.status||'—'}</span>
                    </td>
                    <td style={{ padding:'9px 12px', color:r.atraso==='SIM'?T.rose:T.emerald, fontWeight:700, textShadow:r.atraso==='SIM'?`0 0 10px ${T.rose}`:'' }}>{r.atraso||'—'}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </GlowCard>
    </div>
  )
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
    if (search) { const q=search.toLowerCase(); r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))) }
    return r
  }, [rows, periodoFn, ufFilt, search])

  const stats = useMemo(() => {
    const comEspera = filtered.filter(d=>d.tempo_espera_min!==null&&d.tempo_espera_min!==undefined)
    const total   = filtered.length
    const totalPac= filtered.reduce((a,d)=>a+(d.qt_pacientes_aguardando||0),0)
    const soma    = comEspera.reduce((a,d)=>a+d.tempo_espera_min,0)
    const media   = comEspera.length>0?soma/comEspera.length:null
    const maxE    = comEspera.length>0?Math.max(...comEspera.map(d=>d.tempo_espera_min)):null
    const unidMap={}, medMap={}
    comEspera.forEach(d=>{
      const u=d.nm_local||'Sem Unidade'; if(!unidMap[u]) unidMap[u]={soma:0,cnt:0}; unidMap[u].soma+=d.tempo_espera_min; unidMap[u].cnt+=1
      const m=d.nm_medico||''; if(!m) return; if(!medMap[m]) medMap[m]={soma:0,cnt:0}; medMap[m].soma+=d.tempo_espera_min; medMap[m].cnt+=1
    })
    const topUnidades = Object.entries(unidMap).map(([n,{soma,cnt}])=>({n,v:Math.round(soma/cnt),cnt})).sort((a,b)=>b.v-a.v).slice(0,8)
    const topMedicos  = Object.entries(medMap).map(([n,{soma,cnt}])=>({n,v:Math.round(soma/cnt),cnt})).sort((a,b)=>b.v-a.v).slice(0,8)
    const dayMap={}
    comEspera.forEach(d=>{ if(!d.data_agenda) return; if(!dayMap[d.data_agenda]) dayMap[d.data_agenda]={soma:0,cnt:0}; dayMap[d.data_agenda].soma+=d.tempo_espera_min; dayMap[d.data_agenda].cnt+=1 })
    const byDate = Object.entries(dayMap).map(([k,{soma,cnt}])=>({k,v:Math.round(soma/cnt)})).sort((a,b)=>a.k.localeCompare(b.k)).slice(-28)
    const faixas = {'0–15min':0,'16–30min':0,'31–60min':0,'61–120min':0,'+120min':0}
    comEspera.forEach(d=>{ const m=d.tempo_espera_min; if(m<=15) faixas['0–15min']++; else if(m<=30) faixas['16–30min']++; else if(m<=60) faixas['31–60min']++; else if(m<=120) faixas['61–120min']++; else faixas['+120min']++ })
    const distEspera = Object.entries(faixas).map(([k,v])=>({k,v}))
    return { total, totalPac, media, maxE, topUnidades, topMedicos, byDate, distEspera }
  }, [filtered])

  if (!rows.length) return (
    <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', minHeight:420, gap:16 }}>
      <div style={{ fontSize:56, filter:'drop-shadow(0 0 20px #00D4FF88)' }}>⏱️</div>
      <div style={{ fontSize:22, fontWeight:900, color:T.text, fontFamily:"'Syne',sans-serif" }}>Nenhum dado de espera carregado</div>
      <div style={{ fontSize:13, color:T.muted }}>Carregue uma planilha para começar</div>
    </div>
  )

  const { total, totalPac, media, maxE, topUnidades, topMedicos, byDate, distEspera } = stats
  const maxUnid = topUnidades[0]?.v||1, maxMed = topMedicos[0]?.v||1
  const distColors = [T.emerald, T.cyan, T.amber, T.rose, '#FF1040']

  return (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20, flexWrap:'wrap', gap:12 }}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('')}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={`${total.toLocaleString('pt-BR')} registros`} />
      </div>
      <Filters search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}} />

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:14, marginBottom:20 }}>
        <KpiCard icon="⏱️" label="Registros de Espera"  value={total.toLocaleString('pt-BR')}    color={T.cyan}    grad={T.gCyan}   />
        <KpiCard icon="👥" label="Pacientes na Fila"    value={totalPac.toLocaleString('pt-BR')} color={T.blue}    grad={T.gBlue}   />
        <KpiCard icon="📊" label="Espera Média"          value={fmtMin(media)}  sub="TEMPO_DE_ESPERA"  color={T.amber}   grad="linear-gradient(135deg,#FFB020,#CC8000)" />
        <KpiCard icon="🔴" label="Maior Espera"          value={fmtMin(maxE)}   sub="pior caso no período" color={T.rose} grad={T.gRose}   />
      </div>

      {/* Distribuição por faixa */}
      <GlowCard color={T.amber} style={{ marginBottom:14 }}>
        <SecTitle sub="Réguas de classificação a definir — distribuição por TEMPO_DE_ESPERA">⏳ Distribuição por Faixa de Espera</SecTitle>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(5,1fr)', gap:10 }}>
          {distEspera.map(({k,v},i) => {
            const c = distColors[i]
            const maxD = Math.max(...distEspera.map(d=>d.v),1)
            return (
              <div key={k} style={{ background:`${c}0A`, border:`1px solid ${c}33`, borderRadius:14, padding:'16px 14px', textAlign:'center', position:'relative', overflow:'hidden' }}>
                <div style={{ position:'absolute', bottom:0, left:0, right:0, height:3, background:`linear-gradient(90deg,${c},${c}44)`, opacity:(v/maxD) }} />
                <div style={{ fontSize:9, fontWeight:800, color:c, textTransform:'uppercase', letterSpacing:'.09em', marginBottom:10, textShadow:`0 0 8px ${c}` }}>{k}</div>
                <div style={{ fontSize:30, fontWeight:900, color:c, lineHeight:1, textShadow:`0 0 20px ${c}88`, fontFamily:"'Syne',sans-serif" }}>{v.toLocaleString('pt-BR')}</div>
              </div>
            )
          })}
        </div>
      </GlowCard>

      {/* Tendência */}
      {byDate.length > 1 && (
        <GlowCard color={T.amber} style={{ marginBottom:14 }}>
          <SecTitle sub={`Média diária de TEMPO_DE_ESPERA — ${byDate.length} dias`}>📈 Evolução da Espera Média</SecTitle>
          <NeonSpark data={byDate} color={T.amber} height={64} />
          <div style={{ display:'flex', justifyContent:'space-between', marginTop:8 }}>
            <span style={{ fontSize:10, color:T.muted }}>{fmtDate(byDate[0]?.k)} · {fmtMin(byDate[0]?.v)}</span>
            <span style={{ fontSize:10, color:T.muted }}>{fmtDate(byDate[byDate.length-1]?.k)} · {fmtMin(byDate[byDate.length-1]?.v)}</span>
          </div>
        </GlowCard>
      )}

      {/* Rankings */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:14 }}>
        <GlowCard color={T.rose}>
          <SecTitle sub="Média de TEMPO_DE_ESPERA">🏥 Unidades — Maior Espera</SecTitle>
          {topUnidades.map((u,i)=>(
            <NeonBar key={u.n} rank={i} label={u.n} value={fmtMin(u.v)} max={maxUnid} sub={`${u.cnt} registros`}
              color={i===0?T.rose:i<3?T.amber:T.cyan} />
          ))}
          {!topUnidades.length && <div style={{color:T.muted,fontSize:12}}>Sem dados no período.</div>}
        </GlowCard>
        <GlowCard color={T.violet}>
          <SecTitle sub="Média de TEMPO_DE_ESPERA">👨‍⚕️ Médicos — Maior Espera</SecTitle>
          {topMedicos.map((m,i)=>(
            <NeonBar key={m.n} rank={i} label={m.n} value={fmtMin(m.v)} max={maxMed} sub={`${m.cnt} registros`}
              color={i<3?T.violet:T.blue} />
          ))}
          {!topMedicos.length && <div style={{color:T.muted,fontSize:12}}>Sem dados no período.</div>}
        </GlowCard>
      </div>

      {/* Tabela */}
      <GlowCard color={T.border}>
        <SecTitle sub="Ordenado por maior TEMPO_DE_ESPERA — top 50">📋 Detalhamento de Espera</SecTitle>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11.5 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${T.border}` }}>
                {['Data','Unidade','Médico','UF','Pac. Aguardando','Tempo de Espera'].map(h=>(
                  <th key={h} style={{ padding:'8px 12px', textAlign:'left', color:T.muted, fontWeight:700, fontSize:9.5, textTransform:'uppercase', letterSpacing:'.08em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...filtered].sort((a,b)=>(b.tempo_espera_min||0)-(a.tempo_espera_min||0)).slice(0,50).map((r,i)=>{
                const m=r.tempo_espera_min
                const c=m>120?T.rose:m>60?T.amber:m>30?T.cyan:T.emerald
                return (
                  <tr key={i} style={{ borderBottom:`1px solid rgba(99,120,255,0.07)`, transition:'background .15s' }}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(79,123,255,0.06)'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <td style={{ padding:'9px 12px', color:T.sub, whiteSpace:'nowrap', fontSize:11 }}>{fmtDate(r.data_agenda)}</td>
                    <td style={{ padding:'9px 12px', color:T.text, fontWeight:600, maxWidth:190, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.nm_local}</td>
                    <td style={{ padding:'9px 12px', color:T.sub, whiteSpace:'nowrap' }}>{r.nm_medico}</td>
                    <td style={{ padding:'9px 12px', color:T.muted }}>{r.uf}</td>
                    <td style={{ padding:'9px 12px', color:T.cyan, fontWeight:700, textAlign:'center' }}>{r.qt_pacientes_aguardando??'—'}</td>
                    <td style={{ padding:'9px 12px' }}>
                      <span style={{ fontSize:13, fontWeight:900, color:c, textShadow:`0 0 12px ${c}88` }}>{fmtMin(m)}</span>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </GlowCard>
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

  const loadTable = useCallback(async (table) => {
    const PAGE=1000; let all=[], offset=0
    while(true) {
      const res = await fetch(`${SB_URL}/rest/v1/${table}?select=*&order=id.asc&limit=${PAGE}&offset=${offset}`,{headers:SBH})
      if(!res.ok) break
      const batch = await res.json()
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
        const [ag,esp] = await Promise.all([loadTable('agendas'),loadTable('espera')])
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

  const handleUpload = useCallback(async (e) => {
    const file=e.target.files[0]; if(!file) return; e.target.value=''
    setStoring(true)
    const now=new Date(),pad=n=>String(n).padStart(2,'0')
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
        setStoreMsg(`Salvando agendas… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`)
        await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:json.slice(i,i+CHUNK),ts,table:'agendas'})})
      }
      for(let i=0;i<json.length;i+=CHUNK){
        setStoreMsg(`Salvando espera… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`)
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
    <div style={{
      minHeight:'100vh',
      fontFamily:"'DM Sans','Segoe UI',sans-serif",
      color:T.text,
      background:`radial-gradient(ellipse 80% 60% at 50% -10%,rgba(79,123,255,0.22) 0%,transparent 60%),
                  radial-gradient(ellipse 60% 40% at 90% 50%,rgba(155,93,255,0.12) 0%,transparent 50%),
                  radial-gradient(ellipse 40% 30% at 10% 80%,rgba(0,212,255,0.08) 0%,transparent 40%),
                  linear-gradient(180deg,#03050E 0%,#060B1A 50%,#040810 100%)`,
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800;900&family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(99,120,255,0.3);border-radius:4px}
        select option{background:#0A0E23;color:#EEF2FF}
        input::placeholder{color:#3D5080}
        select{appearance:auto}
        .upload-btn{transition:all .25s}
        .upload-btn:hover{filter:brightness(1.15);transform:translateY(-2px);box-shadow:0 8px 32px rgba(79,123,255,0.4)!important}
      `}</style>

      {/* HEADER */}
      <div style={{
        backdropFilter:'blur(20px)',
        background:'rgba(6,11,26,0.8)',
        borderBottom:`1px solid ${T.border}`,
        padding:'0 32px',
        display:'flex', alignItems:'center', justifyContent:'space-between',
        height:62, position:'sticky', top:0, zIndex:100,
        boxShadow:`0 1px 0 rgba(99,120,255,0.15), 0 4px 24px rgba(0,0,0,0.4)`,
      }}>
        {/* LOGO */}
        <div style={{ display:'flex', alignItems:'center', gap:12 }}>
          <div style={{
            width:34, height:34, borderRadius:10,
            background:'linear-gradient(135deg,#4F7BFF,#9B5DFF)',
            display:'flex', alignItems:'center', justifyContent:'center', fontSize:17,
            boxShadow:'0 0 20px rgba(79,123,255,0.6)',
          }}>🏥</div>
          <div>
            <div style={{ fontSize:15, fontWeight:900, fontFamily:"'Syne',sans-serif", letterSpacing:'-.4px', background:'linear-gradient(90deg,#EEF2FF,#8898CC)', WebkitBackgroundClip:'text', WebkitTextFillColor:'transparent' }}>
              Monitor Clínicas
            </div>
            <div style={{ fontSize:9.5, color:T.muted, letterSpacing:'.08em', textTransform:'uppercase' }}>Agendas · Espera · Operacional</div>
          </div>
        </div>

        {/* TABS */}
        <div style={{ display:'flex', gap:2, background:'rgba(255,255,255,0.03)', border:`1px solid ${T.border}`, borderRadius:13, padding:4, backdropFilter:'blur(8px)' }}>
          {[{key:'agendas',icon:'🗓',label:'Agendas',count:storageInfo.agendas},{key:'espera',icon:'⏱',label:'Espera',count:storageInfo.espera}].map(t=>(
            <button key={t.key} onClick={()=>setTab(t.key)} style={{
              padding:'7px 22px', borderRadius:10, border:'none', cursor:'pointer', transition:'all .2s',
              fontSize:12.5, fontWeight:700,
              background: tab===t.key ? 'linear-gradient(135deg,#4F7BFF,#2952E3)' : 'transparent',
              color: tab===t.key ? '#fff' : T.muted,
              boxShadow: tab===t.key ? '0 0 20px rgba(79,123,255,0.5)' : 'none',
            }}>
              {t.icon} {t.label}
              {t.count>0&&<span style={{ marginLeft:7, fontSize:9.5, background:'rgba(255,255,255,0.15)', borderRadius:10, padding:'1px 7px' }}>{t.count.toLocaleString('pt-BR')}</span>}
            </button>
          ))}
        </div>

        {/* CONTROLS */}
        <div style={{ display:'flex', alignItems:'center', gap:10 }}>
          {timestamp && (
            <div style={{ fontSize:9, color:T.muted, textAlign:'right', lineHeight:1.6 }}>
              <div style={{ textTransform:'uppercase', letterSpacing:'.08em' }}>Verificação</div>
              <div style={{ color:T.sub, fontWeight:600 }}>{timestamp}</div>
            </div>
          )}
          <div style={{ fontSize:9.5, textAlign:'right', lineHeight:1.7, minWidth:110 }}>
            {(storageInfo.agendas>0||storageInfo.espera>0)&&!storeMsg&&(
              <div style={{ color:T.emerald, fontWeight:700, textShadow:`0 0 10px ${T.emerald}88` }}>
                ☁ {(storageInfo.agendas+storageInfo.espera).toLocaleString('pt-BR')} registros
              </div>
            )}
            {storeMsg&&<div style={{ color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?T.emerald:T.amber, fontWeight:600 }}>{storeMsg}</div>}
          </div>
          {(storageInfo.agendas>0||storageInfo.espera>0)&&!storing&&(
            <button onClick={handleClear} style={{ background:'rgba(255,64,96,0.08)', border:`1px solid rgba(255,64,96,0.3)`, borderRadius:9, color:T.rose, fontSize:12, padding:'6px 10px', cursor:'pointer', transition:'all .2s' }}>🗑</button>
          )}
          <label className="upload-btn" style={{
            background: storing ? 'rgba(255,255,255,0.08)' : 'linear-gradient(135deg,#4F7BFF,#2952E3)',
            color:'#fff', fontWeight:700, fontSize:12.5, padding:'8px 18px',
            borderRadius:10, cursor:storing?'default':'pointer',
            whiteSpace:'nowrap', boxShadow:'0 0 24px rgba(79,123,255,0.35)',
            border:'1px solid rgba(79,123,255,0.4)',
          }}>
            {storing?'Salvando…':'+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={storing||loadingDB} />
          </label>
        </div>
      </div>

      {/* CONTENT */}
      <div style={{ padding:'24px 32px' }}>
        {loadingDB ? (
          <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', minHeight:'calc(100vh - 110px)', gap:16 }}>
            <div style={{ fontSize:48, filter:'drop-shadow(0 0 30px #4F7BFF)' }}>⏳</div>
            <div style={{ fontSize:18, color:T.sub, fontFamily:"'Syne',sans-serif", fontWeight:800 }}>Conectando ao banco de dados…</div>
            {storeMsg&&<div style={{ fontSize:12, color:T.amber, background:'rgba(255,176,32,0.08)', padding:'10px 18px', borderRadius:10, border:`1px solid ${T.amber}44` }}>{storeMsg}</div>}
          </div>
        ) : (
          <>
            {tab==='agendas'&&<TabAgendas rows={agendas}/>}
            {tab==='espera' &&<TabEspera  rows={espera} />}
          </>
        )}
      </div>
    </div>
  )
}