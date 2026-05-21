'use client'

import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'

const C = {
  bg:'#06080F', amber:'#F59E0B', orange:'#F97316', rose:'#F43F5E', teal:'#00C9A7',
  blue:'#3B82F6', violet:'#8B5CF6', emerald:'#10B981', cyan:'#06B6D4',
  text:'#F1F5F9', sub:'#94A3B8', muted:'#475569',
  border:'rgba(245,158,11,0.1)', border2:'rgba(245,158,11,0.22)',
}
const SB_URL='https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY='eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'
const SBH={'apikey':SB_KEY,'Authorization':`Bearer ${SB_KEY}`,'Content-Type':'application/json'}

const fmtMin=m=>{if(m==null)return'—';const abs=Math.round(Math.abs(m)),s=m<0?'-':'';if(abs<60)return`${s}${abs}min`;return`${s}${Math.floor(abs/60)}h${abs%60>0?` ${abs%60}m`:''}`}
const norm10=d=>d?String(d).slice(0,10):''
const buildFilter=(allDates,periodo,df,dt)=>{
  if(!allDates.length)return()=>true
  const s=[...allDates].sort(),max=s[s.length-1]
  const sysToday=()=>{const n=new Date();return n.getUTCFullYear()+'-'+String(n.getUTCMonth()+1).padStart(2,'0')+'-'+String(n.getUTCDate()).padStart(2,'0')}
  const sysYest=()=>{const n=new Date();n.setUTCDate(n.getUTCDate()-1);return n.toISOString().slice(0,10)}
  if(periodo==='HOJE')return d=>norm10(d)===sysToday()
  if(periodo==='ONTEM')return d=>norm10(d)===sysYest()
  if(periodo==='SEMANA'){const r=new Date(max+'T00:00:00Z');r.setUTCDate(r.getUTCDate()-6);const c=r.toISOString().slice(0,10);return d=>norm10(d)>=c&&norm10(d)<=max}
  if(periodo==='MES')return d=>norm10(d).slice(0,7)===max.slice(0,7)
  if(periodo==='ANO')return d=>norm10(d).slice(0,4)===max.slice(0,4)
  if(periodo==='PERIODO'){const from=df||s[0],to=dt||max;return d=>norm10(d)>=from&&norm10(d)<=to}
  return()=>true
}
const getStatusCfg=s=>{
  const cfg={
    'OK':{label:'OK',color:'#10B981'},
    'ATRASO':{label:'Atraso 31-45min',color:'#F59E0B'},
    'ATRASO CRÍTICO':{label:'Atraso Crítico',color:'#F97316'},
    'ATRASO GRAVE':{label:'Atraso Grave',color:'#F43F5E'},
    'Falta Médica':{label:'Médico Faltou',color:'#3B82F6'},
    'SEM PONTO':{label:'Sem Ponto',color:'#64748B'},
  }
  if(!s)return cfg['OK'];if(cfg[s])return cfg[s]
  const u=s.toUpperCase()
  if(u.includes('CRÍTICO')||u.includes('CRITICO'))return cfg['ATRASO CRÍTICO']
  if(u.includes('GRAVE'))return cfg['ATRASO GRAVE']
  if(u.includes('ATRASO'))return cfg['ATRASO']
  if(u.includes('FALTA'))return cfg['Falta Médica']
  if(u.includes('SEM PONTO')||u.includes('SEM_PONTO'))return cfg['SEM PONTO']
  return cfg['OK']
}
const clsEspera=m=>{
  if(!m||m<15)return null
  if(m<=30)return{label:'Espera Moderada',color:'#F59E0B',bg:'rgba(245,158,11,0.08)',border:'rgba(245,158,11,0.22)'}
  if(m<=89)return{label:'Espera Grave',color:'#F97316',bg:'rgba(249,115,22,0.1)',border:'rgba(249,115,22,0.28)'}
  return{label:'Espera Crítica',color:'#F43F5E',bg:'rgba(244,63,94,0.1)',border:'rgba(244,63,94,0.28)'}
}
const PERIODOS=[{key:'HOJE',label:'Hoje'},{key:'ONTEM',label:'Ontem'},{key:'SEMANA',label:'Semana'},{key:'MES',label:'Mês'},{key:'ANO',label:'Ano'},{key:'PERIODO',label:'Período'}]

function PeriodoBar({value,onChange,allDates,dateFrom,dateTo,onDateFrom,onDateTo,label}){
  return(
    <div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap'}}>
      <div style={{display:'flex',background:'rgba(255,255,255,0.03)',border:`1px solid ${C.border}`,borderRadius:10,padding:3,gap:2}}>
        {PERIODOS.map(p=><button key={p.key} onClick={()=>onChange(p.key)} style={{padding:'5px 12px',borderRadius:7,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .2s',background:value===p.key?C.amber:'transparent',color:value===p.key?'#1a0800':C.muted}}>{p.label}</button>)}
      </div>
      {value==='PERIODO'&&(<>
        <input type="date" value={dateFrom} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateFrom(e.target.value)} style={{background:'rgba(255,255,255,0.04)',border:`1px solid ${C.border2}`,borderRadius:8,color:C.text,fontSize:11,padding:'5px 9px',outline:'none',colorScheme:'dark'}}/>
        <span style={{color:C.muted}}>→</span>
        <input type="date" value={dateTo} min={allDates[0]||''} max={allDates[allDates.length-1]||''} onChange={e=>onDateTo(e.target.value)} style={{background:'rgba(255,255,255,0.04)',border:`1px solid ${C.border2}`,borderRadius:8,color:C.text,fontSize:11,padding:'5px 9px',outline:'none',colorScheme:'dark'}}/>
      </>)}
      {label&&<span style={{fontSize:11,color:C.muted}}>{label}</span>}
    </div>
  )
}

function SearchBar({search,onSearch,uf,onUf,ufs,showClear,onClear}){
  return(
    <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center',marginBottom:18}}>
      <div style={{position:'relative',flex:1,minWidth:220}}>
        <span style={{position:'absolute',left:12,top:'50%',transform:'translateY(-50%)',color:C.muted,fontSize:14}}>🔍</span>
        <input value={search} onChange={e=>onSearch(e.target.value)} placeholder="Buscar unidade, médico, especialidade…" style={{width:'100%',background:'rgba(255,255,255,0.04)',border:`1px solid ${C.border}`,borderRadius:10,color:C.text,fontSize:12,padding:'8px 12px 8px 34px',outline:'none'}}/>
      </div>
      <select value={uf} onChange={e=>onUf(e.target.value)} style={{background:'rgba(6,8,15,0.95)',border:`1px solid ${C.border}`,borderRadius:10,color:C.text,fontSize:12,padding:'8px 12px',outline:'none',cursor:'pointer'}}>
        <option value="TODOS">Todos os Estados</option>
        {ufs.map(u=><option key={u}>{u}</option>)}
      </select>
      {showClear&&<button onClick={onClear} style={{background:'rgba(244,63,94,0.08)',border:`1px solid rgba(244,63,94,0.25)`,borderRadius:10,color:C.rose,fontSize:12,padding:'8px 12px',cursor:'pointer'}}>✕ Limpar</button>}
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB AGENDAS
// ══════════════════════════════════════════════════════════════════════════════
const fmtHHMM=m=>{if(m==null)return'—';const h=Math.floor(m/60),min=m%60;return String(h).padStart(2,'0')+':'+String(min).padStart(2,'0')}

function TabAgendas({rows}){
  const[periodo,setPeriodo]=useState('MES')
  const[dateFrom,setDateFrom]=useState('')
  const[dateTo,setDateTo]=useState('')
  const[ufFilt,setUfFilt]=useState('TODOS')
  const[search,setSearch]=useState('')
  const[horasFilt,setHorasFilt]=useState([])
  const[unidFilt,setUnidFilt]=useState('')
  const[trendView,setTrendView]=useState('real')
  const[aTip,setATip]=useState(null)
  const[allAgJust,setAllAgJust]=useState([])
  const[agJustDraft,setAgJustDraft]=useState({})
  const[agJustSaving,setAgJustSaving]=useState(false)
  const[justExpanded,setJustExpanded]=useState({})
  const[expandedHoras,setExpandedHoras]=useState({}) // accordion: {hora: bool}

  const allDates=useMemo(()=>[...new Set(rows.map(r=>norm10(r.dt_registro||r.data_agenda)).filter(Boolean))].sort(),[rows])
  const periodoFn=useMemo(()=>buildFilter(allDates,periodo,dateFrom,dateTo),[allDates,periodo,dateFrom,dateTo])
  const ufs=useMemo(()=>[...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(),[rows])
  const dataRefAg=useMemo(()=>allDates.length?allDates[allDates.length-1]:'',[allDates])

  useEffect(()=>{
    fetch(`${SB_URL}/rest/v1/justif_agenda?select=data_ref,nm_local,nm_medico,hora,texto`,{headers:SBH})
      .then(r=>r.json()).then(d=>{if(Array.isArray(d))setAllAgJust(d)}).catch(()=>{})
  },[rows])

  const saveAgJust=useCallback(async(nm_medico,hora,texto)=>{
    if(!unidFilt||!dataRefAg)return
    setAgJustSaving(true)
    const horaVal=hora!=null?parseInt(hora):-1
    try{
      await fetch(`${SB_URL}/rest/v1/justif_agenda`,{
        method:'POST',headers:{...SBH,'Prefer':'resolution=merge-duplicates'},
        body:JSON.stringify({data_ref:dataRefAg,nm_local:unidFilt,nm_medico,hora:horaVal,texto})
      })
      setAllAgJust(prev=>{
        const sem=prev.filter(j=>!(j.data_ref===dataRefAg&&j.nm_local===unidFilt&&j.nm_medico===nm_medico&&j.hora===horaVal))
        return[...sem,{data_ref:dataRefAg,nm_local:unidFilt,nm_medico,hora:horaVal,texto}]
      })
    }catch(e){console.error(e)}
    setAgJustSaving(false)
  },[unidFilt,dataRefAg])

  const horasDisp=useMemo(()=>{
    const set=new Set()
    rows.filter(d=>periodoFn(d.dt_registro||d.data_agenda)).forEach(d=>{
      const h=d.hr_inicio_min;if(h==null)return
      const hora=Math.floor(h/60);if(hora>=0&&hora<=23)set.add(hora)
    })
    return[...set].sort((a,b)=>a-b)
  },[rows,periodoFn])

  const toggleHora=useCallback(h=>setHorasFilt(prev=>prev.includes(h)?prev.filter(x=>x!==h):[...prev,h]),[])

  const filtered=useMemo(()=>{
    let r=rows.filter(d=>periodoFn(d.dt_registro||d.data_agenda))
    if(ufFilt!=='TODOS')r=r.filter(d=>d.uf===ufFilt)
    if(search){const q=search.toLowerCase();r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q)))}
    if(horasFilt.length>0)r=r.filter(d=>{const h=d.hr_inicio_min;if(h==null)return false;return horasFilt.includes(Math.floor(h/60))})
    if(unidFilt)r=r.filter(d=>d.nm_local===unidFilt)
    return r
  },[rows,periodoFn,ufFilt,search,horasFilt,unidFilt])

  // filteredGlobal: igual a filtered mas SEM unidFilt → para % global do período
  const filteredGlobal=useMemo(()=>{
    let r=rows.filter(d=>periodoFn(d.dt_registro||d.data_agenda))
    if(ufFilt!=='TODOS')r=r.filter(d=>d.uf===ufFilt)
    if(search){const q=search.toLowerCase();r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q)))}
    // SEM unidFilt e SEM horasFilt — % global do período
    return r
  },[rows,periodoFn,ufFilt,search,horasFilt])

  // justAgGlobal: todas as justificativas do período (sem filtro de unidade)
  const justAgGlobal=useMemo(()=>allAgJust.filter(x=>periodoFn(x.data_ref)),[allAgJust,periodoFn])

  // % global: total de (médico+hora+unidade) com problema vs com justificativa no período
  const globalJustStats=useMemo(()=>{
    const statusU=d=>String(d.status||'').toUpperCase()
    const atrasoU=d=>String(d.atraso||'').toUpperCase()
    const isFn=d=>atrasoU(d)==='FALTA'||statusU(d).includes('FALTA')||atrasoU(d)==='SIM'
    const totalSet=new Set()
    filteredGlobal.forEach(d=>{
      if(!isFn(d))return
      const nm=d.nm_medico;if(!nm)return
      const h=d.hr_inicio_min!=null?Math.floor(d.hr_inicio_min/60):-1
      totalSet.add(`${nm}::${h}::${d.nm_local||''}`)
    })
    const justSet=new Set(justAgGlobal.map(j=>`${j.nm_medico}::${j.hora!=null?j.hora:-1}::${j.nm_local||''}`))
    const justified=[...totalSet].filter(k=>justSet.has(k)).length
    const total=totalSet.size
    const pct=total>0?Math.min(Math.round(justified/total*100),100):0
    return{pct,justified,total}
  },[filteredGlobal,justAgGlobal])

  const globalJustColor=globalJustStats.pct>=80?C.emerald:globalJustStats.pct>=50?C.amber:C.rose

    const docHoraMap=useMemo(()=>{
    const map={}
    filtered.forEach(d=>{
      if(!d.nm_medico)return
      const h=d.hr_inicio_min
      if(h!=null&&map[d.nm_medico]==null)map[d.nm_medico]=Math.floor(h/60)
    })
    return map
  },[filtered])

  const justAgFiltered=useMemo(()=>{
    let j=allAgJust.filter(x=>periodoFn(x.data_ref))
    if(unidFilt)j=j.filter(x=>x.nm_local===unidFilt)
    if(horasFilt.length>0)j=j.filter(x=>x.hora===-1||horasFilt.includes(parseInt(x.hora)))
    return j
  },[allAgJust,periodoFn,unidFilt,horasFilt])

  const docsJustSet=useMemo(()=>{
    const s=new Set()
    justAgFiltered.forEach(j=>{const hora=j.hora===-1?null:j.hora;s.add(j.nm_medico+'::'+String(hora??'x'))})
    return s
  },[justAgFiltered])

  // ── Helpers para TEMPO DE ATRASO (funções puras — sem hook) ────────────────
  const parseTempoMinFn=(t)=>{
    if(t==null||t===''||t===false)return null
    const s=String(t).trim()
    if(!s||s==='None'||s==='0'||s==='false')return null
    if(s.includes(':')){
      const neg=s.startsWith('-')
      const clean=s.replace(/^[-+]/,'')
      const pts=clean.split(':')
      const mins=parseInt(pts[0]||'0')*60+parseInt(pts[1]||'0')
      return neg?mins:-mins // negativo no arquivo = atraso → positivo retornado
    }
    return null
  }
  const fmtTempoMinFn=(m)=>{
    if(!m||m<=0)return null
    if(m<60)return m+'min'
    return Math.floor(m/60)+'h'+(m%60>0?' '+m%60+'m':'')
  }

    const agStats=useMemo(()=>{
    const totalRows=filtered.length
    const totalCons=filtered.reduce((a,d)=>a+(d.qt_consulta||0),0)
    const totalEnc =filtered.reduce((a,d)=>a+(d.qt_encaixe ||0),0)
    const totalAg  =totalCons+totalEnc

    // Classificação por status e atraso
    const statusU=d=>String(d.status||'').toUpperCase()
    const atrasoU=d=>String(d.atraso||'').toUpperCase()
    const isFaltaFn =d=>atrasoU(d)==='FALTA'||statusU(d).includes('FALTA')
    const isCritFn  =d=>atrasoU(d)==='SIM'&&(statusU(d).includes('CRÍTICO')||statusU(d).includes('CRITICO'))
    const isGrvFn   =d=>atrasoU(d)==='SIM'&&statusU(d).includes('GRAVE')
    const isAtrSoFn =d=>atrasoU(d)==='SIM'&&!isCritFn(d)&&!isGrvFn(d)&&!statusU(d).includes('SEM PONTO')

    const faltaRows=filtered.filter(isFaltaFn)
    const critRows =filtered.filter(isCritFn)
    const grvRows  =filtered.filter(isGrvFn)
    const atrRows  =filtered.filter(isAtrSoFn)
    const faltaDocs=[...new Set(faltaRows.map(d=>d.nm_medico).filter(Boolean))]
    const critDocs =[...new Set(critRows .map(d=>d.nm_medico).filter(Boolean))]
    const grvDocs  =[...new Set(grvRows  .map(d=>d.nm_medico).filter(Boolean))]
    const atrDocs  =[...new Set(atrRows  .map(d=>d.nm_medico).filter(Boolean))]
    const faltaAg=faltaRows.reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const critAg =critRows .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const grvAg  =grvRows  .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const atrAg  =atrRows  .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)

    const uMap={}
    filtered.forEach(d=>{
      const u=d.nm_local||'?'
      if(!uMap[u])uMap[u]={nm_local:u,uf:d.uf||'',cidade:d.cidade||'',consultas:0,encaixe:0,total:0,faltas:0,atrasos:0}
      uMap[u].consultas+=(d.qt_consulta||0);uMap[u].encaixe+=(d.qt_encaixe||0);uMap[u].total+=(d.qt_consulta||0)+(d.qt_encaixe||0)
      if(isFaltaFn(d))uMap[u].faltas++;if(atrasoU(d)==='SIM')uMap[u].atrasos++
    })
    const feedList=Object.values(uMap).sort((a,b)=>b.total-a.total).slice(0,20)

    const sMap={}
    filtered.forEach(d=>{
      const s=d.ds_especialidade||'Não informado'
      if(!sMap[s])sMap[s]={name:s,total:0}
      sMap[s].total+=(d.qt_consulta||0)+(d.qt_encaixe||0)
    })
    const topSpec=Object.values(sMap).sort((a,b)=>b.total-a.total).slice(0,8)

    const dMap={}
    filtered.forEach(d=>{
      const dt=norm10(d.data_agenda);if(!dt)return
      if(!dMap[dt])dMap[dt]={date:dt,agendas:0,consultas:0,encaixe:0,impactadas:0}
      dMap[dt].agendas+=(d.qt_consulta||0)+(d.qt_encaixe||0)
      dMap[dt].consultas+=(d.qt_consulta||0);dMap[dt].encaixe+=(d.qt_encaixe||0)
      if(isFaltaFn(d)||atrasoU(d)==='SIM')dMap[dt].impactadas+=(d.qt_consulta||0)+(d.qt_encaixe||0)
    })
    const byDate=Object.values(dMap).map(d=>({...d,delivered:Math.max(0,d.agendas-d.impactadas)})).sort((a,b)=>a.date.localeCompare(b.date))
    const addDay=(ds,n)=>{const d=new Date(ds+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)}
    const pts=byDate.slice(-3)
    const mkSlope=(key)=>pts.length>=2?(pts[pts.length-1][key]-pts[0][key])/(pts.length-1):0
    const lastDate=byDate[byDate.length-1]?.date||''
    const projData=lastDate?Array.from({length:5},(_,i)=>({
      date:addDay(lastDate,i+1),
      agendas:Math.max(0,Math.round((byDate[byDate.length-1]?.agendas||0)+mkSlope('agendas')*(i+1))),
      impactadas:Math.max(0,Math.round((byDate[byDate.length-1]?.impactadas||0)+mkSlope('impactadas')*(i+1))),
      delivered:Math.max(0,Math.round((byDate[byDate.length-1]?.delivered||0)+mkSlope('delivered')*(i+1))),
      consultas:0,encaixe:0,isProj:true,
    })):[]

    const docsFaltaU=unidFilt?faltaDocs.map(nm=>({nm,status:'Falta Médica',color:C.blue})):[]
    const docsCritU =unidFilt?critDocs .map(nm=>({nm,status:'Atraso Crítico',color:C.rose})):[]
    const docsGrvU  =unidFilt?grvDocs  .map(nm=>({nm,status:'Atraso Grave',color:C.orange})):[]
    const docsAtrU  =unidFilt?atrDocs  .map(nm=>({nm,status:'Atraso 31–45min',color:C.amber})):[]

    // ── SITUAÇÃO DE PONTO ─────────────────────────────────────────────────────
    // Col O (hr_entrada): null = sem ponto | preenchida = chegou
    // Col S (atraso): 'FALTA' | 'SIM' | 'NÃO'
    // Col T (tempo_de_atraso): '-00:31:00' string (negativo = atrasado)
    // Col U (status): 'Falta Médica' | 'Remarcação Adm' | 'ATRASO' | 'ATRASO GRAVE' | 'ATRASO CRÍTICO'
    //
    // Visão 1: hr_entrada=null + atraso='FALTA' → faltou (mostra STATUS como motivo)
    // Visão 2: hr_entrada=null + atraso='SIM'   → sem ponto em atraso (mostra TEMPO + classificação)
    // Visão 3: hr_entrada=preenchida + atraso='SIM' → com ponto atrasado (mostra TEMPO + STATUS)

    // parseTempoMin e fmtTempoMin definidos como parseTempoMinFn/fmtTempoMinFn no escopo do componente

    const rowsPonto=rows
      .filter(d=>periodoFn(d.dt_registro||d.data_agenda))
      .filter(d=>ufFilt==='TODOS'||d.uf===ufFilt)
      .filter(d=>{if(!search)return true;const q=search.toLowerCase();return[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q))})
      .filter(d=>{if(horasFilt.length===0)return true;const h=d.hr_inicio_min;if(h==null)return false;return horasFilt.includes(Math.floor(h/60))})

    const semPontoFaltaMap  = {} // hora → [{nm, status}]
    const semPontoAtrasoMap = {} // hora → [{nm, tempo_min, status}]
    const comPontoAtrasoList= [] // [{nm, tempo_min, status, hora}]

    rowsPonto.forEach(d=>{
      const hrRaw=d.hr_entrada??d.HR_ENTRADA??null
      const temPonto=hrRaw!==null&&hrRaw!==undefined&&String(hrRaw).trim()!==''&&String(hrRaw).trim()!=='0'&&hrRaw!==0
      const atrasoU=String(d.atraso||'').toUpperCase().trim()
      const isFalta=atrasoU==='FALTA'
      const isSIM  =atrasoU==='SIM'
      if(!isFalta&&!isSIM)return

      const h=d.hr_inicio_min!=null?Math.floor(d.hr_inicio_min/60):null
      if(h==null||h<0||h>23)return

      const nm=d.nm_medico||'—'
      const tempo_min=parseTempoMinFn(d.tempo_de_atraso??d['tempo_de_atraso']??d['TEMPO DE ATRASO']??null)
      const status=d.status||''

      if(!temPonto){
        // SEM PONTO
        const hr_ini_min=d.hr_inicio_min??null
        const hr_fim_min=d.hr_fim_min??d.hr_fim??null
        if(isFalta){
          if(!semPontoFaltaMap[h])semPontoFaltaMap[h]=[]
          if(!semPontoFaltaMap[h].find(x=>x.nm===nm&&x.nm_local===(d.nm_local||'')))
            semPontoFaltaMap[h].push({nm,status,nm_local:d.nm_local||'',hr_ini_min,hr_fim_min})
        }else if(isSIM){
          if(!semPontoAtrasoMap[h])semPontoAtrasoMap[h]=[]
          if(!semPontoAtrasoMap[h].find(x=>x.nm===nm&&x.nm_local===(d.nm_local||'')))
            semPontoAtrasoMap[h].push({nm,tempo_min,status,nm_local:d.nm_local||'',hr_ini_min,hr_fim_min})
        }
      }else{
        // COM PONTO + ATRASO
        if(isSIM&&!comPontoAtrasoList.find(x=>x.nm===nm&&x.hora===h&&x.nm_local===(d.nm_local||'')))
          comPontoAtrasoList.push({nm,tempo_min,status,hora:h,nm_local:d.nm_local||'',hr_ini_min,hr_fim_min:d.hr_fim_min??d.hr_fim??null})
      }
    })

    const allHorasPonto=[...new Set([
      ...Object.keys(semPontoFaltaMap),
      ...Object.keys(semPontoAtrasoMap),
      ...comPontoAtrasoList.map(d=>String(d.hora))
    ])].map(Number).filter(h=>!isNaN(h)).sort((a,b)=>a-b)

    const situacaoPontoHoras=allHorasPonto.map(hora=>({
      hora,
      semPontoFalta :[...(semPontoFaltaMap [hora]||[])],
      semPontoAtraso:[...(semPontoAtrasoMap[hora]||[])],
      comPontoAtraso:[...new Map(comPontoAtrasoList.filter(d=>d.hora===hora).map(d=>[d.nm,d])).values()],
    }))
    const semPontoFaltaTotal =[...new Set(Object.values(semPontoFaltaMap ).flatMap(arr=>arr.map(x=>x.nm)))].length
    const semPontoAtrasoTotal=[...new Set(Object.values(semPontoAtrasoMap).flatMap(arr=>arr.map(x=>x.nm)))].length
    const comPontoAtrasoTotal=[...new Set(comPontoAtrasoList.map(d=>d.nm))].length

        return{totalRows,totalCons,totalEnc,totalAg,
      faltaDocs,critDocs,grvDocs,atrDocs,faltaAg,critAg,grvAg,atrAg,
      feedList,topSpec,byDate,projData,
      docsFaltaU,docsCritU,docsGrvU,docsAtrU,
      situacaoPontoHoras,semPontoAtrasoTotal,semPontoFaltaTotal,comPontoAtrasoTotal}
  },[filtered,unidFilt,rows,periodoFn,ufFilt,search,horasFilt])

  const{totalRows,totalCons,totalEnc,totalAg,
    faltaDocs,critDocs,grvDocs,atrDocs,faltaAg,critAg,grvAg,atrAg,
    feedList,topSpec,byDate,projData,
    docsFaltaU,docsCritU,docsGrvU,docsAtrU,
    situacaoPontoHoras,semPontoAtrasoTotal,semPontoFaltaTotal,comPontoAtrasoTotal}=agStats

  const allDocsList=useMemo(()=>[...docsFaltaU,...docsAtrU,...docsGrvU,...docsCritU],[docsFaltaU,docsAtrU,docsGrvU,docsCritU])
  // justPct, justColor, justifiedCount removidos → usar globalJustStats e globalJustColor

  const trendAllData=trendView==='real'?byDate:[...byDate,...projData]
  const tMaxAg=Math.max(...trendAllData.map(d=>d.agendas||0),1)
  const tRealCnt=byDate.length
  const tLastReal=byDate[byDate.length-1]
  const tTotalGap=byDate.reduce((a,d)=>a+(d.impactadas||0),0)
  const tAvgAg=byDate.length>0?Math.round(byDate.reduce((a,d)=>a+(d.agendas||0),0)/byDate.length):0
  const tSlope=projData.length>0?((projData[0]?.agendas||0)-(tLastReal?.agendas||0)):0

  const AVW=800,AVH=150,APL=52,APR=20,APT=10,APB=32
  const ACW=AVW-APL-APR,ACH=AVH-APT-APB
  const aNT=trendAllData.length
  const atxP=i=>aNT<=1?APL+ACW/2:APL+i/(aNT-1)*ACW
  const atyS=v=>APT+ACH-(tMaxAg>0?(v||0)/tMaxAg*ACH:0)
  const atSmooth=(data,key,scFn,off)=>{
    if(!data.length)return''
    const pts=data.map((d,i)=>({x:parseFloat(atxP(i+(off||0)).toFixed(2)),y:parseFloat(scFn(d[key]||0).toFixed(2))}))
    if(pts.length===1)return'M'+pts[0].x+','+pts[0].y
    let p='M'+pts[0].x+','+pts[0].y
    for(let i=1;i<pts.length;i++){
      const p0=pts[i-2]||pts[i-1],p1=pts[i-1],p2=pts[i],p3=pts[i+1]||pts[i]
      const c1x=(p1.x+(p2.x-p0.x)/4).toFixed(2),c1y=(p1.y+(p2.y-p0.y)/4).toFixed(2)
      const c2x=(p2.x-(p3.x-p1.x)/4).toFixed(2),c2y=(p2.y-(p3.y-p1.y)/4).toFixed(2)
      p+=' C'+c1x+','+c1y+' '+c2x+','+c2y+' '+p2.x+','+p2.y
    }
    return p
  }
  const maxFeed=feedList[0]?.total||1
  const card={background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}

  if(!rows.length)return(
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'60vh',gap:14}}>
      <div style={{fontSize:44}}>📋</div>
      <div style={{fontSize:18,fontWeight:700,color:C.text}}>Nenhuma agenda carregada</div>
    </div>
  )

  return(
    <div>
      <div style={{marginBottom:14}}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('');setUnidFilt('');setHorasFilt([])}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={totalRows.toLocaleString('pt-BR')+' registros'}/>
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search||!!unidFilt}
        onClear={()=>{setUfFilt('TODOS');setSearch('');setUnidFilt('')}}/>

      {/* Filtro de hora início removido */}

      {/* KPI BAR */}
      <div style={{display:'flex',background:'rgba(255,255,255,0.025)',border:'0.5px solid rgba(255,255,255,0.07)',borderRadius:12,overflow:'hidden',marginBottom:14}}>
        <div style={{flex:'1.4',padding:'16px 20px',borderRight:'0.5px solid rgba(255,255,255,0.07)'}}>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:5}}>Total Agendas</div>
          <div style={{fontSize:30,fontWeight:800,color:C.amber,letterSpacing:'-1px',lineHeight:1}}>{totalAg.toLocaleString('pt-BR')}</div>
          <div style={{display:'flex',gap:14,marginTop:8}}>
            <div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:18,height:3,background:C.teal,borderRadius:2}}/><span style={{fontSize:9,color:C.muted}}>{totalCons.toLocaleString('pt-BR')} consultas</span></div>
            <div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:18,height:3,background:C.violet,borderRadius:2}}/><span style={{fontSize:9,color:C.muted}}>{totalEnc.toLocaleString('pt-BR')} encaixe</span></div>
          </div>
          <div style={{marginTop:8,height:4,background:'rgba(255,255,255,0.04)',borderRadius:2,overflow:'hidden',display:'flex',gap:1}}>
            <div style={{flex:totalCons,background:C.teal,opacity:.7}}/><div style={{flex:totalEnc,background:C.violet,opacity:.7}}/>
          </div>
        </div>
        <div style={{flex:'1',padding:'16px 20px',borderRight:'0.5px solid rgba(255,255,255,0.07)'}}>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:5}}>Afetadas por problema</div>
          <div style={{fontSize:30,fontWeight:800,color:C.rose,letterSpacing:'-1px',lineHeight:1}}>{(faltaAg+atrAg+critAg+grvAg).toLocaleString('pt-BR')}</div>
          <div style={{fontSize:9,color:C.muted,marginTop:6}}>{totalAg>0?((faltaAg+atrAg+critAg+grvAg)/totalAg*100).toFixed(1):'0'}% das agendas</div>
          <div style={{marginTop:8,height:4,background:'rgba(255,255,255,0.04)',borderRadius:2,overflow:'hidden'}}>
            <div style={{width:(totalAg>0?((faltaAg+atrAg+critAg+grvAg)/totalAg*100).toFixed(1):0)+'%',height:'100%',background:C.rose,opacity:.7,borderRadius:2}}/>
          </div>
        </div>
        <div style={{flex:'1.8',padding:'16px 24px',position:'relative',overflow:'hidden'}}>
          <div style={{position:'absolute',top:-20,right:-20,width:120,height:120,borderRadius:'50%',background:'radial-gradient(circle,rgba(16,185,129,0.06),transparent 70%)',pointerEvents:'none'}}/>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.12em',marginBottom:6}}>Médicos com Ocorrências</div>
          <div style={{fontSize:42,fontWeight:900,color:C.emerald,lineHeight:1,letterSpacing:'-2px'}}>{faltaDocs.length+atrDocs.length+critDocs.length+grvDocs.length}</div>
          <div style={{fontSize:9,color:C.muted,marginTop:8}}>médicos afetados no período</div>
          <div style={{marginTop:10,display:'flex',gap:10}}>
            <div style={{display:'flex',alignItems:'center',gap:4}}><div style={{width:6,height:6,borderRadius:'50%',background:C.blue}}/><span style={{fontSize:9,color:C.muted}}>{faltaDocs.length} falta{faltaDocs.length!==1?'s':''}</span></div>
            <div style={{display:'flex',alignItems:'center',gap:4}}><div style={{width:6,height:6,borderRadius:'50%',background:C.rose}}/><span style={{fontSize:9,color:C.muted}}>{critDocs.length+grvDocs.length+atrDocs.length} atraso{(critDocs.length+grvDocs.length+atrDocs.length)!==1?'s':''}</span></div>
          </div>
        </div>
      </div>

      {/* FEED + PAINEL DIREITO EXECUTIVO */}
      <div style={{display:'grid',gridTemplateColumns:'1fr 460px',gap:14,marginBottom:14}}>
        {/* Feed */}
        <div style={card}>
          {unidFilt&&(
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12,padding:'8px 12px',background:'rgba(245,158,11,0.08)',border:'0.5px solid rgba(245,158,11,0.25)',borderRadius:9}}>
              <div style={{width:6,height:6,borderRadius:'50%',background:C.amber}}/><span style={{fontSize:11,color:C.amber,flex:1,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Filtrando: {unidFilt}</span>
              <button onClick={()=>setUnidFilt('')} style={{background:'transparent',border:'none',color:C.muted,cursor:'pointer',fontSize:12}}>✕</button>
            </div>
          )}
          <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:4}}>Unidades por volume de agendas</div>
          <div style={{fontSize:10,color:C.muted,marginBottom:14}}>{unidFilt?'clique ✕ para limpar':'clique para filtrar e justificar médicos'}</div>
          <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:500,overflowY:'auto'}}>
            {feedList.slice(0,15).map((item,i)=>{
              const isSel=unidFilt===item.nm_local,pct=maxFeed>0?(item.total/maxFeed)*100:0
              return(
                <div key={i} onClick={()=>setUnidFilt(isSel?'':item.nm_local)} style={{display:'flex',alignItems:'center',gap:10,padding:'9px 12px',borderRadius:10,cursor:'pointer',background:isSel?'rgba(245,158,11,0.1)':'rgba(255,255,255,0.02)',border:isSel?'1px solid rgba(245,158,11,0.4)':'0.5px solid rgba(255,255,255,0.05)',transition:'all .15s'}}
                  onMouseEnter={e=>{if(!isSel)e.currentTarget.style.background='rgba(255,255,255,0.04)'}}
                  onMouseLeave={e=>{if(!isSel)e.currentTarget.style.background='rgba(255,255,255,0.02)'}}>
                  <div style={{fontSize:10,fontWeight:800,color:i<3?C.amber:C.muted,minWidth:22}}>#{ i+1}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:11,fontWeight:600,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{item.nm_local}</div>
                    <div style={{height:3,background:'rgba(255,255,255,0.06)',borderRadius:2,marginTop:4,overflow:'hidden'}}>
                      <div style={{height:'100%',borderRadius:2,background:i<3?C.amber:C.teal,width:pct+'%',transition:'width .5s'}}/>
                    </div>
                  </div>
                  <div style={{textAlign:'right',flexShrink:0}}>
                    <div style={{fontSize:12,fontWeight:700,color:i<3?C.amber:C.text}}>{item.total.toLocaleString('pt-BR')}</div>
                    <div style={{fontSize:8,color:C.muted}}>{item.consultas}c+{item.encaixe}e</div>
                  </div>
                  <div style={{display:'flex',flexDirection:'column',gap:2,flexShrink:0}}>
                    {item.faltas>0&&<span style={{fontSize:8,padding:'1px 5px',borderRadius:10,background:'rgba(59,130,246,0.15)',color:C.blue,border:'0.5px solid rgba(59,130,246,0.3)',whiteSpace:'nowrap'}}>{item.faltas}f</span>}
                    {item.atrasos>0&&<span style={{fontSize:8,padding:'1px 5px',borderRadius:10,background:'rgba(244,63,94,0.12)',color:C.rose,border:'0.5px solid rgba(244,63,94,0.3)',whiteSpace:'nowrap'}}>{item.atrasos}a</span>}
                  </div>
                </div>
              )
            })}
          </div>
        </div>

        {/* ── PAINEL DIREITO EXECUTIVO ── */}
        <div style={{display:'flex',flexDirection:'column',gap:12}}>

          {/* BLOCO 1: KPIs grandes — Médicos Impacto */}
          <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:14,padding:'20px',position:'relative',overflow:'hidden'}}>
            {/* Glow de fundo */}
            <div style={{position:'absolute',top:-40,right:-40,width:180,height:180,borderRadius:'50%',background:'radial-gradient(circle,rgba(244,63,94,0.06),transparent 70%)',pointerEvents:'none'}}/>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.14em',marginBottom:16}}>⚡ Médicos — Impacto nas Agendas</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
              {[
                {label:'Falta Médica',docs:faltaDocs.length,ag:faltaAg,color:C.blue,icon:'🔵'},
                {label:'Atraso',docs:atrDocs.length,ag:atrAg,color:C.amber,icon:'🟡'},
                {label:'Atraso Grave',docs:grvDocs.length,ag:grvAg,color:C.orange,icon:'🟠'},
                {label:'Atraso Crítico',docs:critDocs.length,ag:critAg,color:C.rose,icon:'🔴'},
              ].map(k=>(
                <div key={k.label} style={{borderRadius:12,padding:'14px 16px',background:`linear-gradient(145deg,${k.color}14,${k.color}06)`,border:`1px solid ${k.color}30`,position:'relative',overflow:'hidden'}}>
                  <div style={{fontSize:8,fontWeight:700,color:k.color,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:10}}>{k.label}</div>
                  <div style={{fontSize:44,fontWeight:900,color:k.color,lineHeight:1,letterSpacing:'-2px'}}>{k.docs}</div>
                  <div style={{fontSize:9,color:C.muted,marginTop:6}}>médico{k.docs!==1?'s':''}</div>
                  <div style={{marginTop:10,paddingTop:8,borderTop:`0.5px solid ${k.color}20`,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <span style={{fontSize:8,color:C.muted}}>agendas</span>
                    <span style={{fontSize:13,fontWeight:800,color:k.color}}>{k.ag.toLocaleString('pt-BR')}</span>
                  </div>
                  <div style={{position:'absolute',bottom:-8,right:4,fontSize:48,opacity:0.05,fontWeight:900,color:k.color,lineHeight:1}}>{k.docs}</div>
                </div>
              ))}
            </div>
          </div>

          {/* BLOCO 2: % Justificativas GLOBAL — destaque visual */}
          <div style={{borderRadius:14,padding:'20px',background:`linear-gradient(135deg,${globalJustColor}12,${globalJustColor}05)`,border:`1px solid ${globalJustColor}30`,transition:'all .4s',position:'relative',overflow:'hidden'}}>
            <div style={{position:'absolute',top:-30,left:-30,width:150,height:150,borderRadius:'50%',background:`radial-gradient(circle,${globalJustColor}08,transparent 70%)`,pointerEvents:'none'}}/>
            <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14}}>
              <div>
                <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.12em',marginBottom:4}}>Justificativas · Meta 80%</div>
                <div style={{fontSize:9,color:C.muted,marginTop:2}}>
                  {periodo==='HOJE'?'Hoje':periodo==='ONTEM'?'Ontem':periodo==='SEMANA'?'Semana':periodo==='MES'?'Mês':periodo==='ANO'?'Ano':'Período'} · todas as unidades
                </div>
                {unidFilt&&<div style={{fontSize:9,color:C.amber,marginTop:4,fontStyle:'italic'}}>↳ justificando: {unidFilt}</div>}
              </div>
              <div style={{textAlign:'right'}}>
                <div style={{fontSize:52,fontWeight:900,color:globalJustColor,lineHeight:1,letterSpacing:'-2px'}}>{globalJustStats.pct}<span style={{fontSize:26}}>%</span></div>
                <div style={{fontSize:9,color:C.muted,marginTop:2}}>{globalJustStats.justified}/{globalJustStats.total} méd. just.</div>
              </div>
            </div>
            <div style={{position:'relative',height:12,background:'rgba(255,255,255,0.06)',borderRadius:6,overflow:'hidden',marginBottom:6}}>
              <div style={{position:'absolute',top:0,left:0,height:'100%',width:`${globalJustStats.pct}%`,background:`linear-gradient(90deg,${globalJustColor},${globalJustColor}bb)`,borderRadius:6,transition:'width .7s ease',boxShadow:`0 0 12px ${globalJustColor}40`}}/>
              <div style={{position:'absolute',top:0,left:'80%',width:2.5,height:'100%',background:'rgba(255,255,255,0.5)'}}/>
            </div>
            <div style={{display:'flex',justifyContent:'space-between',fontSize:9,color:C.muted}}>
              <span>{globalJustStats.justified} justificado{globalJustStats.justified!==1?'s':''} de {globalJustStats.total} total de ocorrências</span>
              <span style={{fontWeight:700,color:globalJustColor}}>← meta 80%</span>
            </div>
          </div>

          {/* BLOCO 3: Lista médicos + justificativas */}
          {unidFilt&&allDocsList.length>0&&(
            <div style={{...card,flex:1}}>
              <div style={{fontSize:9,fontWeight:700,color:C.amber,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:10}}>
                ✎ Médicos · {horasFilt.length>0?horasFilt.map(h=>String(h).padStart(2,'0')+'h').join(', '):'todas as horas'}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:4,maxHeight:240,overflowY:'auto'}}>
                {allDocsList.map((doc,i)=>{
                  const hora=docHoraMap[doc.nm]
                  const key=doc.nm+'::'+String(hora??'x')
                  const isJust=docsJustSet.has(key)
                  const savedJust=justAgFiltered.find(j=>j.nm_medico===doc.nm&&(hora==null||parseInt(j.hora)===hora||j.hora===-1))
                  const draftText=agJustDraft[key]!==undefined?agJustDraft[key]:(savedJust?.texto||'')
                  const isExpanded=justExpanded[key]||false
                  const cfg=getStatusCfg(doc.status)
                  return(
                    <div key={i} style={{borderRadius:8,background:isJust?`${cfg.color}08`:'rgba(255,255,255,0.02)',border:`0.5px solid ${isJust?cfg.color+'25':'rgba(255,255,255,0.06)'}`,overflow:'hidden'}}>
                      <div style={{display:'flex',alignItems:'center',gap:7,padding:'7px 10px'}}>
                        {hora!=null&&<span style={{fontFamily:'monospace',fontSize:9,color:C.muted,flexShrink:0,minWidth:26}}>{String(hora).padStart(2,'0')}h</span>}
                        <div style={{width:5,height:5,borderRadius:'50%',background:cfg.color,flexShrink:0}}/>
                        <span style={{fontSize:10,fontWeight:600,color:C.text,flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{doc.nm}</span>
                        <span style={{fontSize:8,fontWeight:700,padding:'1px 6px',borderRadius:10,background:`${cfg.color}15`,color:cfg.color,border:`0.5px solid ${cfg.color}28`,flexShrink:0,whiteSpace:'nowrap'}}>{cfg.label}</span>
                        {isJust&&<span style={{fontSize:9,color:C.emerald}}>✓</span>}
                        <button onClick={()=>setJustExpanded(prev=>({...prev,[key]:!isExpanded}))}
                          style={{background:isExpanded?'rgba(245,158,11,0.15)':'rgba(255,255,255,0.05)',border:`0.5px solid ${isExpanded?'rgba(245,158,11,0.4)':'rgba(255,255,255,0.1)'}`,borderRadius:5,color:isExpanded?C.amber:C.muted,fontSize:11,cursor:'pointer',flexShrink:0,padding:'2px 6px',transition:'all .15s'}}>
                          {isExpanded?'▲':'✎'}
                        </button>
                      </div>
                      {!isExpanded&&isJust&&savedJust?.texto&&(
                        <div style={{padding:'0 10px 6px',fontSize:9,color:C.muted,fontStyle:'italic',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',borderTop:`0.5px solid ${cfg.color}12`}}>↳ {savedJust.texto}</div>
                      )}
                      {isExpanded&&(
                        <div style={{padding:'8px 10px 10px',borderTop:`0.5px solid ${cfg.color}18`,background:`${cfg.color}04`}}>
                          <textarea value={draftText} onChange={e=>setAgJustDraft(prev=>({...prev,[key]:e.target.value}))}
                            placeholder={`Justificativa para ${doc.nm}${hora!=null?' às '+String(hora).padStart(2,'0')+'h':''}…`} rows={2}
                            style={{width:'100%',background:'rgba(255,255,255,0.05)',border:`0.5px solid ${cfg.color}35`,borderRadius:7,color:C.text,fontSize:10,padding:'6px 8px',outline:'none',resize:'vertical',fontFamily:"'DM Sans',sans-serif"}}/>
                          <div style={{display:'flex',justifyContent:'flex-end',gap:5,marginTop:5}}>
                            {agJustDraft[key]!==undefined&&<button onClick={()=>setAgJustDraft(prev=>{const n={...prev};delete n[key];return n})} style={{background:'transparent',border:'0.5px solid rgba(255,255,255,0.1)',borderRadius:5,color:C.muted,fontSize:9,padding:'3px 8px',cursor:'pointer'}}>Descartar</button>}
                            <button onClick={async()=>{if(!draftText.trim())return;await saveAgJust(doc.nm,hora,draftText);setJustExpanded(prev=>({...prev,[key]:false}));setAgJustDraft(prev=>{const n={...prev};delete n[key];return n})}}
                              disabled={agJustSaving||!draftText.trim()}
                              style={{background:draftText.trim()?`${cfg.color}20`:'rgba(255,255,255,0.04)',border:`0.5px solid ${draftText.trim()?cfg.color+'45':'rgba(255,255,255,0.1)'}`,borderRadius:5,color:draftText.trim()?cfg.color:C.muted,fontSize:10,fontWeight:700,padding:'3px 12px',cursor:draftText.trim()?'pointer':'default'}}>
                              {agJustSaving?'Salvando…':'Salvar'}
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          )}
          {unidFilt&&allDocsList.length===0&&(
            <div style={{...card,fontSize:11,color:C.muted,textAlign:'center',padding:'24px 20px'}}>Sem ocorrências nesta unidade no período.</div>
          )}
        </div>
      </div>

      {/* ── SITUAÇÃO DE PONTO — Accordion ── */}
      <div style={{...card,marginBottom:14}}>
        {/* Header + resumo */}
        <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:18}}>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:4}}>Situação de Ponto · por Hora de Agenda</div>
            <div style={{fontSize:10,color:C.muted}}>
              Col O (HR_ENTRADA) · Col S (ATRASO) · Col T (TEMPO DE ATRASO) · Col U (STATUS)
              {horasFilt.length>0?` · filtrando ${horasFilt.map(h=>String(h).padStart(2,'0')+'h').join(', ')}`:' · todas as horas'}
            </div>
          </div>
          <div style={{display:'flex',gap:10,flexShrink:0}}>
            {[
              {label:'Visão 1',sub:'Faltou',value:semPontoFaltaTotal,color:'#185FA5',bg:'#E6F1FB',border:'#B5D4F4'},
              {label:'Visão 2',sub:'Sem Ponto · Atraso',value:semPontoAtrasoTotal,color:'#854F0B',bg:'#FAEEDA',border:'#FAC775'},
              {label:'Com Ponto',sub:'Chegou · Atrasado',value:comPontoAtrasoTotal,color:'#0F6E56',bg:'#E1F5EE',border:'#9FE1CB'},
            ].map(k=>(
              <div key={k.label} style={{textAlign:'center',padding:'12px 18px',borderRadius:12,background:k.bg,border:`1px solid ${k.border}`,minWidth:110}}>
                <div style={{fontSize:9,fontWeight:500,color:k.color,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:2}}>{k.label}</div>
                <div style={{fontSize:32,fontWeight:500,color:k.color,lineHeight:1,letterSpacing:'-1px'}}>{k.value}</div>
                <div style={{fontSize:8,color:k.color,opacity:.7,marginTop:4,lineHeight:1.3}}>{k.sub}</div>
              </div>
            ))}
          </div>
        </div>

        {situacaoPontoHoras.length===0?(
          <div style={{textAlign:'center',padding:'32px 0',color:C.muted,fontSize:12}}>
            {horasFilt.length>0
              ?`Nenhuma ocorrência nas ${horasFilt.map(h=>String(h).padStart(2,'0')+'h').join(', ')}.`
              :'Nenhuma ocorrência de ponto no período.'}
          </div>
        ):(
          <div style={{display:'flex',flexDirection:'column',gap:6}}>
            {situacaoPontoHoras.map(({hora,semPontoFalta,semPontoAtraso,comPontoAtraso})=>{
              const isOpen=expandedHoras[hora]||false
              const totalHora=semPontoFalta.length+semPontoAtraso.length+comPontoAtraso.length
              return(
                <div key={hora} style={{borderRadius:10,border:'0.5px solid rgba(255,255,255,0.1)',overflow:'hidden'}}>
                  {/* Hora header — clicável */}
                  <div onClick={()=>setExpandedHoras(prev=>({...prev,[hora]:!isOpen}))}
                    style={{display:'flex',alignItems:'center',gap:14,padding:'11px 18px',background:isOpen?'rgba(255,255,255,0.05)':'rgba(255,255,255,0.02)',cursor:'pointer',userSelect:'none',transition:'background .15s'}}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.05)'}
                    onMouseLeave={e=>e.currentTarget.style.background=isOpen?'rgba(255,255,255,0.05)':'rgba(255,255,255,0.02)'}>
                    <span style={{fontFamily:'monospace',fontSize:16,fontWeight:500,color:C.text,minWidth:52}}>{String(hora).padStart(2,'0')}:00</span>
                    <div style={{display:'flex',gap:6,flexWrap:'wrap',flex:1}}>
                      {semPontoFalta.length>0&&(
                        <span style={{fontSize:11,padding:'2px 10px',borderRadius:20,background:'#E6F1FB',color:'#185FA5',border:'0.5px solid #B5D4F4',fontWeight:500}}>
                          {semPontoFalta.length} faltou
                        </span>
                      )}
                      {semPontoAtraso.length>0&&(
                        <span style={{fontSize:11,padding:'2px 10px',borderRadius:20,background:'#FAEEDA',color:'#854F0B',border:'0.5px solid #FAC775',fontWeight:500}}>
                          {semPontoAtraso.length} sem ponto · atraso
                        </span>
                      )}
                      {comPontoAtraso.length>0&&(
                        <span style={{fontSize:11,padding:'2px 10px',borderRadius:20,background:'#E1F5EE',color:'#0F6E56',border:'0.5px solid #9FE1CB',fontWeight:500}}>
                          {comPontoAtraso.length} com ponto · atrasado
                        </span>
                      )}
                    </div>
                    <span style={{fontSize:12,color:C.muted,flexShrink:0}}>{isOpen?'▾ recolher':'▸ expandir'}</span>
                  </div>

                  {/* Conteúdo expandido */}
                  {isOpen&&(
                    <div style={{padding:'14px 18px',borderTop:'0.5px solid rgba(255,255,255,0.07)',display:'flex',flexDirection:'column',gap:14}}>

                      {/* VISÃO 1: FALTOU */}
                      {semPontoFalta.length>0&&(
                        <div>
                          <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:8}}>
                            <div style={{width:18,height:1.5,background:'#B5D4F4',borderRadius:1}}/>
                            <span style={{fontSize:10,fontWeight:500,color:'#185FA5',textTransform:'uppercase',letterSpacing:'.1em'}}>
                              Visão 1 · Faltou (HR_ENTRADA vazia + ATRASO=FALTA)
                            </span>
                          </div>
                          <div style={{display:'flex',flexDirection:'column',gap:5}}>
                            {semPontoFalta.map((doc,i)=>{
                              const motivoMap={
                                'Falta Médica':{label:'Falta Médica',color:'#185FA5',bg:'#E6F1FB',border:'#B5D4F4'},
                                'Remarcação Adm':{label:'Remarcação Adm',color:'#534AB7',bg:'#EEEDFE',border:'#CECBF6'},
                                'Remarcação Médico':{label:'Remarcação Médico',color:'#185FA5',bg:'#E6F1FB',border:'#B5D4F4'},
                                'Remarcação médico':{label:'Remarcação Médico',color:'#185FA5',bg:'#E6F1FB',border:'#B5D4F4'},
                              }
                              const motivo=motivoMap[doc.status]||{label:doc.status||'Falta',color:'#185FA5',bg:'#E6F1FB',border:'#B5D4F4'}
                              return(
                                <div key={i} style={{display:'flex',alignItems:'center',borderRadius:8,border:`0.5px solid ${motivo.border}`,width:'100%'}}>
                                  <div style={{padding:'8px 14px',background:motivo.bg,display:'flex',alignItems:'center',gap:10,flex:1,minWidth:0}}>
                                    <div style={{width:5,height:5,borderRadius:'50%',background:motivo.color,flexShrink:0}}/>
                                    <span style={{fontSize:11,color:motivo.color,fontWeight:500}}>{doc.nm}</span>
                                    {doc.nm_local&&<span style={{fontSize:10,color:motivo.color,opacity:.65,whiteSpace:'nowrap'}}>· {doc.nm_local}</span>}
                                  </div>
                                  <div style={{padding:'8px 14px',background:`${motivo.bg}`,borderLeft:`0.5px solid ${motivo.border}`,display:'flex',alignItems:'center',gap:12,flexShrink:0}}>
                                    {doc.hr_ini_min!=null&&<span style={{fontSize:10,color:motivo.color,fontFamily:'monospace',whiteSpace:'nowrap'}}>{fmtHHMM(doc.hr_ini_min)} → {fmtHHMM(doc.hr_fim_min)}</span>}
                                    <span style={{fontSize:10,fontWeight:500,color:motivo.color,whiteSpace:'nowrap'}}>{motivo.label}</span>
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      )}

                      {/* VISÃO 2: SEM PONTO EM ATRASO */}
                      {semPontoAtraso.length>0&&(
                        <div>
                          <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:8}}>
                            <div style={{width:18,height:1.5,background:'#FAC775',borderRadius:1}}/>
                            <span style={{fontSize:10,fontWeight:500,color:'#854F0B',textTransform:'uppercase',letterSpacing:'.1em'}}>
                              Visão 2 · Sem Ponto em Atraso (HR_ENTRADA vazia + ATRASO=SIM)
                            </span>
                          </div>
                          <div style={{display:'flex',flexDirection:'column',gap:5}}>
                            {semPontoAtraso.map((doc,i)=>{
                              const tempoStr=fmtTempoMinFn(doc.tempo_min)
                              const cfg=getStatusCfg(doc.status)
                              const atrColors={
                                'ATRASO':{color:'#854F0B',bg:'#FAEEDA',border:'#FAC775',badge:'#FAC775'},
                                'ATRASO GRAVE':{color:'#993C1D',bg:'#FAECE7',border:'#F5C4B3',badge:'#F5C4B3'},
                                'ATRASO CRÍTICO':{color:'#A32D2D',bg:'#FCEBEB',border:'#F7C1C1',badge:'#F7C1C1'},
                              }
                              const atrCfg=atrColors[doc.status]||atrColors['ATRASO']
                              return(
                                <div key={i} style={{display:'flex',alignItems:'center',borderRadius:8,border:`0.5px solid ${atrCfg.border}`,width:'100%'}}>
                                  <div style={{padding:'8px 14px',background:atrCfg.bg,display:'flex',alignItems:'center',gap:10,flex:1,minWidth:0}}>
                                    <div style={{width:5,height:5,borderRadius:'50%',background:atrCfg.color,flexShrink:0}}/>
                                    <span style={{fontSize:11,color:atrCfg.color,fontWeight:500}}>{doc.nm}</span>
                                    {doc.nm_local&&<span style={{fontSize:10,color:atrCfg.color,opacity:.65,whiteSpace:'nowrap'}}>· {doc.nm_local}</span>}
                                  </div>
                                  <div style={{padding:'8px 14px',background:atrCfg.bg,borderLeft:`0.5px solid ${atrCfg.border}`,display:'flex',alignItems:'center',gap:10,flexShrink:0}}>
                                    {doc.hr_ini_min!=null&&<span style={{fontSize:10,color:atrCfg.color,fontFamily:'monospace',whiteSpace:'nowrap'}}>{fmtHHMM(doc.hr_ini_min)}→{fmtHHMM(doc.hr_fim_min)}</span>}
                                    {tempoStr&&<span style={{fontSize:12,fontWeight:500,color:atrCfg.color,whiteSpace:'nowrap'}}>{tempoStr} em atraso</span>}
                                    <span style={{fontSize:10,padding:'2px 8px',borderRadius:8,background:atrCfg.badge,color:atrCfg.color,whiteSpace:'nowrap',fontWeight:500}}>{cfg.label}</span>
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      )}

                      {/* COM PONTO ATRASADO */}
                      {comPontoAtraso.length>0&&(
                        <div>
                          <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:8}}>
                            <div style={{width:18,height:1.5,background:'#9FE1CB',borderRadius:1}}/>
                            <span style={{fontSize:10,fontWeight:500,color:'#0F6E56',textTransform:'uppercase',letterSpacing:'.1em'}}>
                              Com Ponto · Chegou Atrasado (HR_ENTRADA preenchida + ATRASO=SIM)
                            </span>
                          </div>
                          <div style={{display:'flex',flexDirection:'column',gap:5}}>
                            {comPontoAtraso.map((doc,i)=>{
                              const tempoStr=fmtTempoMinFn(doc.tempo_min)
                              const cfg=getStatusCfg(doc.status)
                              return(
                                <div key={i} style={{display:'flex',alignItems:'center',borderRadius:8,border:'0.5px solid #9FE1CB',width:'100%'}}>
                                  <div style={{padding:'8px 14px',background:'#E1F5EE',display:'flex',alignItems:'center',gap:10,flex:1,minWidth:0}}>
                                    <div style={{width:5,height:5,borderRadius:'50%',background:'#0F6E56',flexShrink:0}}/>
                                    <span style={{fontSize:11,color:'#0F6E56',fontWeight:500}}>{doc.nm}</span>
                                    {doc.nm_local&&<span style={{fontSize:10,color:'#0F6E56',opacity:.65,whiteSpace:'nowrap'}}>· {doc.nm_local}</span>}
                                  </div>
                                  <div style={{padding:'8px 14px',background:'#E1F5EE',borderLeft:'0.5px solid #9FE1CB',display:'flex',alignItems:'center',gap:10,flexShrink:0}}>
                                    {doc.hr_ini_min!=null&&<span style={{fontSize:10,color:'#0F6E56',fontFamily:'monospace',whiteSpace:'nowrap'}}>{fmtHHMM(doc.hr_ini_min)}→{fmtHHMM(doc.hr_fim_min)}</span>}
                                    {tempoStr&&<span style={{fontSize:12,fontWeight:500,color:'#0F6E56',whiteSpace:'nowrap'}}>{tempoStr} em atraso</span>}
                                    <span style={{fontSize:10,padding:'2px 8px',borderRadius:8,background:'#9FE1CB',color:'#085041',whiteSpace:'nowrap',fontWeight:500}}>{cfg.label}</span>
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )
            })}
          </div>
        )}
      </div>

            {/* TENDÊNCIA */}
      <div style={{...card,marginBottom:14}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14}}>
          <div>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.12em'}}>{'TENDÊNCIA · '+tRealCnt+' DIA'+(tRealCnt!==1?'S':'')}</div>
            <div style={{fontSize:10,color:C.muted,marginTop:3}}>Agendadas vs Entregues</div>
          </div>
          <div style={{display:'flex',background:'rgba(255,255,255,0.04)',border:'0.5px solid rgba(255,255,255,0.08)',borderRadius:9,padding:3,gap:2}}>
            {[{key:'real',label:'Real'},{key:'proj',label:'Projeção'}].map(v=><button key={v.key} onClick={()=>setTrendView(v.key)} style={{padding:'5px 14px',borderRadius:6,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .15s',background:trendView===v.key?C.amber:'transparent',color:trendView===v.key?'#1a0800':C.muted}}>{v.label}</button>)}
          </div>
        </div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10,marginBottom:14}}>
          {[
            {label:'Agendadas hoje',value:tLastReal?tLastReal.agendas.toLocaleString('pt-BR'):'—',color:C.amber},
            {label:'Entregues hoje',value:tLastReal?tLastReal.delivered.toLocaleString('pt-BR'):'—',color:C.emerald},
            {label:'Gap acumulado',value:tTotalGap.toLocaleString('pt-BR'),color:C.rose},
            {label:'Média diária',value:tAvgAg.toLocaleString('pt-BR'),color:C.teal},
          ].map(k=><div key={k.label} style={{background:k.color+'12',border:'0.5px solid '+k.color+'20',borderRadius:9,padding:'10px 14px'}}><div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.07em',marginBottom:5}}>{k.label}</div><div style={{fontSize:20,fontWeight:800,color:k.color}}>{k.value}</div></div>)}
        </div>
        {aTip&&(
          <div style={{position:'fixed',left:aTip.x+14,top:aTip.y-70,zIndex:999,pointerEvents:'none',background:'#0A0D16',border:'1px solid rgba(245,158,11,0.3)',borderRadius:10,padding:'10px 14px',boxShadow:'0 8px 32px rgba(0,0,0,0.6)',minWidth:150}}>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,marginBottom:6}}>{aTip.date?aTip.date.slice(5).replace('-','/'):'—'}{aTip.isProj?' · Projeção':''}</div>
            <div style={{display:'flex',flexDirection:'column',gap:4}}>
              <div style={{display:'flex',justifyContent:'space-between',gap:14,paddingBottom:4,borderBottom:'0.5px solid rgba(255,255,255,0.07)'}}><span style={{fontSize:11,color:C.muted}}>Agendas</span><span style={{fontSize:13,fontWeight:800,color:C.amber}}>{(aTip.agendas||0).toLocaleString('pt-BR')}</span></div>
              <div style={{display:'flex',justifyContent:'space-between',gap:14}}><span style={{fontSize:11,color:C.rose}}>Afetadas</span><span style={{fontSize:12,fontWeight:800,color:C.rose}}>{(aTip.impactadas||0).toLocaleString('pt-BR')}</span></div>
            </div>
          </div>
        )}
        {trendAllData.length===0?<div style={{textAlign:'center',padding:'32px 0',color:C.muted,fontSize:12}}>Sem dados suficientes.</div>:(
          <div>
            <svg width="100%" viewBox={'0 0 '+AVW+' '+AVH} style={{display:'block',overflow:'visible',cursor:'crosshair'}}
              onMouseMove={e=>{if(!trendAllData.length)return;const rc=e.currentTarget.getBoundingClientRect();const mx=(e.clientX-rc.left)/rc.width*AVW;let near=null,minD=Infinity;trendAllData.forEach((d,i)=>{const xv=aNT<=1?APL+ACW/2:APL+i/(aNT-1)*ACW;const dist=Math.abs(mx-xv);if(dist<minD){minD=dist;near={...d,idx:i}}});if(near&&minD<(ACW/Math.max(aNT-1,1))*.8)setATip({x:e.clientX,y:e.clientY,...near});else setATip(null)}}
              onMouseLeave={()=>setATip(null)}>
              <defs><clipPath id="aClip"><rect x={APL} y={APT-2} width={ACW} height={ACH+4}/></clipPath></defs>
              {[0,Math.round(tMaxAg*.33),Math.round(tMaxAg*.66),tMaxAg].map((v,ti)=>(<g key={ti}><line x1={APL} y1={atyS(v)} x2={AVW-APR} y2={atyS(v)} stroke="rgba(255,255,255,0.05)" strokeWidth=".5"/><text x={APL-7} y={atyS(v)+4} textAnchor="end" fontSize="9" fill="#475569">{v>=1000?(v/1000).toFixed(0)+'k':v}</text></g>))}
              {trendView==='proj'&&tRealCnt>0&&aNT>tRealCnt&&<line x1={atxP(tRealCnt-1)} y1={APT} x2={atxP(tRealCnt-1)} y2={APT+ACH} stroke="rgba(245,158,11,0.22)" strokeWidth="1" strokeDasharray="4,3"/>}
              <g clipPath="url(#aClip)">
                {tRealCnt>=2&&<path d={atSmooth(byDate,'agendas',atyS,0)+' '+[...byDate].reverse().map((d,i)=>'L'+atxP(tRealCnt-1-i).toFixed(1)+','+atyS(d.delivered||0).toFixed(1)).join(' ')+' Z'} fill="rgba(244,63,94,0.12)"/>}
                {tRealCnt>=2&&<path d={atSmooth(byDate,'delivered',atyS,0)} fill="none" stroke="#10B981" strokeWidth="2.5"/>}
                {tRealCnt>=2&&<path d={atSmooth(byDate,'agendas',atyS,0)} fill="none" stroke="#F59E0B" strokeWidth="2.5"/>}
                {trendView==='proj'&&projData.length>=1&&tRealCnt>=1&&<path d={'M'+atxP(tRealCnt-1).toFixed(1)+','+atyS(tLastReal?.agendas||0).toFixed(1)+atSmooth(projData,'agendas',atyS,tRealCnt).slice(1)} fill="none" stroke="#F59E0B" strokeWidth="1.5" strokeOpacity={.45} strokeDasharray="6,4"/>}
              </g>
              {byDate.map((d,i)=>[<circle key={'a'+i} cx={atxP(i)} cy={atyS(d.agendas||0)} r="4" fill="#F59E0B" stroke="#06080F" strokeWidth="1.5"/>,<circle key={'e'+i} cx={atxP(i)} cy={atyS(d.delivered||0)} r="4" fill="#10B981" stroke="#06080F" strokeWidth="1.5"/>])}
              {aTip&&aTip.idx!=null&&<circle cx={aNT<=1?APL+ACW/2:APL+(aTip.idx)/(aNT-1)*ACW} cy={atyS(aTip.agendas||0)} r="7" fill="none" stroke="#F59E0B" strokeWidth="2" strokeOpacity={.8}/>}
              {trendAllData.map((d,i)=>{const show=aNT<=10||i===0||i===aNT-1||i%Math.ceil(aNT/8)===0;return show?<text key={'x'+i} x={atxP(i)} y={AVH-3} textAnchor="middle" fontSize="9" fill={d.isProj?'rgba(245,158,11,0.4)':'#334155'}>{d.date.slice(5).replace('-','/')}</text>:null})}
              <line x1={APL} y1={APT+ACH} x2={AVW-APR} y2={APT+ACH} stroke="rgba(255,255,255,0.07)" strokeWidth=".5"/>
            </svg>
          </div>
        )}
      </div>

      {/* TOP UNIDADES + ESPECIALIDADES */}
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
        <div style={card}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:14}}>🏥 Top Unidades por Agendas</div>
          {feedList.slice(0,8).map((u,i)=>{
            const pct=maxFeed>0?(u.total/maxFeed)*100:0
            return(<div key={u.nm_local} style={{marginBottom:10}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4,gap:8}}>
                <div style={{display:'flex',alignItems:'center',gap:6,minWidth:0}}><span style={{fontSize:9,fontWeight:800,color:i===0?C.amber:C.muted,minWidth:20}}>#{i+1}</span><span style={{fontSize:11,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{u.nm_local}</span></div>
                <span style={{fontSize:11,fontWeight:700,color:i<3?C.amber:C.teal,flexShrink:0}}>{u.total.toLocaleString('pt-BR')}</span>
              </div>
              <div style={{height:5,borderRadius:3,overflow:'hidden',background:'rgba(255,255,255,0.04)'}}><div style={{height:'100%',background:i<3?C.amber:C.teal,width:pct+'%',transition:'width .5s',borderRadius:3}}/></div>
            </div>)
          })}
        </div>
        <div style={card}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:14}}>🩺 Top Especialidades</div>
          {topSpec.map((s,i)=>{
            const pct=topSpec[0]?.total>0?(s.total/topSpec[0].total)*100:0
            return(<div key={s.name} style={{marginBottom:10}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4,gap:8}}>
                <div style={{display:'flex',alignItems:'center',gap:6,minWidth:0}}><span style={{fontSize:9,fontWeight:800,color:i===0?C.amber:C.muted,minWidth:20}}>#{i+1}</span><span style={{fontSize:11,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.name}</span></div>
                <span style={{fontSize:11,fontWeight:700,color:i<3?C.amber:C.cyan,flexShrink:0}}>{s.total.toLocaleString('pt-BR')}</span>
              </div>
              <div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:4,overflow:'hidden'}}><div style={{height:'100%',background:i<3?C.amber:C.cyan,width:pct+'%',transition:'width .6s',borderRadius:3}}/></div>
            </div>)
          })}
        </div>
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════════
// TAB ESPERA — sem alterações
// ══════════════════════════════════════════════════════════════════════════════
function TabEspera({rows}){
  const[periodo,setPeriodo]=useState('MES')
  const[dateFrom,setDateFrom]=useState('')
  const[dateTo,setDateTo]=useState('')
  const[ufFilt,setUfFilt]=useState('TODOS')
  const[search,setSearch]=useState('')
  const[horaFilt,setHoraFilt]=useState('TODAS')
  const[horaFiltFim,setHoraFiltFim]=useState('TODAS')
  const[sevFilt,setSevFilt]=useState(['MOD','GRV','CRIT'])
  const[unidFilt,setUnidFilt]=useState('')
  const[trendView,setTrendView]=useState('real')
  const[tTip,setTTip]=useState(null)

  const allDates=useMemo(()=>[...new Set(rows.map(r=>norm10(r.dt_registro||r.data_agenda)).filter(Boolean))].sort(),[rows])
  const periodoFn=useMemo(()=>buildFilter(allDates,periodo,dateFrom,dateTo),[allDates,periodo,dateFrom,dateTo])
  const ufs=useMemo(()=>[...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(),[rows])
  const horasDisp=useMemo(()=>{
    const set=new Set()
    rows.filter(d=>periodoFn(d.dt_registro||d.data_agenda)&&d.tempo_espera_min>=15).forEach(d=>{const h=d.hr_registro_espera_min;if(h==null)return;const hora=Math.floor(h/60);if(hora>=0&&hora<=23)set.add(hora)})
    return[...set].sort((a,b)=>a-b)
  },[rows,periodoFn])
  const filtered=useMemo(()=>{
    let r=rows.filter(d=>periodoFn(d.dt_registro||d.data_agenda))
    if(ufFilt!=='TODOS')r=r.filter(d=>d.uf===ufFilt)
    if(search){const q=search.toLowerCase();r=r.filter(d=>[d.nm_local,d.nm_medico,d.cidade].some(v=>String(v||'').toLowerCase().includes(q)))}
    if(horaFilt!=='TODAS'){const ini=parseInt(horaFilt,10),fim=horaFiltFim==='TODAS'?ini:parseInt(horaFiltFim,10);r=r.filter(d=>{const h=d.hr_registro_espera_min;if(h==null)return false;const hora=Math.floor(h/60);return hora>=ini&&hora<=fim})}
    if(unidFilt)r=r.filter(d=>d.nm_local===unidFilt)
    return r
  },[rows,periodoFn,ufFilt,search,horaFilt,horaFiltFim,unidFilt])
  const espStats=useMemo(()=>{
    const comEsp=filtered.filter(d=>d.tempo_espera_min!=null&&d.tempo_espera_min>=15)
    const totalReg=filtered.length
    const grupoMap={}
    comEsp.forEach(d=>{
      const h=d.hr_registro_espera_min;if(h==null)return
      const hora=Math.floor(h/60),mins=Math.floor(h%60);if(hora<0||hora>23)return
      const horaStr=String(hora).padStart(2,'0')+':'+String(mins).padStart(2,'0')
      const unidade=d.nm_local||'Sem Unidade',key=horaStr+'||'+unidade
      if(!grupoMap[key])grupoMap[key]={horaStr,hora,nm_local:unidade,uf:d.uf||'',cidade:d.cidade||'',maxTempo:0,pac:0}
      if(d.tempo_espera_min>grupoMap[key].maxTempo){grupoMap[key].maxTempo=d.tempo_espera_min;grupoMap[key].pac=d.qt_pacientes_aguardando||0}
    })
    const incidentes=Object.values(grupoMap)
    const feedList=incidentes.slice().sort((a,b)=>b.maxTempo-a.maxTempo)
    const modCnt=incidentes.filter(g=>g.maxTempo>=15&&g.maxTempo<=30).length
    const grvCnt=incidentes.filter(g=>g.maxTempo>30&&g.maxTempo<=89).length
    const critCnt=incidentes.filter(g=>g.maxTempo>=90).length
    const faltasList=filtered.filter(d=>String(d.atraso||'').toUpperCase()==='FALTA')
    const atrasosList=filtered.filter(d=>{if(String(d.atraso||'').toUpperCase()!=='SIM')return false;const t=d.tempo_atraso_min;return t!=null&&Math.abs(t)>31})
    const sMap={};atrasosList.forEach(d=>{const s=d.status||'Sem Status';sMap[s]=(sMap[s]||0)+1})
    const statusAt=Object.entries(sMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)
    const gMap={}
    comEsp.forEach(d=>{
      const dt=d.data_agenda;if(!dt)return
      const h=d.hr_registro_espera_min;if(h==null)return
      const hora=Math.floor(h/60),unidade=d.nm_local||'?'
      const key=dt+'||'+String(hora).padStart(2,'0')+'||'+unidade
      if(!gMap[key])gMap[key]={date:dt,maxTempo:0,pac:0}
      if(d.tempo_espera_min>gMap[key].maxTempo){gMap[key].maxTempo=d.tempo_espera_min;gMap[key].pac=d.qt_pacientes_aguardando||0}
    })
    const dMap={}
    Object.values(gMap).forEach(g=>{
      const dt=g.date;if(!dMap[dt])dMap[dt]={date:dt,mod:0,grv:0,crit:0,pac:0}
      const m=g.maxTempo
      if(m>=90)dMap[dt].crit++;else if(m>=31)dMap[dt].grv++;else if(m>=15)dMap[dt].mod++
      dMap[dt].pac+=g.pac
    })
    const byDate=Object.values(dMap).map(d=>({...d,total:d.mod+d.grv+d.crit})).sort((a,b)=>a.date.localeCompare(b.date))
    const addDay=(ds,n)=>{const d=new Date(ds+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)}
    const pts=byDate.slice(-3)
    const slopeMod=pts.length>=2?(pts[pts.length-1].mod-pts[0].mod)/(pts.length-1):0
    const slopeGrv=pts.length>=2?(pts[pts.length-1].grv-pts[0].grv)/(pts.length-1):0
    const slopeCrt=pts.length>=2?(pts[pts.length-1].crit-pts[0].crit)/(pts.length-1):0
    const lastDate=byDate[byDate.length-1]?.date||''
    const lM=byDate[byDate.length-1]?.mod||0,lG=byDate[byDate.length-1]?.grv||0,lC=byDate[byDate.length-1]?.crit||0
    const projData=lastDate?Array.from({length:5},(_,i)=>({date:addDay(lastDate,i+1),mod:Math.max(0,Math.round(lM+slopeMod*(i+1))),grv:Math.max(0,Math.round(lG+slopeGrv*(i+1))),crit:Math.max(0,Math.round(lC+slopeCrt*(i+1))),total:0,pac:0,isProj:true})):[]
    const docsFalta=unidFilt?Object.entries(faltasList.filter(d=>d.nm_local===unidFilt).reduce((a,d)=>{const nm=d.nm_medico||'—';a[nm]=(a[nm]||0)+1;return a},{})).map(([nm,cnt])=>({nm,cnt})).sort((a,b)=>b.cnt-a.cnt):[]
    const docsAtraso=unidFilt?Object.entries(atrasosList.filter(d=>d.nm_local===unidFilt).reduce((a,d)=>{const nm=d.nm_medico||'—';if(!a[nm])a[nm]={cnt:0,status:d.status};a[nm].cnt++;return a},{})).map(([nm,v])=>({nm,...v})).sort((a,b)=>b.cnt-a.cnt):[]
    return{totalReg,modCnt,grvCnt,critCnt,feedList,faltasList,atrasosList,statusAt,byDate,projData,docsFalta,docsAtraso}
  },[filtered,unidFilt])
  const feedListFiltered=useMemo(()=>espStats.feedList.filter(item=>{
    if(sevFilt.includes('CRIT')&&item.maxTempo>=90)return true
    if(sevFilt.includes('GRV')&&item.maxTempo>=31&&item.maxTempo<90)return true
    if(sevFilt.includes('MOD')&&item.maxTempo>=15&&item.maxTempo<=30)return true
    return false
  }),[espStats.feedList,sevFilt])
  const{totalReg,modCnt,grvCnt,critCnt,feedList,faltasList,atrasosList,statusAt,byDate,projData,docsFalta,docsAtraso}=espStats
  const medTotalProb=faltasList.length+atrasosList.length
  const medFPct=medTotalProb>0?Math.round(faltasList.length/medTotalProb*100):0
  const medAPct=medTotalProb>0?Math.round(atrasosList.length/medTotalProb*100):0
  const trendAllData=trendView==='real'?byDate:[...byDate,...projData]
  const trendMaxAll=Math.max(...trendAllData.map(d=>Math.max(d.mod||0,d.grv||0,d.crit||0)),1)
  const trendRealCnt=byDate.length
  const trendLastReal=byDate[byDate.length-1]
  const trendMaxDay=byDate.length>0?byDate.reduce((a,d)=>(d.crit||0)>(a.crit||0)?d:a,byDate[0]):null
  const trendSlope=projData.length>0?((projData[0]?.crit||0)-(trendLastReal?.crit||0)):0
  const VW=800,VH=160,PL=50,PR=20,PT=10,PB=34
  const CW=VW-PL-PR,CH=VH-PT-PB
  const nTotal=trendAllData.length
  const txP=i=>nTotal<=1?PL+CW/2:PL+i/(nTotal-1)*CW
  const tyA=v=>PT+CH-(trendMaxAll>0?(v||0)/trendMaxAll*CH:0)
  const tSmooth=(data,key,scFn,off)=>{
    if(!data.length)return''
    const pts=data.map((d,i)=>({x:parseFloat(txP(i+(off||0)).toFixed(2)),y:parseFloat(scFn(d[key]||0).toFixed(2))}))
    if(pts.length===1)return'M'+pts[0].x+','+pts[0].y
    let p='M'+pts[0].x+','+pts[0].y
    for(let i=1;i<pts.length;i++){
      const p0=pts[i-2]||pts[i-1],p1=pts[i-1],p2=pts[i],p3=pts[i+1]||pts[i]
      const c1x=(p1.x+(p2.x-p0.x)/4).toFixed(2),c1y=(p1.y+(p2.y-p0.y)/4).toFixed(2)
      const c2x=(p2.x-(p3.x-p1.x)/4).toFixed(2),c2y=(p2.y-(p3.y-p1.y)/4).toFixed(2)
      p+=' C'+c1x+','+c1y+' '+c2x+','+c2y+' '+p2.x+','+p2.y
    }
    return p
  }
  if(!rows.length)return(<div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'60vh',gap:14}}><div style={{fontSize:44}}>⏱️</div><div style={{fontSize:18,fontWeight:700,color:'#F1F5F9'}}>Nenhum dado de espera</div></div>)
  const horasDispFim=horaFilt==='TODAS'?[]:horasDisp.filter(h=>h>parseInt(horaFilt))
  const cardE={background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}
  return(
    <div>
      <div style={{marginBottom:16}}><PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('');setUnidFilt('');setHoraFilt('TODAS');setHoraFiltFim('TODAS')}} allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo} label={`${totalReg.toLocaleString('pt-BR')} registros`}/></div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs} showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}}/>
      <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:12,padding:'12px 18px',marginBottom:14}}>
        <div style={{display:'flex',alignItems:'center',gap:16}}>
          {[{label:'Moderada',value:modCnt,sub:'15–30min',color:'#F59E0B'},{label:'Grave',value:grvCnt,sub:'31–1h29',color:'#F97316'},{label:'Crítica',value:critCnt,sub:'+1h30',color:'#F43F5E'}].map((k,i)=>(
            <div key={k.label} style={{display:'flex',flexDirection:'column',alignItems:'center',gap:2,minWidth:58,paddingLeft:i>0?16:0,borderLeft:i>0?'0.5px solid rgba(255,255,255,0.07)':'none'}}>
              <span style={{fontSize:9,fontWeight:700,color:k.color,textTransform:'uppercase',letterSpacing:'.07em'}}>{k.label}</span>
              <span style={{fontSize:22,fontWeight:800,color:k.color,lineHeight:1}}>{k.value.toLocaleString('pt-BR')}</span>
              <span style={{fontSize:8,color:'#475569'}}>{k.sub}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 320px',gap:14,marginBottom:14}}>
        <div style={cardE}>
          {unidFilt&&<div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12,padding:'8px 12px',background:'rgba(245,158,11,0.08)',border:'0.5px solid rgba(245,158,11,0.25)',borderRadius:9}}><div style={{width:6,height:6,borderRadius:'50%',background:'#F59E0B',flexShrink:0}}/><span style={{fontSize:11,color:'#F59E0B',flex:1,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Filtrando: {unidFilt}</span><button onClick={()=>setUnidFilt('')} style={{background:'transparent',border:'none',color:'#475569',cursor:'pointer',fontSize:12}}>✕</button></div>}
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
            <div>
              <div style={{display:'flex',alignItems:'center',gap:8}}><div style={{fontSize:13,fontWeight:700,color:'#F1F5F9'}}>Feed de Esperas por Hora</div><span style={{fontSize:11,fontWeight:700,padding:'2px 8px',borderRadius:10,background:'rgba(245,158,11,0.15)',color:'#F59E0B',border:'0.5px solid rgba(245,158,11,0.3)'}}>{feedListFiltered.length} espera{feedListFiltered.length!==1?'s':''}</span></div>
              <div style={{fontSize:10.5,color:'#475569',marginTop:3}}>TEMPO_DE_ESPERA · hora via HR_REGISTRO_ESPERA</div>
              <div style={{display:'flex',gap:12,marginTop:8}}>{[{label:'Moderada 15–30min',color:'#F59E0B'},{label:'Grave 31–89min',color:'#F97316'},{label:'Crítica ≥90min',color:'#F43F5E'}].map(k=><div key={k.label} style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:8,height:8,borderRadius:'50%',background:k.color}}/><span style={{fontSize:9,color:'#475569'}}>{k.label}</span></div>)}</div>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:8,flexShrink:0,flexWrap:'wrap'}}>
              <div style={{display:'flex',alignItems:'center',gap:6,background:'rgba(255,255,255,0.03)',border:'0.5px solid rgba(255,255,255,0.08)',borderRadius:8,padding:'4px 10px'}}>
                <span style={{fontSize:10,color:'#475569',fontWeight:600}}>Exibir:</span>
                {[{key:'MOD',label:'Mod',color:'#F59E0B'},{key:'GRV',label:'Grv',color:'#F97316'},{key:'CRIT',label:'Crt',color:'#F43F5E'}].map(s=>{const sel=sevFilt.includes(s.key);return<button key={s.key} onClick={()=>setSevFilt(prev=>prev.includes(s.key)?prev.filter(x=>x!==s.key):[...prev,s.key])} style={{display:'flex',alignItems:'center',gap:4,padding:'3px 8px',borderRadius:6,border:'none',background:sel?s.color+'20':'transparent',cursor:'pointer'}}><div style={{width:12,height:12,borderRadius:3,border:`1.5px solid ${s.color}`,background:sel?s.color:'transparent',display:'flex',alignItems:'center',justifyContent:'center'}}>{sel&&<span style={{color:'#0A0D16',fontSize:9,fontWeight:900}}>✓</span>}</div><span style={{fontSize:10,fontWeight:600,color:sel?s.color:'#475569'}}>{s.label}</span></button>})}
              </div>
              <select value={horaFilt} onChange={e=>{setHoraFilt(e.target.value);setHoraFiltFim('TODAS');setUnidFilt('')}} style={{background:'rgba(255,255,255,0.05)',border:'0.5px solid rgba(245,158,11,0.2)',borderRadius:8,color:'#F1F5F9',fontSize:11,padding:'5px 8px',outline:'none',cursor:'pointer'}}><option value="TODAS">Todas</option>{horasDisp.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}</select>
              {horaFilt!=='TODAS'&&<><span style={{fontSize:11,color:'#475569'}}>→</span><select value={horaFiltFim} onChange={e=>setHoraFiltFim(e.target.value)} style={{background:'rgba(255,255,255,0.05)',border:'0.5px solid rgba(245,158,11,0.2)',borderRadius:8,color:'#F1F5F9',fontSize:11,padding:'5px 8px',outline:'none',cursor:'pointer'}}><option value="TODAS">{String(horaFilt).padStart(2,'0')}:00 só</option>{horasDispFim.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}</select></>}
            </div>
          </div>
          {feedListFiltered.length===0?<div style={{textAlign:'center',padding:'32px 0',color:'#475569',fontSize:12}}>Sem esperas ≥ 15min no período/filtro.</div>:(
            <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:600,overflowY:'auto'}}>
              {feedListFiltered.map((item,i)=>{
                const cls=clsEspera(item.maxTempo),isCrit=item.maxTempo>=90,isGrv=item.maxTempo>=31&&item.maxTempo<90,isMod=item.maxTempo>=15&&item.maxTempo<=30,isSel=unidFilt===item.nm_local
                return<div key={i} onClick={()=>setUnidFilt(isSel?'':item.nm_local)} style={{display:'flex',alignItems:'center',gap:12,padding:'10px 14px',borderRadius:10,cursor:'pointer',background:isSel?`${cls.color}18`:isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':isMod?'rgba(245,158,11,0.03)':'rgba(255,255,255,0.02)',border:isSel?`1px solid ${cls.color}55`:`0.5px solid ${i<3?cls.border:'rgba(255,255,255,0.05)'}`,transition:'all .15s'}} onMouseEnter={e=>{if(!isSel)e.currentTarget.style.background=cls.bg}} onMouseLeave={e=>{if(!isSel)e.currentTarget.style.background=isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':isMod?'rgba(245,158,11,0.03)':'rgba(255,255,255,0.02)'}}>
                  <div style={{width:8,height:8,borderRadius:'50%',background:cls.color,flexShrink:0,boxShadow:isCrit?`0 0 8px ${cls.color}`:'none'}}/>
                  <div style={{fontFamily:'monospace',fontSize:13,fontWeight:700,color:'#94A3B8',flexShrink:0,minWidth:44}}>{item.horaStr}</div>
                  <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,color:'#F1F5F9',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{item.nm_local}</div><div style={{fontSize:10,color:'#475569',marginTop:2}}>{[item.cidade,item.uf].filter(Boolean).join(' · ')}</div></div>
                  <div style={{textAlign:'center',flexShrink:0,minWidth:42}}><div style={{fontSize:12,fontWeight:700,color:'#0EA5E9'}}>{item.pac>0?item.pac:'—'}</div><div style={{fontSize:9,color:'#475569'}}>pac.</div></div>
                  <div style={{fontSize:15,fontWeight:900,color:cls.color,flexShrink:0,minWidth:52,textAlign:'right'}}>{fmtMin(item.maxTempo)}</div>
                  <span style={{fontSize:9.5,fontWeight:700,padding:'3px 9px',borderRadius:20,background:cls.bg,color:cls.color,border:`0.5px solid ${cls.border}`,whiteSpace:'nowrap',flexShrink:0}}>{cls.label}</span>
                </div>
              })}
            </div>
          )}
        </div>
        <div style={cardE}>
          <div style={{fontSize:13,fontWeight:700,color:'#F1F5F9',marginBottom:14}}>Médicos — Falta e Atraso</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1px 1fr',gap:0,alignItems:'stretch'}}>
            <div style={{paddingRight:16,display:'flex',flexDirection:'column',gap:10}}>
              {[{label:'Faltas',value:faltasList.length,color:'#F43F5E',pct:medFPct},{label:'Atrasos >31min',value:atrasosList.length,color:'#F59E0B',pct:medAPct}].map(k=>(
                <div key={k.label}><div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',marginBottom:5}}><div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:6,height:6,borderRadius:'50%',background:k.color}}/><span style={{fontSize:11,color:'#94A3B8'}}>{k.label}</span></div><span style={{fontSize:13,fontWeight:700,color:k.color}}>{k.value} <span style={{fontSize:9,color:'#475569',fontWeight:400}}>{k.pct}%</span></span></div><div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:4,overflow:'hidden'}}><div style={{height:'100%',background:k.color,width:`${k.pct}%`,borderRadius:3,transition:'width .6s'}}/></div></div>
              ))}
              <div style={{display:'flex',alignItems:'center',gap:6,paddingTop:8,borderTop:'0.5px solid rgba(255,255,255,0.06)',marginTop:2}}><span style={{fontSize:22,fontWeight:800,color:'#F97316',lineHeight:1}}>{medTotalProb}</span><span style={{fontSize:10,color:'#475569'}}>total</span></div>
            </div>
            <div style={{background:'rgba(255,255,255,0.07)',margin:'0 16px'}}/>
            <div style={{paddingLeft:16,display:'flex',flexDirection:'column',gap:8}}>
              <span style={{fontSize:9,fontWeight:700,color:'#475569',textTransform:'uppercase',letterSpacing:'.09em'}}>Classificação</span>
              {statusAt.length===0&&<div style={{color:'#475569',fontSize:11}}>Nenhuma ocorrência.</div>}
              {statusAt.map(({k,v})=>{const cfg=getStatusCfg(k),pctAt=statusAt[0]?.v>0?Math.round(v/statusAt[0].v*100):0;return<div key={k}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:3}}><span style={{fontSize:10,color:'#94A3B8'}}>{cfg.label}</span><div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontSize:12,fontWeight:700,color:cfg.color}}>{v}</span><span style={{fontSize:9,color:'#475569'}}>{pctAt}%</span></div></div><div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:5,overflow:'hidden'}}><div style={{height:'100%',background:`linear-gradient(90deg,${cfg.color},${cfg.color}88)`,width:`${pctAt}%`,borderRadius:3,transition:'width .6s'}}/></div></div>})}
            </div>
          </div>
          {unidFilt&&(docsFalta.length>0||docsAtraso.length>0)&&(
            <div style={{marginTop:12,paddingTop:12,borderTop:'0.5px solid rgba(255,255,255,0.06)'}}>
              <div style={{fontSize:9,fontWeight:700,color:'#475569',textTransform:'uppercase',letterSpacing:'.09em',marginBottom:8}}>Médicos · {unidFilt}</div>
              <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:140,overflowY:'auto'}}>
                {docsFalta.map(d=><div key={'f'+d.nm} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0'}}><div style={{width:6,height:6,borderRadius:'50%',background:'#F43F5E',flexShrink:0}}/><span style={{fontSize:10.5,color:'#F1F5F9',flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.nm}</span><span style={{fontSize:9,fontWeight:700,padding:'1px 7px',borderRadius:20,background:'rgba(244,63,94,0.12)',color:'#F43F5E',border:'0.5px solid rgba(244,63,94,0.3)',flexShrink:0}}>Falta</span></div>)}
                {docsAtraso.map(d=>{const cfg=getStatusCfg(d.status);return<div key={'a'+d.nm} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0'}}><div style={{width:6,height:6,borderRadius:'50%',background:'#F59E0B',flexShrink:0}}/><span style={{fontSize:10.5,color:'#F1F5F9',flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.nm}</span><span style={{fontSize:9,fontWeight:700,padding:'1px 7px',borderRadius:20,background:`${cfg.color}18`,color:cfg.color,border:`0.5px solid ${cfg.color}30`,flexShrink:0,whiteSpace:'nowrap'}}>{cfg.label}</span></div>})}
              </div>
            </div>
          )}
          {unidFilt&&docsFalta.length===0&&docsAtraso.length===0&&<div style={{marginTop:10,fontSize:11,color:'#475569',textAlign:'center'}}>Sem ocorrências nesta unidade.</div>}
        </div>
      </div>
      <div style={{...cardE,marginBottom:14}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
          <div><div style={{fontSize:10,fontWeight:700,color:'#475569',textTransform:'uppercase',letterSpacing:'.13em'}}>{'TENDÊNCIA · '+trendRealCnt+' DIA'+(trendRealCnt!==1?'S':'')}</div><div style={{fontSize:10.5,color:'#475569',marginTop:4}}>Esperas críticas ao longo do tempo</div></div>
          <div style={{display:'flex',background:'rgba(255,255,255,0.04)',border:'0.5px solid rgba(255,255,255,0.08)',borderRadius:9,padding:3,gap:2}}>{[{key:'real',label:'Real'},{key:'proj',label:'Projeção'}].map(v=><button key={v.key} onClick={()=>setTrendView(v.key)} style={{padding:'5px 16px',borderRadius:6,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .15s',background:trendView===v.key?'#F59E0B':'transparent',color:trendView===v.key?'#1a0800':'#475569'}}>{v.label}</button>)}</div>
        </div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10,marginBottom:20}}>
          {(trendView==='real'?[{label:'Crítica hoje',value:String(trendLastReal?.crit||'—'),color:'#F43F5E'},{label:'Grave hoje',value:String(trendLastReal?.grv||'—'),color:'#F97316'},{label:'Moderada hoje',value:String(trendLastReal?.mod||'—'),color:'#F59E0B'},{label:'Pior dia',value:trendMaxDay?(trendMaxDay.crit+' em '+trendMaxDay.date.slice(5).replace('-','/')):'—',color:'#F43F5E'}]:[{label:'Proj. crítica',value:String(projData[0]?.crit||'—'),color:'#F43F5E'},{label:'Proj. grave',value:String(projData[0]?.grv||'—'),color:'#F97316'},{label:'Proj. moderada',value:String(projData[0]?.mod||'—'),color:'#F59E0B'},{label:'Tendência',value:trendSlope>=0?'+'+trendSlope:String(trendSlope),color:trendSlope>0?'#F43F5E':trendSlope<0?'#10B981':'#475569'}]).map(k=><div key={k.label} style={{background:k.color+'14',border:'0.5px solid '+k.color+'22',borderRadius:10,padding:'11px 14px'}}><div style={{fontSize:9,fontWeight:700,color:'#475569',textTransform:'uppercase',letterSpacing:'.08em',marginBottom:7}}>{k.label}</div><div style={{fontSize:22,fontWeight:800,color:k.color}}>{k.value}</div></div>)}
        </div>
        {trendAllData.length===0?<div style={{textAlign:'center',padding:'40px 0',color:'#475569',fontSize:12}}>Sem dados suficientes.</div>:(
          <div>
            {tTip&&<div style={{position:'fixed',left:tTip.x+14,top:tTip.y-60,zIndex:999,pointerEvents:'none',background:'#0A0D16',border:'1px solid rgba(244,63,94,0.35)',borderRadius:10,padding:'10px 14px',boxShadow:'0 8px 32px rgba(0,0,0,0.6)',minWidth:140}}><div style={{fontSize:10,fontWeight:700,color:'#475569',marginBottom:6}}>{tTip.date?tTip.date.slice(5).replace('-','/'):'—'}{tTip.isProj?' · Projeção':''}</div><div style={{display:'flex',flexDirection:'column',gap:5}}><div style={{display:'flex',justifyContent:'space-between',gap:16}}><span style={{fontSize:11,color:'#F43F5E'}}>● Crítica</span><span style={{fontSize:13,fontWeight:800,color:'#F43F5E'}}>{tTip.crit||0}</span></div><div style={{display:'flex',justifyContent:'space-between',gap:16}}><span style={{fontSize:11,color:'#F97316'}}>● Grave</span><span style={{fontSize:12,fontWeight:700,color:'#F97316'}}>{tTip.grv||0}</span></div><div style={{display:'flex',justifyContent:'space-between',gap:16}}><span style={{fontSize:11,color:'#F59E0B'}}>● Moderada</span><span style={{fontSize:12,fontWeight:700,color:'#F59E0B'}}>{tTip.mod||0}</span></div></div></div>}
            <svg width="100%" viewBox={'0 0 '+VW+' '+VH} style={{display:'block',overflow:'visible',cursor:'crosshair'}}
              onMouseMove={e=>{if(!trendAllData.length)return;const rc=e.currentTarget.getBoundingClientRect();const mx=(e.clientX-rc.left)/rc.width*VW;let near=null,minD=Infinity;trendAllData.forEach((d,i)=>{const xv=nTotal<=1?PL+CW/2:PL+i/(nTotal-1)*CW;const dist=Math.abs(mx-xv);if(dist<minD){minD=dist;near={...d,idx:i}}});if(near&&minD<(CW/Math.max(nTotal-1,1))*.8)setTTip({x:e.clientX,y:e.clientY,...near});else setTTip(null)}}
              onMouseLeave={()=>setTTip(null)}>
              <defs><clipPath id="tClip"><rect x={PL} y={PT-2} width={CW} height={CH+4}/></clipPath></defs>
              {[0,Math.round(trendMaxAll*.33),Math.round(trendMaxAll*.66),trendMaxAll].map((v,ti)=><g key={ti}><line x1={PL} y1={tyA(v)} x2={VW-PR} y2={tyA(v)} stroke="rgba(255,255,255,0.05)" strokeWidth=".5"/><text x={PL-7} y={tyA(v)+4} textAnchor="end" fontSize="9" fill="#475569">{v}</text></g>)}
              {trendView==='proj'&&trendRealCnt>0&&nTotal>trendRealCnt&&<line x1={txP(trendRealCnt-1)} y1={PT} x2={txP(trendRealCnt-1)} y2={PT+CH} stroke="rgba(245,158,11,0.25)" strokeWidth="1" strokeDasharray="4,3"/>}
              <g clipPath="url(#tClip)">
                {trendRealCnt>=2&&<path d={tSmooth(byDate,'crit',tyA,0)+' L'+txP(trendRealCnt-1).toFixed(1)+','+(PT+CH).toFixed(1)+' L'+txP(0).toFixed(1)+','+(PT+CH).toFixed(1)+' Z'} fill="rgba(244,63,94,0.10)"/>}
                {trendRealCnt>=2&&<path d={tSmooth(byDate,'mod',tyA,0)} fill="none" stroke="#F59E0B" strokeWidth="2" strokeOpacity={.8}/>}
                {trendRealCnt>=2&&<path d={tSmooth(byDate,'grv',tyA,0)} fill="none" stroke="#F97316" strokeWidth="2.5" strokeOpacity={.9}/>}
                {trendRealCnt>=2&&<path d={tSmooth(byDate,'crit',tyA,0)} fill="none" stroke="#F43F5E" strokeWidth="3"/>}
                {trendView==='proj'&&projData.length>=1&&trendRealCnt>=1&&<path d={'M'+txP(trendRealCnt-1).toFixed(1)+','+tyA(trendLastReal?.crit||0).toFixed(1)+tSmooth(projData,'crit',tyA,trendRealCnt).slice(1)} fill="none" stroke="#F43F5E" strokeWidth="2" strokeOpacity={.45} strokeDasharray="6,4"/>}
              </g>
              {byDate.map((d,i)=>[<circle key={'m'+i} cx={txP(i)} cy={tyA(d.mod||0)} r="3" fill="#F59E0B" stroke="#06080F" strokeWidth="1.5"/>,<circle key={'g'+i} cx={txP(i)} cy={tyA(d.grv||0)} r="3.5" fill="#F97316" stroke="#06080F" strokeWidth="1.5"/>,<circle key={'c'+i} cx={txP(i)} cy={tyA(d.crit||0)} r="4" fill="#F43F5E" stroke="#06080F" strokeWidth="1.5"/>])}
              {tTip&&tTip.idx!=null&&<circle cx={nTotal<=1?PL+CW/2:PL+(tTip.idx)/(nTotal-1)*CW} cy={tyA(tTip.crit||0)} r="7" fill="none" stroke="#F43F5E" strokeWidth="2" strokeOpacity={.8}/>}
              {trendAllData.map((d,i)=>{const show=nTotal<=10||i===0||i===nTotal-1||i%Math.ceil(nTotal/8)===0;return show?<text key={'x'+i} x={txP(i)} y={VH-3} textAnchor="middle" fontSize="9" fill={d.isProj?'rgba(245,158,11,0.4)':'#334155'}>{d.date.slice(5).replace('-','/')}</text>:null})}
              <line x1={PL} y1={PT+CH} x2={VW-PR} y2={PT+CH} stroke="rgba(255,255,255,0.07)" strokeWidth=".5"/>
            </svg>
          </div>
        )}
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════════
// HOME
// ══════════════════════════════════════════════════════════════════════════════
export default function Home(){
  const[tab,setTab]=useState('espera')
  const[agendas,setAgendas]=useState([])
  const[espera,setEspera]=useState([])
  const[loadingDB,setLoadingDB]=useState(true)
  const[storing,setStoring]=useState(false)
  const[storeMsg,setStoreMsg]=useState('')
  const[timestamp,setTimestamp]=useState('')
  const[storageInfo,setStorageInfo]=useState({agendas:0,espera:0})

  const loadTable=useCallback(async table=>{
    const PAGE=1000;let all=[],offset=0
    while(true){
      const res=await fetch(`${SB_URL}/rest/v1/${table}?select=*&order=id.asc&limit=${PAGE}&offset=${offset}`,{headers:SBH})
      if(!res.ok)break
      const batch=await res.json()
      if(!Array.isArray(batch)||!batch.length)break
      all=all.concat(batch);if(batch.length<PAGE)break;offset+=PAGE
    }
    return all
  },[])

  useEffect(()=>{
    const load=async()=>{
      try{
        setStoreMsg('Conectando…')
        const[ag,esp]=await Promise.all([loadTable('agendas'),loadTable('espera')])
        if(ag.length||esp.length){setAgendas(ag);setEspera(esp);setStorageInfo({agendas:ag.length,espera:esp.length});const ts=ag[0]?.verif_ts||esp[0]?.verif_ts||'';if(ts)setTimestamp(ts);setStoreMsg(`☁ ${ag.length.toLocaleString('pt-BR')} agendas · ${esp.length.toLocaleString('pt-BR')} esperas`);setTimeout(()=>setStoreMsg(''),4000)}else setStoreMsg('')
      }catch(e){setStoreMsg(`Erro: ${e.message}`)}
      setLoadingDB(false)
    }
    load()
  },[loadTable])

  const handleUpload=useCallback(async e=>{
    const file=e.target.files[0];if(!file)return;e.target.value=''
    setStoring(true)
    const now=new Date(),pad=n=>String(n).padStart(2,'0')
    const ts=`${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`
    setTimestamp(ts)
    try{
      setStoreMsg('Lendo planilha…')
      const buf=await file.arrayBuffer()
      const wb=XLSX.read(buf,{type:'buffer'})
      const ws=wb.Sheets['PONTOS']||wb.Sheets[wb.SheetNames[0]]
      const json=XLSX.utils.sheet_to_json(ws,{range:3,defval:''})
      setStoreMsg(`${json.length.toLocaleString('pt-BR')} linhas — salvando…`)
      const CHUNK=500
      for(let i=0;i<json.length;i+=CHUNK){setStoreMsg(`Agendas… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`);await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:json.slice(i,i+CHUNK),ts,table:'agendas'})})}
      for(let i=0;i<json.length;i+=CHUNK){setStoreMsg(`Espera… ${Math.min(i+CHUNK,json.length).toLocaleString('pt-BR')}/${json.length.toLocaleString('pt-BR')}`);await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:json.slice(i,i+CHUNK),ts,table:'espera'})})}
      setStoreMsg('Recarregando…')
      const[ag,esp]=await Promise.all([loadTable('agendas'),loadTable('espera')])
      setAgendas(ag);setEspera(esp);setStorageInfo({agendas:ag.length,espera:esp.length})
      setStoreMsg(`✓ ${ag.length.toLocaleString('pt-BR')} agendas · ${esp.length.toLocaleString('pt-BR')} esperas`)
      setTimeout(()=>setStoreMsg(''),4000)
    }catch(e){setStoreMsg(`⚠ Erro: ${e.message}`)}
    setStoring(false)
  },[loadTable])

  const handleClear=async()=>{
    if(!confirm('Apagar TODOS os dados do banco?'))return
    setStoring(true);setStoreMsg('Apagando…')
    try{
      await fetch(`${SB_URL}/rest/v1/agendas?id=gte.0`,{method:'DELETE',headers:SBH})
      await fetch(`${SB_URL}/rest/v1/espera?id=gte.0`,{method:'DELETE',headers:SBH})
    }catch(e){console.error(e)}
    setAgendas([]);setEspera([]);setStorageInfo({agendas:0,espera:0});setTimestamp('');setStoreMsg('');setStoring(false)
  }

  return(
    <div style={{minHeight:'100vh',background:'#06080F',fontFamily:"'DM Sans','Segoe UI',sans-serif",color:'#F1F5F9'}}>
      <style suppressHydrationWarning>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800;900&family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(245,158,11,0.2);border-radius:4px}
        select option{background:#0A0D16;color:#F1F5F9}
        input::placeholder{color:#475569}
        textarea::placeholder{color:#475569}
        select{appearance:auto}
      `}</style>
      <div style={{background:'#070910',borderBottom:'1px solid rgba(245,158,11,0.1)',display:'flex',alignItems:'center',padding:'0 32px',height:52,position:'sticky',top:0,zIndex:100,backdropFilter:'blur(12px)'}}>
        <div style={{position:'absolute',top:-40,left:'50%',transform:'translateX(-50%)',width:500,height:80,background:'radial-gradient(ellipse,rgba(245,158,11,0.07),transparent 70%)',pointerEvents:'none'}}/>
        <div style={{display:'flex',alignItems:'center',gap:10,marginRight:32}}>
          <div style={{width:28,height:28,borderRadius:8,background:'linear-gradient(135deg,#F59E0B,#F97316)',display:'flex',alignItems:'center',justifyContent:'center'}}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
          </div>
          <div><span style={{fontSize:14,fontWeight:900,color:'#F1F5F9',fontFamily:"'Syne',sans-serif",letterSpacing:'-.3px'}}>Monitor </span><span style={{fontSize:14,fontWeight:900,color:'#F59E0B',fontFamily:"'Syne',sans-serif",letterSpacing:'-.3px'}}>Clínicas</span></div>
        </div>
        <div style={{display:'flex',height:'100%',gap:0}}>
          {[{key:'espera',label:'Fila de Espera',count:storageInfo.espera},{key:'agendas',label:'Agendas Médicas',count:storageInfo.agendas}].map(t=>(
            <button key={t.key} onClick={()=>setTab(t.key)} style={{padding:'0 20px',border:'none',background:'transparent',cursor:'pointer',fontSize:12,fontWeight:700,height:'100%',color:tab===t.key?'#F59E0B':'#475569',borderBottom:tab===t.key?'2px solid #F59E0B':'2px solid transparent',transition:'all .2s',display:'flex',alignItems:'center',gap:7}}>
              {t.label}
              {t.count>0&&<span style={{fontSize:9,padding:'1px 6px',borderRadius:10,background:tab===t.key?'rgba(245,158,11,0.2)':'rgba(255,255,255,0.06)',color:tab===t.key?'#F59E0B':'#475569',fontWeight:700}}>{t.count.toLocaleString('pt-BR')}</span>}
            </button>
          ))}
        </div>
        <div style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:14}}>
          <div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:6,height:6,borderRadius:'50%',background:'#10B981',boxShadow:'0 0 6px #10B981'}}/><span style={{fontSize:10,color:'#475569'}}>Banco conectado</span></div>
          {timestamp&&<span style={{fontSize:10,color:'#475569'}}>{timestamp}</span>}
          {storeMsg&&<span style={{fontSize:10,color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?'#10B981':'#F59E0B',fontWeight:600,maxWidth:220,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{storeMsg}</span>}
          {(storageInfo.agendas>0||storageInfo.espera>0)&&!storing&&<button onClick={handleClear} style={{background:'transparent',border:'0.5px solid rgba(244,63,94,0.3)',borderRadius:8,color:'#F43F5E',fontSize:11,padding:'5px 10px',cursor:'pointer'}}>🗑</button>}
          <label style={{background:storing?'rgba(255,255,255,0.06)':'linear-gradient(135deg,#F59E0B,#F97316)',color:storing?'#475569':'#1a0800',fontWeight:800,fontSize:12,padding:'7px 16px',borderRadius:9,cursor:storing?'default':'pointer',transition:'all .2s',whiteSpace:'nowrap',fontFamily:"'Syne',sans-serif"}}>
            {storing?'Salvando…':'+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={storing||loadingDB}/>
          </label>
        </div>
      </div>
      <div style={{padding:'24px 32px'}}>
        {loadingDB?(
          <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'calc(100vh - 120px)',gap:16}}>
            <div style={{fontSize:36,filter:'drop-shadow(0 0 20px #F59E0B88)'}}>⏳</div>
            <div style={{fontSize:16,color:'#94A3B8',fontFamily:"'Syne',sans-serif",fontWeight:700}}>Conectando ao banco…</div>
            {storeMsg&&<div style={{fontSize:12,color:'#F59E0B'}}>{storeMsg}</div>}
          </div>
        ):(<>
          {tab==='agendas'&&<TabAgendas rows={agendas}/>}
          {tab==='espera'&&<TabEspera rows={espera}/>}
        </>)}
      </div>
    </div>
  )
}