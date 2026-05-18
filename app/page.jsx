'use client'

import { useState, useEffect, useMemo, useCallback } from 'react'
import * as XLSX from 'xlsx'

const C = {
  bg:'#06080F', bgCard:'#0A0D16', bgCard2:'#0D1020',
  amber:'#F59E0B', orange:'#F97316', rose:'#F43F5E', teal:'#00C9A7',
  blue:'#3B82F6', violet:'#8B5CF6', emerald:'#10B981', cyan:'#06B6D4',
  text:'#F1F5F9', sub:'#94A3B8', muted:'#475569',
  border:'rgba(245,158,11,0.1)', border2:'rgba(245,158,11,0.22)',
  borderGray:'rgba(255,255,255,0.05)',
}
const SB_URL='https://fwdvzsywudpieqlqnxkp.supabase.co'
const SB_KEY='eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3ZHZ6c3l3dWRwaWVxbHFueGtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODcyNzEsImV4cCI6MjA5NDE2MzI3MX0.SkyfE_HVulz_TyQldI6XpENSJAuu6xDgUEDz4vObKYQ'
const SBH={'apikey':SB_KEY,'Authorization':`Bearer ${SB_KEY}`,'Content-Type':'application/json'}

const fmtMin=m=>{if(m==null)return'—';const abs=Math.round(Math.abs(m)),s=m<0?'-':'';if(abs<60)return`${s}${abs}min`;return`${s}${Math.floor(abs/60)}h${abs%60>0?` ${abs%60}m`:''}`}
const fmtDate=s=>{if(!s)return'—';const[y,mo,d]=String(s).split('-');return`${d}/${mo}/${y}`}
const buildFilter=(allDates,periodo,df,dt)=>{
  if(!allDates.length)return()=>true
  const s=[...allDates].sort(),max=s[s.length-1]
  const sysToday=()=>{const n=new Date();return n.getUTCFullYear()+'-'+String(n.getUTCMonth()+1).padStart(2,'0')+'-'+String(n.getUTCDate()).padStart(2,'0')}
  const sysYest=()=>{const n=new Date();n.setUTCDate(n.getUTCDate()-1);return n.toISOString().slice(0,10)}
 if(periodo==='HOJE')return d => String(d) === String(max)

if(periodo==='ONTEM'){
  const idx=s.indexOf(max)
  const ontem=idx>0?s[idx-1]:max
  return d => String(d) === String(ontem)
}
  
  if(periodo==='SEMANA'){const r=new Date(max+'T00:00:00Z');r.setUTCDate(r.getUTCDate()-6);const c=r.toISOString().slice(0,10);return d=>d>=c&&d<=max}
  if(periodo==='MES')return d=>d.slice(0,7)===max.slice(0,7)
  if(periodo==='ANO')return d=>d.slice(0,4)===max.slice(0,4)
  if(periodo==='PERIODO'){const from=df||s[0],to=dt||max;return d=>d>=from&&d<=to}
  return()=>true
}
const STATUS_CFG={
  'OK':{label:'OK',color:'#10B981',glow:'rgba(16,185,129,0.2)'},
  'ATRASO':{label:'Atraso 31–45min',color:'#F59E0B',glow:'rgba(245,158,11,0.2)'},
  'ATRASO CRÍTICO':{label:'Atraso Crítico',color:'#F97316',glow:'rgba(249,115,22,0.2)'},
  'ATRASO GRAVE':{label:'Atraso Grave',color:'#F43F5E',glow:'rgba(244,63,94,0.2)'},
  'Falta Médica':{label:'Médico Faltou',color:'#3B82F6',glow:'rgba(59,130,246,0.2)'},
  'SEM PONTO':{label:'Sem Ponto',color:'#64748B',glow:'rgba(100,116,139,0.2)'},
}
const getStatusCfg=s=>{
  if(!s)return STATUS_CFG['OK'];if(STATUS_CFG[s])return STATUS_CFG[s]
  const u=s.toUpperCase()
  if(u.includes('CRÍTICO')||u.includes('CRITICO'))return STATUS_CFG['ATRASO CRÍTICO']
  if(u.includes('GRAVE'))return STATUS_CFG['ATRASO GRAVE']
  if(u.includes('ATRASO'))return STATUS_CFG['ATRASO']
  if(u.includes('FALTA'))return STATUS_CFG['Falta Médica']
  return STATUS_CFG['OK']
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
        {PERIODOS.map(p=>(
          <button key={p.key} onClick={()=>onChange(p.key)} style={{padding:'5px 12px',borderRadius:7,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .2s',background:value===p.key?C.amber:'transparent',color:value===p.key?'#1a0800':C.muted}}>{p.label}</button>
        ))}
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

function SearchBar({search,onSearch,uf,onUf,ufs,extra,showClear,onClear}){
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
      {extra}
      {showClear&&<button onClick={onClear} style={{background:'rgba(244,63,94,0.08)',border:`1px solid rgba(244,63,94,0.25)`,borderRadius:10,color:C.rose,fontSize:12,padding:'8px 12px',cursor:'pointer'}}>✕ Limpar</button>}
    </div>
  )
}

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

  const allDates=useMemo(()=>[...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(),[rows])
  const periodoFn=useMemo(()=>buildFilter(allDates,periodo,dateFrom,dateTo),[allDates,periodo,dateFrom,dateTo])
  const ufs=useMemo(()=>[...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(),[rows])

  const horasDisp=useMemo(()=>{
    const set=new Set()
    rows.filter(d=>periodoFn(d.data_agenda)).forEach(d=>{
      const h=d.hr_inicio_min;if(h==null)return
      const hora=Math.floor(h/60);if(hora>=0&&hora<=23)set.add(hora)
    })
    return[...set].sort((a,b)=>a-b)
  },[rows,periodoFn])

  const toggleHora=useCallback(h=>{
    setHorasFilt(prev=>prev.includes(h)?prev.filter(x=>x!==h):[...prev,h])
  },[])

  const filtered=useMemo(()=>{
    let r=rows.filter(d=>periodoFn(d.data_agenda))
    if(ufFilt!=='TODOS')r=r.filter(d=>d.uf===ufFilt)
    if(search){const q=search.toLowerCase();r=r.filter(d=>[d.nm_local,d.nm_medico,d.ds_especialidade,d.cidade].some(v=>String(v||'').toLowerCase().includes(q)))}
    if(horasFilt.length>0)r=r.filter(d=>{const h=d.hr_inicio_min;if(h==null)return false;return horasFilt.includes(Math.floor(h/60))})
    if(unidFilt)r=r.filter(d=>d.nm_local===unidFilt)
    return r
  },[rows,periodoFn,ufFilt,search,horasFilt,unidFilt])

  const agStats=useMemo(()=>{
    const totalRows=filtered.length
    const totalCons=filtered.reduce((a,d)=>a+(d.qt_consulta||0),0)
    const totalEnc =filtered.reduce((a,d)=>a+(d.qt_encaixe ||0),0)
    const totalAg  =totalCons+totalEnc

    // Médicos c/ problemas
    const faltaRows=filtered.filter(d=>String(d.atraso||'').toUpperCase()==='FALTA'||String(d.status||'').toLowerCase().includes('falta'))
    const critRows =filtered.filter(d=>String(d.atraso||'').toUpperCase()==='SIM'&&(String(d.status||'').toUpperCase().includes('CRÍTICO')||String(d.status||'').toUpperCase().includes('CRITICO')))
    const grvRows  =filtered.filter(d=>String(d.atraso||'').toUpperCase()==='SIM'&&String(d.status||'').toUpperCase().includes('GRAVE'))
    const atrRows  =filtered.filter(d=>String(d.atraso||'').toUpperCase()==='SIM'&&!String(d.status||'').toUpperCase().includes('CRÍTICO')&&!String(d.status||'').toUpperCase().includes('CRITICO')&&!String(d.status||'').toUpperCase().includes('GRAVE'))

    const faltaDocs=[...new Set(faltaRows.map(d=>d.nm_medico).filter(Boolean))]
    const critDocs =[...new Set(critRows .map(d=>d.nm_medico).filter(Boolean))]
    const grvDocs  =[...new Set(grvRows  .map(d=>d.nm_medico).filter(Boolean))]
    const atrDocs  =[...new Set(atrRows  .map(d=>d.nm_medico).filter(Boolean))]
    const faltaAg=faltaRows.reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const critAg =critRows .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const grvAg  =grvRows  .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)
    const atrAg  =atrRows  .reduce((a,d)=>a+(d.qt_consulta||0)+(d.qt_encaixe||0),0)

    // Feed de unidades (por QT_CONSULTA + QT_ENCAIXE)
    const uMap={}
    filtered.forEach(d=>{
      const u=d.nm_local||'?'
      if(!uMap[u])uMap[u]={nm_local:u,uf:d.uf||'',cidade:d.cidade||'',consultas:0,encaixe:0,total:0,faltas:0,atrasos:0}
      uMap[u].consultas+=(d.qt_consulta||0)
      uMap[u].encaixe  +=(d.qt_encaixe ||0)
      uMap[u].total    +=(d.qt_consulta||0)+(d.qt_encaixe||0)
      if(String(d.atraso||'').toUpperCase()==='FALTA')uMap[u].faltas++
      if(String(d.atraso||'').toUpperCase()==='SIM')  uMap[u].atrasos++
    })
    const feedList=Object.values(uMap).sort((a,b)=>b.total-a.total).slice(0,20)

    // Top especialidades
    const sMap={}
    filtered.forEach(d=>{
      const s=d.ds_especialidade||'Não informado'
      if(!sMap[s])sMap[s]={name:s,total:0,consultas:0,encaixe:0}
      sMap[s].total    +=(d.qt_consulta||0)+(d.qt_encaixe||0)
      sMap[s].consultas+=(d.qt_consulta||0)
      sMap[s].encaixe  +=(d.qt_encaixe ||0)
    })
    const topSpec=Object.values(sMap).sort((a,b)=>b.total-a.total).slice(0,8)

    // Por data — tendência
    const dMap={}
    filtered.forEach(d=>{
      const dt=d.data_agenda;if(!dt)return
      if(!dMap[dt])dMap[dt]={date:dt,agendas:0,consultas:0,encaixe:0,impactadas:0}
      dMap[dt].agendas   +=(d.qt_consulta||0)+(d.qt_encaixe||0)
      dMap[dt].consultas +=(d.qt_consulta||0)
      dMap[dt].encaixe   +=(d.qt_encaixe ||0)
      const prob=String(d.atraso||'').toUpperCase()==='FALTA'||String(d.atraso||'').toUpperCase()==='SIM'
      if(prob)dMap[dt].impactadas+=(d.qt_consulta||0)+(d.qt_encaixe||0)
    })
    const byDate=Object.values(dMap).map(d=>({...d,delivered:Math.max(0,d.agendas-d.impactadas)})).sort((a,b)=>a.date.localeCompare(b.date))

    // Projeção
    const addDay=(ds,n)=>{const d=new Date(ds+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)}
    const pts=byDate.slice(-3)
    const slopeAg =pts.length>=2?(pts[pts.length-1].agendas   -pts[0].agendas   )/(pts.length-1):0
    const slopeImp=pts.length>=2?(pts[pts.length-1].impactadas-pts[0].impactadas)/(pts.length-1):0
    const slopeDel=pts.length>=2?(pts[pts.length-1].delivered -pts[0].delivered )/(pts.length-1):0
    const lastDate=byDate[byDate.length-1]?.date||''
    const lastAg  =byDate[byDate.length-1]?.agendas   ||0
    const lastImp =byDate[byDate.length-1]?.impactadas||0
    const lastDel =byDate[byDate.length-1]?.delivered ||0
    const projData=lastDate?Array.from({length:5},(_,i)=>({
      date:addDay(lastDate,i+1),
      agendas:   Math.max(0,Math.round(lastAg  +slopeAg *(i+1))),
      impactadas:Math.max(0,Math.round(lastImp +slopeImp*(i+1))),
      delivered: Math.max(0,Math.round(lastDel +slopeDel*(i+1))),
      consultas:0,encaixe:0,isProj:true,
    })):[]

    // Lista de médicos quando unidade filtrada
    const docsFaltaU =unidFilt?faltaDocs.map(nm=>({nm,tipo:'FALTA', status:'Falta Médica'})):[]
    const docsCritU  =unidFilt?critDocs .map(nm=>({nm,tipo:'CRIT',  status:'Atraso Crítico'})):[]
    const docsGrvU   =unidFilt?grvDocs  .map(nm=>({nm,tipo:'GRV',   status:'Atraso Grave'})):[]
    const docsAtrU   =unidFilt?atrDocs  .map(nm=>({nm,tipo:'ATR',   status:'Atraso 31–45min'})):[]

    return{totalRows,totalCons,totalEnc,totalAg,
      faltaDocs,critDocs,grvDocs,atrDocs,faltaAg,critAg,grvAg,atrAg,
      feedList,topSpec,byDate,projData,
      docsFaltaU,docsCritU,docsGrvU,docsAtrU}
  },[filtered,unidFilt])

  const{totalRows,totalCons,totalEnc,totalAg,
    faltaDocs,critDocs,grvDocs,atrDocs,faltaAg,critAg,grvAg,atrAg,
    feedList,topSpec,byDate,projData,
    docsFaltaU,docsCritU,docsGrvU,docsAtrU}=agStats

  // Trend vars
  const trendAllData=trendView==='real'?byDate:[...byDate,...projData]
  const tMaxAg  =Math.max(...trendAllData.map(d=>d.agendas||0),1)
  const tRealCnt=byDate.length
  const tLastReal=byDate[byDate.length-1]
  const tLastDel=tLastReal?.delivered||0
  const tTotalGap=byDate.reduce((a,d)=>a+(d.impactadas||0),0)
  const tVarAg=byDate.length>=2?((tLastReal?.agendas||0)-(byDate[0]?.agendas||0)):0
  const tAvgAg=byDate.length>0?Math.round(byDate.reduce((a,d)=>a+(d.agendas||0),0)/byDate.length):0
  const tSlope=projData.length>0?((projData[0]?.agendas||0)-(tLastReal?.agendas||0)):0

  // SVG chart vars
  const AVW=800,AVH=150,APL=52,APR=20,APT=10,APB=32
  const ACW=AVW-APL-APR,ACH=AVH-APT-APB
  const aNT=trendAllData.length
  const atxP=i=>aNT<=1?APL+ACW/2:APL+i/(aNT-1)*ACW
  const atyS=(v)=>APT+ACH-(tMaxAg>0?(v||0)/tMaxAg*ACH:0)
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

  if(!rows.length)return(
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'60vh',gap:14}}>
      <div style={{fontSize:44}}>📋</div>
      <div style={{fontSize:18,fontWeight:700,color:C.text}}>Nenhuma agenda carregada</div>
      <div style={{fontSize:13,color:C.muted}}>Use o botão para carregar uma planilha</div>
    </div>
  )

  const maxFeed=feedList[0]?.total||1

  return(
    <div>
      {/* PERÍODO + BUSCA */}
      <div style={{marginBottom:14}}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('');setUnidFilt('');setHorasFilt([])}}
          allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo}
          label={totalRows.toLocaleString('pt-BR')+' registros'}/>
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs}
        showClear={ufFilt!=='TODOS'||!!search||horasFilt.length>0||!!unidFilt}
        onClear={()=>{setUfFilt('TODOS');setSearch('');setHorasFilt([]);setUnidFilt('')}}/>

      {/* FILTRO DE HORA — MÚLTIPLA SELEÇÃO */}
      {horasDisp.length>0&&(
        <div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap',marginBottom:14,padding:'10px 14px',background:'rgba(255,255,255,0.02)',border:'0.5px solid rgba(255,255,255,0.06)',borderRadius:10}}>
          <span style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.09em',flexShrink:0}}>Hora início</span>
          {horasDisp.map(h=>{
            const sel=horasFilt.includes(h)
            return(
              <button key={h} onClick={()=>toggleHora(h)} style={{
                padding:'4px 10px',borderRadius:6,border:sel?'none':'0.5px solid rgba(255,255,255,0.08)',
                background:sel?C.amber:'rgba(255,255,255,0.03)',
                color:sel?'#1a0800':C.muted,fontSize:10,fontWeight:sel?700:400,cursor:'pointer',transition:'all .15s',
              }}>{String(h).padStart(2,'0')}h</button>
            )
          })}
          {horasFilt.length>0&&(
            <button onClick={()=>setHorasFilt([])} style={{padding:'4px 10px',borderRadius:6,border:'0.5px solid rgba(244,63,94,0.3)',background:'rgba(244,63,94,0.07)',color:C.rose,fontSize:10,cursor:'pointer'}}>✕ Limpar</button>
          )}
        </div>
      )}

      {/* KPIs — Barra horizontal Opção A */}
      <div style={{display:'flex',background:'rgba(255,255,255,0.025)',border:'0.5px solid rgba(255,255,255,0.07)',borderRadius:12,overflow:'hidden',marginBottom:14}}>

        {/* Bloco 1 — Total + consultas/encaixe */}
        <div style={{flex:'1.4',padding:'16px 20px',borderRight:'0.5px solid rgba(255,255,255,0.07)'}}>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:5}}>Total Agendas</div>
          <div style={{fontSize:30,fontWeight:800,color:C.amber,letterSpacing:'-1px',lineHeight:1}}>{totalAg.toLocaleString('pt-BR')}</div>
          <div style={{display:'flex',gap:14,marginTop:8}}>
            <div style={{display:'flex',alignItems:'center',gap:5}}>
              <div style={{width:18,height:3,background:C.teal,borderRadius:2}}/>
              <span style={{fontSize:9,color:C.muted}}>{totalCons.toLocaleString('pt-BR')} consultas</span>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:5}}>
              <div style={{width:18,height:3,background:C.violet,borderRadius:2}}/>
              <span style={{fontSize:9,color:C.muted}}>{totalEnc.toLocaleString('pt-BR')} encaixe</span>
            </div>
          </div>
          <div style={{marginTop:8,height:4,background:'rgba(255,255,255,0.04)',borderRadius:2,overflow:'hidden',display:'flex',gap:1}}>
            <div style={{flex:totalCons,background:C.teal,opacity:.7}}/>
            <div style={{flex:totalEnc,background:C.violet,opacity:.7}}/>
          </div>
        </div>

        {/* Bloco 2 — Afetadas */}
        <div style={{flex:'1',padding:'16px 20px',borderRight:'0.5px solid rgba(255,255,255,0.07)'}}>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:5}}>Afetadas por problema</div>
          <div style={{fontSize:30,fontWeight:800,color:C.rose,letterSpacing:'-1px',lineHeight:1}}>{(faltaAg+atrAg+critAg+grvAg).toLocaleString('pt-BR')}</div>
          <div style={{fontSize:9,color:C.muted,marginTop:6}}>
            {totalAg>0?((faltaAg+atrAg+critAg+grvAg)/totalAg*100).toFixed(1):'0'}% das agendas totais
          </div>
          <div style={{marginTop:8,height:4,background:'rgba(255,255,255,0.04)',borderRadius:2,overflow:'hidden'}}>
            <div style={{width:(totalAg>0?((faltaAg+atrAg+critAg+grvAg)/totalAg*100).toFixed(1):0)+'%',height:'100%',background:C.rose,opacity:.7,borderRadius:2}}/>
          </div>
        </div>

        {/* Bloco 3 — 4 mini cards por tipo */}
        <div style={{flex:'2.4',padding:'16px 20px'}}>
          <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:10}}>{(faltaDocs.length+atrDocs.length+critDocs.length+grvDocs.length).toLocaleString('pt-BR')} médicos com ocorrências</div>
          <div style={{display:'flex',gap:8}}>
            {[
              {label:'Falta',docs:faltaDocs.length,color:'#3B82F6'},
              {label:'Atraso',docs:atrDocs.length,color:C.amber},
              {label:'Grave',docs:grvDocs.length,color:C.orange},
              {label:'Crítico',docs:critDocs.length,color:C.rose},
            ].map(k=>(
              <div key={k.label} style={{flex:1,background:k.color+'10',border:'0.5px solid '+k.color+'28',borderRadius:8,padding:'8px 10px'}}>
                <div style={{fontSize:8,color:k.color,opacity:.8,marginBottom:4,fontWeight:700}}>{k.label}</div>
                <div style={{fontSize:20,fontWeight:800,color:k.color}}>{k.docs}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* FEED DE UNIDADES + PAINEL MÉDICOS */}
      <div style={{display:'grid',gridTemplateColumns:'1fr 340px',gap:14,marginBottom:14}}>

        {/* Feed de unidades */}
        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}}>
          {unidFilt&&(
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12,padding:'8px 12px',background:'rgba(245,158,11,0.08)',border:'0.5px solid rgba(245,158,11,0.25)',borderRadius:9}}>
              <div style={{width:6,height:6,borderRadius:'50%',background:C.amber,flexShrink:0}}/>
              <span style={{fontSize:11,color:C.amber,flex:1,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Filtrando: {unidFilt}</span>
              <button onClick={()=>setUnidFilt('')} style={{background:'transparent',border:'none',color:C.muted,cursor:'pointer',fontSize:12}}>✕</button>
            </div>
          )}
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:14}}>
            <div>
              <div style={{fontSize:13,fontWeight:700,color:C.text}}>Unidades por volume de agendas</div>
              <div style={{fontSize:10,color:C.muted,marginTop:2}}>{unidFilt?'clique ✕ para limpar':'clique na unidade para filtrar e ver médicos'}</div>
            </div>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:400,overflowY:'auto'}}>
            {feedList.slice(0,15).map((item,i)=>{
              const isSel=unidFilt===item.nm_local
              const pct=maxFeed>0?(item.total/maxFeed)*100:0
              const hasProb=item.faltas>0||item.atrasos>0
              return(
                <div key={i} onClick={()=>setUnidFilt(isSel?'':item.nm_local)} style={{
                  display:'flex',alignItems:'center',gap:10,padding:'9px 12px',borderRadius:10,cursor:'pointer',
                  background:isSel?'rgba(245,158,11,0.1)':'rgba(255,255,255,0.02)',
                  border:isSel?'1px solid rgba(245,158,11,0.4)':'0.5px solid rgba(255,255,255,0.05)',transition:'all .15s',
                }}
                  onMouseEnter={e=>{if(!isSel)e.currentTarget.style.background='rgba(255,255,255,0.04)'}}
                  onMouseLeave={e=>{if(!isSel)e.currentTarget.style.background='rgba(255,255,255,0.02)'}}>
                  <div style={{fontSize:10,fontWeight:800,color:i<3?C.amber:C.muted,minWidth:22,flexShrink:0}}>#{i+1}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:11,fontWeight:600,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{item.nm_local}</div>
                    <div style={{display:'flex',gap:8,marginTop:4}}>
                      <div style={{height:3,background:'rgba(255,255,255,0.06)',borderRadius:2,flex:1,overflow:'hidden'}}>
                        <div style={{height:'100%',borderRadius:2,background:i<3?C.amber:C.teal,width:pct+'%',transition:'width .5s'}}/>
                      </div>
                    </div>
                  </div>
                  <div style={{textAlign:'right',flexShrink:0}}>
                    <div style={{fontSize:12,fontWeight:700,color:i<3?C.amber:C.text}}>{item.total.toLocaleString('pt-BR')}</div>
                    <div style={{fontSize:8,color:C.muted}}>{item.consultas}c + {item.encaixe}e</div>
                  </div>
                  {hasProb&&(
                    <div style={{display:'flex',flexDirection:'column',gap:2,flexShrink:0}}>
                      {item.faltas>0&&<span style={{fontSize:8,padding:'1px 5px',borderRadius:10,background:'rgba(59,130,246,0.15)',color:'#3B82F6',border:'0.5px solid rgba(59,130,246,0.3)',whiteSpace:'nowrap'}}>{item.faltas} falta{item.faltas>1?'s':''}</span>}
                      {item.atrasos>0&&<span style={{fontSize:8,padding:'1px 5px',borderRadius:10,background:'rgba(244,63,94,0.12)',color:C.rose,border:'0.5px solid rgba(244,63,94,0.3)',whiteSpace:'nowrap'}}>{item.atrasos} atraso{item.atrasos>1?'s':''}</span>}
                    </div>
                  )}
                </div>
              )
            })}
          </div>
        </div>

        {/* Painel médicos + lista quando unidade filtrada */}
        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}}>
          <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:14}}>Médicos — Impacto nas Agendas</div>

          {/* 3 blocos: falta, crítico, grave */}
          <div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:12}}>
            {[
              {label:'Falta Médica',   docs:faltaDocs.length,ag:faltaAg,color:'#3B82F6'},
              {label:'Atraso 31–45min',docs:atrDocs.length,  ag:atrAg,  color:C.amber},
              {label:'Atraso Grave',   docs:grvDocs.length,  ag:grvAg,  color:C.orange},
              {label:'Atraso Crítico', docs:critDocs.length, ag:critAg, color:C.rose},
            ].map((k,i)=>(
              <div key={k.label} style={{background:k.color+'0A',border:'0.5px solid '+k.color+'20',borderRadius:8,padding:'10px 12px'}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
                  <div style={{display:'flex',alignItems:'center',gap:6}}>
                    <div style={{width:6,height:6,borderRadius:'50%',background:k.color}}/>
                    <span style={{fontSize:10,color:C.sub,fontWeight:600}}>{k.label}</span>
                  </div>
                  <span style={{fontSize:11,fontWeight:700,color:k.color}}>{k.docs} médico{k.docs!==1?'s':''}</span>
                </div>
                <div style={{display:'flex',justifyContent:'space-between',fontSize:9,color:C.muted}}>
                  <span>Agendas afetadas</span>
                  <span style={{fontWeight:700,color:k.color}}>{k.ag.toLocaleString('pt-BR')}</span>
                </div>
              </div>
            ))}
          </div>

          {/* Lista de médicos quando unidade filtrada */}
          {unidFilt&&(docsFaltaU.length>0||docsCritU.length>0||docsGrvU.length>0)&&(
            <div style={{paddingTop:12,borderTop:'0.5px solid rgba(255,255,255,0.06)'}}>
              <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.09em',marginBottom:8}}>Médicos · {unidFilt}</div>
              <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:180,overflowY:'auto'}}>
                {[...docsFaltaU.map(d=>({...d,color:'#3B82F6'})),...docsAtrU.map(d=>({...d,color:C.amber})),...docsGrvU.map(d=>({...d,color:C.orange})),...docsCritU.map(d=>({...d,color:C.rose}))].map((d,i)=>(
                  <div key={i} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0'}}>
                    <div style={{width:6,height:6,borderRadius:'50%',background:d.color,flexShrink:0}}/>
                    <span style={{fontSize:10,color:C.text,flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.nm}</span>
                    <span style={{fontSize:8,fontWeight:700,padding:'2px 7px',borderRadius:20,background:d.color+'18',color:d.color,border:'0.5px solid '+d.color+'30',flexShrink:0,whiteSpace:'nowrap'}}>{d.status}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          {unidFilt&&docsFaltaU.length===0&&docsCritU.length===0&&docsGrvU.length===0&&docsAtrU.length===0&&(
            <div style={{paddingTop:12,borderTop:'0.5px solid rgba(255,255,255,0.06)',fontSize:11,color:C.muted,textAlign:'center'}}>Sem ocorrências nesta unidade.</div>
          )}
          {!unidFilt&&faltaDocs.length===0&&critDocs.length===0&&grvDocs.length===0&&(
            <div style={{fontSize:11,color:C.muted,textAlign:'center',padding:'16px 0'}}>Nenhum problema registrado.</div>
          )}
        </div>
      </div>

      {/* GRÁFICO TENDÊNCIA */}
      <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'20px 24px',marginBottom:14}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14}}>
          <div>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.12em'}}>{'TENDÊNCIA DE AGENDAS · '+tRealCnt+' DIA'+(tRealCnt!==1?'S':'')}</div>
            <div style={{fontSize:10,color:C.muted,marginTop:3}}>Agendadas (amber) vs Entregues (verde) · área rosa = perdas por falta/atraso</div>
          </div>
          <div style={{display:'flex',background:'rgba(255,255,255,0.04)',border:'0.5px solid rgba(255,255,255,0.08)',borderRadius:9,padding:3,gap:2}}>
            {[{key:'real',label:'Real'},{key:'proj',label:'Projeção'}].map(v=>(
              <button key={v.key} onClick={()=>setTrendView(v.key)} style={{padding:'5px 14px',borderRadius:6,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .15s',background:trendView===v.key?C.amber:'transparent',color:trendView===v.key?'#1a0800':C.muted}}>{v.label}</button>
            ))}
          </div>
        </div>

        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10,marginBottom:16}}>
          {(trendView==='real'?[
            {label:'Agendadas hoje',   value:tLastReal?tLastReal.agendas.toLocaleString('pt-BR'):'—',   color:C.amber},
            {label:'Entregues hoje',   value:tLastReal?tLastDel.toLocaleString('pt-BR'):'—',            color:C.emerald},
            {label:'Gap acumulado',    value:tTotalGap.toLocaleString('pt-BR'),                         color:C.rose,   },
            {label:'Média diária',     value:tAvgAg.toLocaleString('pt-BR'),                            color:C.teal},
          ]:[
            {label:'Proj. agendadas amanhã', value:projData[0]?projData[0].agendas.toLocaleString('pt-BR'):'—',   color:C.amber},
            {label:'Proj. entregues amanhã', value:projData[0]?projData[0].delivered.toLocaleString('pt-BR'):'—', color:C.emerald},
            {label:'Tendência diária',       value:tSlope>=0?'+'+tSlope:String(tSlope),                          color:tSlope>0?C.rose:tSlope<0?C.emerald:C.muted},
            {label:'Gap proj. +5 dias',      value:projData[4]?projData[4].impactadas.toLocaleString('pt-BR'):'—',color:C.rose},
          ]).map((k,i)=>(
            <div key={k.label} style={{background:k.color+'12',border:'0.5px solid '+k.color+'20',borderRadius:9,padding:'10px 14px'}}>
              <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.07em',marginBottom:5}}>{k.label}</div>
              <div style={{fontSize:20,fontWeight:800,color:k.color}}>{k.value}</div>
            </div>
          ))}
        </div>

        {/* Tooltip */}
        {aTip&&(
          <div style={{position:'fixed',left:aTip.x+14,top:aTip.y-70,zIndex:999,pointerEvents:'none',background:'#0A0D16',border:'1px solid rgba(245,158,11,0.3)',borderRadius:10,padding:'10px 14px',boxShadow:'0 8px 32px rgba(0,0,0,0.6)',minWidth:150}}>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,marginBottom:6,textTransform:'uppercase',letterSpacing:'.08em'}}>{aTip.date?aTip.date.slice(5).replace('-','/'):'—'} {aTip.isProj?'· Projeção':''}</div>
            <div style={{display:'flex',flexDirection:'column',gap:4}}>
              <div style={{display:'flex',justifyContent:'space-between',gap:14,paddingBottom:4,borderBottom:'0.5px solid rgba(255,255,255,0.07)'}}>
                <span style={{fontSize:11,color:C.muted}}>Total agendas</span>
                <span style={{fontSize:13,fontWeight:800,color:C.amber}}>{(aTip.agendas||0).toLocaleString('pt-BR')}</span>
              </div>
              {!aTip.isProj&&(<>
                <div style={{display:'flex',justifyContent:'space-between',gap:14}}>
                  <span style={{fontSize:10,color:C.teal}}>Consultas</span>
                  <span style={{fontSize:11,fontWeight:700,color:C.teal}}>{(aTip.consultas||0).toLocaleString('pt-BR')}</span>
                </div>
                <div style={{display:'flex',justifyContent:'space-between',gap:14}}>
                  <span style={{fontSize:10,color:C.violet}}>Encaixe</span>
                  <span style={{fontSize:11,fontWeight:700,color:C.violet}}>{(aTip.encaixe||0).toLocaleString('pt-BR')}</span>
                </div>
              </>)}
              <div style={{display:'flex',justifyContent:'space-between',gap:14,paddingTop:4,borderTop:'0.5px solid rgba(255,255,255,0.07)'}}>
                <span style={{fontSize:11,color:C.rose}}>Afetadas</span>
                <span style={{fontSize:12,fontWeight:800,color:C.rose}}>{(aTip.impactadas||0).toLocaleString('pt-BR')}</span>
              </div>
            </div>
          </div>
        )}

        {trendAllData.length===0?(
          <div style={{textAlign:'center',padding:'32px 0',color:C.muted,fontSize:12}}>Sem dados suficientes.</div>
        ):(
          <div style={{width:'100%',position:'relative'}}>
            <div style={{display:'flex',gap:14,marginBottom:10,flexWrap:'wrap'}}>
              {[{color:'#F59E0B',label:'Agendadas'},{color:'#10B981',label:'Entregues (agendadas − problemas)'},{color:'rgba(244,63,94,0.4)',label:'Gap — perdas por falta/atraso',box:true}].concat(trendView==='proj'?[{color:'#F59E0B',label:'Projeção agendadas',dashed:true},{color:'#10B981',label:'Projeção entregues',dashed:true}]:[]).map(l=>(
                <div key={l.label} style={{display:'flex',alignItems:'center',gap:5}}>
                  <div style={{width:14,height:l.box?7:3,borderRadius:l.box?2:2,background:l.dashed?'transparent':l.color,border:l.dashed?'1.5px dashed '+l.color:'none',opacity:l.dashed?0.7:1}}/>
                  <span style={{fontSize:10,color:C.muted}}>{l.label}</span>
                </div>
              ))}
            </div>
            <svg width="100%" viewBox={'0 0 '+AVW+' '+AVH} style={{display:'block',overflow:'visible',cursor:'crosshair'}}
              onMouseMove={e=>{
                if(!trendAllData.length)return
                const rc=e.currentTarget.getBoundingClientRect()
                const mx=(e.clientX-rc.left)/rc.width*AVW
                let near=null,minD=Infinity
                trendAllData.forEach((d,i)=>{const xv=aNT<=1?APL+ACW/2:APL+i/(aNT-1)*ACW;const dist=Math.abs(mx-xv);if(dist<minD){minD=dist;near={...d,idx:i}}})
                if(near&&minD<(ACW/Math.max(aNT-1,1))*0.8)setATip({x:e.clientX,y:e.clientY,...near})
                else setATip(null)
              }}
              onMouseLeave={()=>setATip(null)}>

              <defs><clipPath id="aClip"><rect x={APL} y={APT-2} width={ACW} height={ACH+4}/></clipPath></defs>

              {/* Y grid */}
              {[0,Math.round(tMaxAg*0.33),Math.round(tMaxAg*0.66),tMaxAg].map((v,ti)=>(
                <g key={ti}>
                  <line x1={APL} y1={atyS(v)} x2={AVW-APR} y2={atyS(v)} stroke="rgba(255,255,255,0.05)" strokeWidth="0.5"/>
                  <text x={APL-7} y={atyS(v)+4} textAnchor="end" fontSize="9" fill="#475569">{v>=1000?(v/1000).toFixed(0)+'k':v}</text>
                </g>
              ))}

              {/* Separator real/proj */}
              {trendView==='proj'&&tRealCnt>0&&aNT>tRealCnt&&(
                <line x1={atxP(tRealCnt-1)} y1={APT} x2={atxP(tRealCnt-1)} y2={APT+ACH} stroke="rgba(245,158,11,0.22)" strokeWidth="1" strokeDasharray="4,3"/>
              )}

              <g clipPath="url(#aClip)">
                {/* Gap fill between agendadas and entregues */}
                {tRealCnt>=2&&(
                  <path d={atSmooth(byDate,'agendas',atyS,0)+' '+[...byDate].reverse().map((d,i)=>'L'+atxP(tRealCnt-1-i).toFixed(1)+','+atyS(d.delivered||0).toFixed(1)).join(' ')+' Z'} fill="rgba(244,63,94,0.13)"/>
                )}

                {/* Linha entregues — emerald */}
                {tRealCnt>=2&&(
                  <path d={atSmooth(byDate,'delivered',atyS,0)} fill="none" stroke="#10B981" strokeWidth="2.5"/>
                )}
                {/* Linha agendadas — amber */}
                {tRealCnt>=2&&(
                  <path d={atSmooth(byDate,'agendas',atyS,0)} fill="none" stroke="#F59E0B" strokeWidth="2.5"/>
                )}

                {/* Proj gap fill */}
                {trendView==='proj'&&projData.length>=2&&tRealCnt>=1&&(
                  <path d={'M'+atxP(tRealCnt-1).toFixed(1)+','+atyS(tLastReal?.agendas||0).toFixed(1)+atSmooth(projData,'agendas',atyS,tRealCnt).slice(1)+' '+[...projData].reverse().map((d,i)=>'L'+atxP(tRealCnt+projData.length-1-i).toFixed(1)+','+atyS(d.delivered||0).toFixed(1)).join(' ')+' Z'} fill="rgba(244,63,94,0.07)"/>
                )}
                {/* Proj agendadas dashed */}
                {trendView==='proj'&&projData.length>=1&&tRealCnt>=1&&(
                  <path d={'M'+atxP(tRealCnt-1).toFixed(1)+','+atyS(tLastReal?.agendas||0).toFixed(1)+atSmooth(projData,'agendas',atyS,tRealCnt).slice(1)} fill="none" stroke="#F59E0B" strokeWidth="1.5" strokeOpacity={0.45} strokeDasharray="6,4"/>
                )}
                {/* Proj entregues dashed */}
                {trendView==='proj'&&projData.length>=1&&tRealCnt>=1&&(
                  <path d={'M'+atxP(tRealCnt-1).toFixed(1)+','+atyS(tLastReal?.delivered||0).toFixed(1)+atSmooth(projData,'delivered',atyS,tRealCnt).slice(1)} fill="none" stroke="#10B981" strokeWidth="1.5" strokeOpacity={0.4} strokeDasharray="5,4"/>
                )}
              </g>

              {/* Dots agendadas */}
              {byDate.map((d,i)=>(<circle key={'a'+i} cx={atxP(i)} cy={atyS(d.agendas||0)} r="4" fill="#F59E0B" stroke="#06080F" strokeWidth="1.5"/>))}
              {/* Dots entregues */}
              {byDate.map((d,i)=>(<circle key={'e'+i} cx={atxP(i)} cy={atyS(d.delivered||0)} r="4" fill="#10B981" stroke="#06080F" strokeWidth="1.5"/>))}
              {/* Proj triangles */}
              {trendView==='proj'&&projData.map((d,i)=>{
                const cx=atxP(i+tRealCnt),cy=atyS(d.agendas||0)
                return(<polygon key={'t'+i} points={cx.toFixed(1)+','+(cy-7).toFixed(1)+' '+(cx-5).toFixed(1)+','+(cy+4).toFixed(1)+' '+(cx+5).toFixed(1)+','+(cy+4).toFixed(1)} fill="#F59E0B" fillOpacity={0.5}/>)
              })}
              {/* Hover highlight */}
              {aTip&&aTip.idx!=null&&(
                <circle cx={aNT<=1?APL+ACW/2:APL+(aTip.idx)/(aNT-1)*ACW} cy={atyS(aTip.agendas||0)} r="7" fill="none" stroke="#F59E0B" strokeWidth="2" strokeOpacity={0.8}/>
              )}
              {/* X labels */}
              {trendAllData.map((d,i)=>{const show=aNT<=10||i===0||i===aNT-1||i%Math.ceil(aNT/8)===0;return show?(<text key={'x'+i} x={atxP(i)} y={AVH-3} textAnchor="middle" fontSize="9" fill={d.isProj?'rgba(245,158,11,0.4)':'#334155'}>{d.date.slice(5).replace('-','/')}</text>):null})}
              <line x1={APL} y1={APT+ACH} x2={AVW-APR} y2={APT+ACH} stroke="rgba(255,255,255,0.07)" strokeWidth="0.5"/>
            </svg>
          </div>
        )}
      </div>

      {/* TOP UNIDADES + TOP ESPECIALIDADES */}
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:14}}>🏥 Top Unidades por Agendas</div>
          {feedList.slice(0,8).map((u,i)=>{
            const pct=maxFeed>0?(u.total/maxFeed)*100:0
            return(
              <div key={u.nm_local} style={{marginBottom:10}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4,gap:8}}>
                  <div style={{display:'flex',alignItems:'center',gap:6,minWidth:0}}>
                    <span style={{fontSize:9,fontWeight:800,color:i===0?C.amber:C.muted,minWidth:20,flexShrink:0}}>#{i+1}</span>
                    <span style={{fontSize:11,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{u.nm_local}</span>
                  </div>
                  <span style={{fontSize:11,fontWeight:700,color:i<3?C.amber:C.teal,flexShrink:0}}>{u.total.toLocaleString('pt-BR')}</span>
                </div>
                <div style={{display:'flex',gap:1,height:5,borderRadius:3,overflow:'hidden',background:'rgba(255,255,255,0.04)'}}>
                  <div style={{width:(u.consultas/Math.max(u.total,1)*pct)+'%',background:C.teal,transition:'width .5s'}}/>
                  <div style={{width:(u.encaixe/Math.max(u.total,1)*pct)+'%',background:C.violet,transition:'width .5s'}}/>
                </div>
              </div>
            )
          })}
          <div style={{display:'flex',gap:12,marginTop:8}}>
            <div style={{display:'flex',alignItems:'center',gap:4}}><div style={{width:8,height:8,borderRadius:2,background:C.teal}}/><span style={{fontSize:9,color:C.muted}}>Consultas</span></div>
            <div style={{display:'flex',alignItems:'center',gap:4}}><div style={{width:8,height:8,borderRadius:2,background:C.violet}}/><span style={{fontSize:9,color:C.muted}}>Encaixe</span></div>
          </div>
        </div>

        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:'.1em',marginBottom:14}}>🩺 Top Especialidades por Pacientes</div>
          {topSpec.map((s,i)=>{
            const pct=topSpec[0]?.total>0?(s.total/topSpec[0].total)*100:0
            return(
              <div key={s.name} style={{marginBottom:10}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4,gap:8}}>
                  <div style={{display:'flex',alignItems:'center',gap:6,minWidth:0}}>
                    <span style={{fontSize:9,fontWeight:800,color:i===0?C.amber:C.muted,minWidth:20,flexShrink:0}}>#{i+1}</span>
                    <span style={{fontSize:11,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.name}</span>
                  </div>
                  <span style={{fontSize:11,fontWeight:700,color:i<3?C.amber:C.cyan,flexShrink:0}}>{s.total.toLocaleString('pt-BR')}</span>
                </div>
                <div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:4,overflow:'hidden'}}>
                  <div style={{height:'100%',background:i<3?C.amber:C.cyan,width:pct+'%',transition:'width .6s ease',borderRadius:3}}/>
                </div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

function TabEspera({rows}){
  const[periodo,setPeriodo]=useState('MES')
  const[dateFrom,setDateFrom]=useState('')
  const[dateTo,setDateTo]=useState('')
  const[ufFilt,setUfFilt]=useState('TODOS')
  const[search,setSearch]=useState('')
const[horaFilt,setHoraFilt]=useState('TODAS')
const[horaFiltFim,setHoraFiltFim]=useState('TODAS')
const[classFiltro,setClassFiltro]=useState('TODAS')
  const[unidFilt,setUnidFilt]=useState('')
  const[justModal,setJustModal]=useState(false)
  const[justificativas,setJustificativas]=useState({})
  const[justLoading,setJustLoading]=useState(false)
  const[trendView,setTrendView]=useState('real')
  const[tTip,setTTip]=useState(null)

  const allDates=useMemo(()=>[...new Set(rows.map(r=>r.data_agenda).filter(Boolean))].sort(),[rows])
  const periodoFn=useMemo(()=>buildFilter(allDates,periodo,dateFrom,dateTo),[allDates,periodo,dateFrom,dateTo])
  const ufs=useMemo(()=>[...new Set(rows.map(r=>r.uf).filter(Boolean))].sort(),[rows])
  const dataRef=useMemo(()=>allDates.length?allDates[allDates.length-1]:'',[allDates])

  useEffect(()=>{
    if(!dataRef)return
    fetch(`${SB_URL}/rest/v1/justificativas_espera?select=hora,quantidade&data_ref=eq.${dataRef}&order=hora.asc`,{headers:SBH})
      .then(r=>r.json()).then(data=>{if(!Array.isArray(data))return;const map={};data.forEach(r=>{map[r.hora]=r.quantidade});setJustificativas(map)}).catch(console.error)
  },[dataRef])

  const saveJustificativa=useCallback(async(hora,qtd)=>{
    if(!dataRef)return;setJustLoading(true)
    try{await fetch(`${SB_URL}/rest/v1/justificativas_espera`,{method:'POST',headers:{...SBH,'Prefer':'resolution=merge-duplicates'},body:JSON.stringify({data_ref:dataRef,hora:parseInt(hora),quantidade:parseInt(qtd)||0})});setJustificativas(prev=>({...prev,[hora]:parseInt(qtd)||0}))}catch(e){console.error(e)}
    setJustLoading(false)
  },[dataRef])

  const horasDisp=useMemo(()=>{
    const set=new Set()
    rows.filter(d=>periodoFn(d.data_agenda)&&d.tempo_espera_min>=15).forEach(d=>{const h=d.hr_registro_espera_min;if(h==null)return;const hora=Math.floor(h/60);if(hora>=0&&hora<=23)set.add(hora)})
    return[...set].sort((a,b)=>a-b)
  },[rows,periodoFn])

  const filtered=useMemo(()=>{
    let r=rows.filter(d=>periodoFn(d.data_agenda))
    if(ufFilt!=='TODOS')r=r.filter(d=>d.uf===ufFilt)
    if(search){const q=search.toLowerCase();r=r.filter(d=>[d.nm_local,d.nm_medico,d.cidade].some(v=>String(v||'').toLowerCase().includes(q)))}
    if(horaFilt!=='TODAS'){const ini=parseInt(horaFilt,10),fim=horaFiltFim==='TODAS'?ini:parseInt(horaFiltFim,10);r=r.filter(d=>{const h=d.hr_registro_espera_min;if(h==null)return false;const hora=Math.floor(h/60);return hora>=ini&&hora<=fim})}
    if(classFiltro!=='TODAS'){
  r=r.filter(d=>{
    const t=d.tempo_espera_min
    if(t==null)return false

    if(classFiltro==='MODERADA')return t>=15&&t<=30
    if(classFiltro==='GRAVE')return t>=31&&t<=89
    if(classFiltro==='CRITICA')return t>=90

    return true
  })
}
    if(unidFilt)r=r.filter(d=>d.nm_local===unidFilt)
    return r
  },[rows,periodoFn,ufFilt,search,horaFilt,horaFiltFim,unidFilt])

  const espStats=useMemo(()=>{
    const comEsp=filtered.filter(d=>d.tempo_espera_min!=null&&d.tempo_espera_min>=15)
    const totalReg=filtered.length
    const totalPac=filtered.reduce((a,d)=>a+(d.qt_pacientes_aguardando||0),0)

    // Agrupa por unidade+hora (mesma lógica do feed) — base para TODOS os counts
    const grupoMap={}
    comEsp.forEach(d=>{
      const h=d.hr_registro_espera_min;if(h==null)return
      const hora=Math.floor(h/60),mins=Math.floor(h%60);if(hora<0||hora>23)return
      const horaStr=String(hora).padStart(2,'0')+':'+String(mins).padStart(2,'0')
      const unidade=d.nm_local||'Sem Unidade',key=horaStr+'||'+unidade
      if(!grupoMap[key])grupoMap[key]={horaStr,hora,nm_local:unidade,uf:d.uf||'',cidade:d.cidade||'',maxTempo:0,pac:0,count:0}
      if(d.tempo_espera_min>grupoMap[key].maxTempo){grupoMap[key].maxTempo=d.tempo_espera_min;grupoMap[key].pac=d.qt_pacientes_aguardando||0}
      grupoMap[key].count+=1
    })
    const incidentes=Object.values(grupoMap)
    const feedList=incidentes
  .slice()
  .sort((a,b)=>b.maxTempo-a.maxTempo)

    // Counts por incidente único (não por linha bruta)
    const modCnt =incidentes.filter(g=>g.maxTempo>=15&&g.maxTempo<=30).length
    const grvCnt =incidentes.filter(g=>g.maxTempo>30&&g.maxTempo<=89).length
    const critCnt=incidentes.filter(g=>g.maxTempo>=90).length
    const totalEsp=modCnt+grvCnt+critCnt

    // Top unidades por incidentes únicos
    const uMap={}
    incidentes.forEach(g=>{
      const u=g.nm_local
      if(!uMap[u])uMap[u]={n:u,total:0,crit:0,grv:0,mod:0}
      uMap[u].total++
      if(g.maxTempo>=90)uMap[u].crit++;else if(g.maxTempo>=31)uMap[u].grv++;else uMap[u].mod++
    })
    const topU=Object.values(uMap).sort((a,b)=>b.crit-a.crit||b.grv-a.grv).slice(0,6)
    const faltasList=filtered.filter(d=>String(d.atraso||'').toUpperCase()==='FALTA')
    const atrasosList=filtered.filter(d=>{if(String(d.atraso||'').toUpperCase()!=='SIM')return false;const t=d.tempo_atraso_min;return t!=null&&Math.abs(t)>31})
    const sMap={};atrasosList.forEach(d=>{const s=d.status||'Sem Status';sMap[s]=(sMap[s]||0)+1})
    const statusAt=Object.entries(sMap).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v)
    // Agrupa por data+unidade+hora (mesma lógica do feed) para contar incidentes únicos
    const gMap={}
    comEsp.forEach(d=>{
      const dt=d.data_agenda;if(!dt)return
      const h=d.hr_registro_espera_min;if(h==null)return
      const hora=Math.floor(h/60)
      const unidade=d.nm_local||'?'
      const key=dt+'||'+String(hora).padStart(2,'0')+'||'+unidade
      if(!gMap[key])gMap[key]={date:dt,maxTempo:0,pac:0}
      if(d.tempo_espera_min>gMap[key].maxTempo){gMap[key].maxTempo=d.tempo_espera_min;gMap[key].pac=d.qt_pacientes_aguardando||0}
    })
    const dMap={}
    Object.values(gMap).forEach(g=>{
      const dt=g.date
      if(!dMap[dt])dMap[dt]={date:dt,mod:0,grv:0,crit:0,pac:0}
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
    const lastMod =byDate[byDate.length-1]?.mod||0
    const lastGrv =byDate[byDate.length-1]?.grv||0
    const lastCrt =byDate[byDate.length-1]?.crit||0
    const projData=lastDate?Array.from({length:5},(_,i)=>({
      date:addDay(lastDate,i+1),
      mod: Math.max(0,Math.round(lastMod+slopeMod*(i+1))),
      grv: Math.max(0,Math.round(lastGrv+slopeGrv*(i+1))),
      crit:Math.max(0,Math.round(lastCrt+slopeCrt*(i+1))),
      total:Math.max(0,Math.round((lastMod+slopeMod*(i+1))+(lastGrv+slopeGrv*(i+1))+(lastCrt+slopeCrt*(i+1)))),
      pac:0,isProj:true,
    })):[]
    return{totalReg,totalPac,modCnt,grvCnt,critCnt,totalEsp,feedList,topU,faltasList,atrasosList,statusAt,byDate,projData}
  },[filtered])

  const totalJust=useMemo(()=>Object.values(justificativas).reduce((a,v)=>a+(parseInt(v)||0),0),[justificativas])
  const metaPct=espStats.totalEsp>0?Math.min(Math.round((totalJust/espStats.totalEsp)*100),100):0
  const metaColor=metaPct>=80?C.emerald:metaPct>=50?C.amber:C.rose
  const{totalReg,totalPac,modCnt,grvCnt,critCnt,totalEsp,feedList,topU,faltasList,atrasosList,statusAt,byDate,projData}=espStats
  const medTotalProb=faltasList.length+atrasosList.length
  const medFPct=medTotalProb>0?Math.round(faltasList.length/medTotalProb*100):0
  const medAPct=medTotalProb>0?Math.round(atrasosList.length/medTotalProb*100):0
  const trendAllData =trendView==='real'?byDate:[...byDate,...projData]
  const trendMaxAll  =Math.max(...trendAllData.map(d=>Math.max(d.mod||0,d.grv||0,d.crit||0)),1)
  const trendRealCnt =byDate.length
  const trendLastReal=byDate[byDate.length-1]
  const trendMaxDay  =byDate.length>0?byDate.reduce((a,d)=>(d.crit||0)>(a.crit||0)?d:a,byDate[0]):null
  const trendSlope   =projData.length>0?((projData[0]?.crit||0)-(trendLastReal?.crit||0)):0

  // SVG line chart vars
  const VW=800,VH=160,PL=50,PR=20,PT=10,PB=34
  const CW=VW-PL-PR,CH=VH-PT-PB
  const nTotal=trendAllData.length
  const txP=i=>nTotal<=1?PL+CW/2:PL+i/(nTotal-1)*CW
  const tyA=(v)=>PT+CH-(trendMaxAll>0?(v||0)/trendMaxAll*CH:0)
  const tSmooth=(data,key,scFn,off)=>{
    if(!data.length)return ''
    const pts=data.map((d,i)=>({x:parseFloat(txP(i+(off||0)).toFixed(2)),y:parseFloat(scFn(d[key]||0).toFixed(2))}))
    if(pts.length===1)return 'M'+pts[0].x+','+pts[0].y
    let p='M'+pts[0].x+','+pts[0].y
    for(let i=1;i<pts.length;i++){
      const p0=pts[i-2]||pts[i-1],p1=pts[i-1],p2=pts[i],p3=pts[i+1]||pts[i]
      const c1x=(p1.x+(p2.x-p0.x)/4).toFixed(2),c1y=(p1.y+(p2.y-p0.y)/4).toFixed(2)
      const c2x=(p2.x-(p3.x-p1.x)/4).toFixed(2),c2y=(p2.y-(p3.y-p1.y)/4).toFixed(2)
      p+=' C'+c1x+','+c1y+' '+c2x+','+c2y+' '+p2.x+','+p2.y
    }
    return p
  }

  // Lista de médicos filtrados pela unidade selecionada
  const docsFalta  = unidFilt ? Object.entries(faltasList.reduce((a,d)=>{const nm=d.nm_medico||'—';a[nm]=(a[nm]||0)+1;return a},{})).map(([nm,cnt])=>({nm,cnt})).sort((a,b)=>b.cnt-a.cnt) : []
  const docsAtraso = unidFilt ? Object.entries(atrasosList.reduce((a,d)=>{const nm=d.nm_medico||'—';if(!a[nm])a[nm]={cnt:0,status:d.status};a[nm].cnt++;return a},{})).map(([nm,v])=>({nm,...v})).sort((a,b)=>b.cnt-a.cnt) : []

  if(!rows.length)return(<div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'60vh',gap:14}}><div style={{fontSize:44}}>⏱️</div><div style={{fontSize:18,fontWeight:700,color:C.text}}>Nenhum dado de espera</div><div style={{fontSize:13,color:C.muted}}>Use o botão para carregar uma planilha</div></div>)

  const horasDispFim=horaFilt==='TODAS'?[]:horasDisp.filter(h=>h>parseInt(horaFilt))

  return(
    <div>
      {justModal&&(
        <div style={{position:'fixed',inset:0,zIndex:200,background:'rgba(0,0,0,0.75)',backdropFilter:'blur(4px)',display:'flex',alignItems:'center',justifyContent:'center'}} onClick={()=>setJustModal(false)}>
          <div style={{background:'#0A0D16',border:`1px solid rgba(245,158,11,0.3)`,borderRadius:16,padding:'24px 28px',minWidth:360,maxWidth:460,boxShadow:'0 24px 80px rgba(0,0,0,0.6)'}} onClick={e=>e.stopPropagation()}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
              <div><div style={{fontSize:14,fontWeight:700,color:C.text}}>Registrar Justificativas</div><div style={{fontSize:10.5,color:C.muted,marginTop:3}}>Meta: 80% das esperas com retorno · {dataRef}</div></div>
              <button onClick={()=>setJustModal(false)} style={{background:'transparent',border:'none',color:C.muted,fontSize:18,cursor:'pointer'}}>✕</button>
            </div>
            <div style={{background:'rgba(255,255,255,0.03)',border:'0.5px solid rgba(255,255,255,0.07)',borderRadius:10,padding:'12px 14px',marginBottom:18}}>
              <div style={{display:'flex',justifyContent:'space-between',marginBottom:7}}><span style={{fontSize:11,color:C.muted}}>Progresso atual</span><span style={{fontSize:12,fontWeight:700,color:metaColor}}>{metaPct}% de 80%</span></div>
              <div style={{background:'rgba(255,255,255,0.06)',borderRadius:4,height:8,overflow:'hidden'}}><div style={{height:'100%',width:`${metaPct}%`,background:metaColor,borderRadius:4,transition:'width .4s ease'}}/></div>
              <div style={{display:'flex',justifyContent:'space-between',marginTop:6}}><span style={{fontSize:10,color:C.muted}}>{totalJust} justificativas</span><span style={{fontSize:10,color:C.muted}}>de {totalEsp} esperas</span></div>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:8,maxHeight:280,overflowY:'auto'}}>
              {horasDisp.length===0&&<div style={{color:C.muted,fontSize:12,textAlign:'center',padding:'16px 0'}}>Sem horas disponíveis.</div>}
              {horasDisp.map(h=>(
                <div key={h} style={{display:'flex',alignItems:'center',gap:10}}>
                  <div style={{width:44,fontSize:12,fontWeight:700,color:C.sub,fontFamily:'monospace',flexShrink:0}}>{String(h).padStart(2,'0')}:00</div>
                  <div style={{flex:1,background:'rgba(255,255,255,0.04)',borderRadius:6,height:4,overflow:'hidden'}}><div style={{height:'100%',borderRadius:6,background:C.amber,transition:'width .3s',width:`${Math.min(((justificativas[h]||0)/Math.max(totalEsp/Math.max(horasDisp.length,1),1))*100,100)}%`}}/></div>
                  <input type="number" min="0" value={justificativas[h]||''} placeholder="0" onChange={e=>saveJustificativa(h,e.target.value)} style={{width:60,background:'rgba(255,255,255,0.05)',border:`0.5px solid ${justificativas[h]>0?'rgba(245,158,11,0.4)':'rgba(255,255,255,0.1)'}`,borderRadius:7,color:justificativas[h]>0?C.amber:C.text,fontSize:12,fontWeight:700,padding:'5px 8px',outline:'none',textAlign:'center'}}/>
                  <span style={{fontSize:10,color:C.muted,minWidth:24}}>just.</span>
                </div>
              ))}
            </div>
            {justLoading&&<div style={{textAlign:'center',marginTop:12,fontSize:11,color:C.amber}}>Salvando…</div>}
          </div>
        </div>
      )}

      <div style={{marginBottom:16}}>
        <PeriodoBar value={periodo} onChange={p=>{setPeriodo(p);setDateFrom('');setDateTo('');setUnidFilt('');setHoraFilt('TODAS');setHoraFiltFim('TODAS')}} allDates={allDates} dateFrom={dateFrom} dateTo={dateTo} onDateFrom={setDateFrom} onDateTo={setDateTo} label={`${totalReg.toLocaleString('pt-BR')} registros`}/>
      </div>
      <SearchBar search={search} onSearch={setSearch} uf={ufFilt} onUf={setUfFilt} ufs={ufs} showClear={ufFilt!=='TODOS'||!!search} onClear={()=>{setUfFilt('TODOS');setSearch('')}}/>

      {/* META */}
      <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:12,padding:'12px 18px',marginBottom:14}}>
        <div style={{display:'flex',alignItems:'center',gap:16}}>
          <span style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.1em',whiteSpace:'nowrap'}}>Meta 80%</span>
          <div style={{flex:1,display:'flex',flexDirection:'column',gap:4}}>
            <div style={{background:'rgba(255,255,255,0.06)',borderRadius:4,height:8,overflow:'hidden',position:'relative'}}>
              <div style={{height:'100%',borderRadius:4,width:`${metaPct}%`,transition:'width .6s ease',background:metaPct>=80?'#10B981':metaPct>=50?C.amber:C.rose}}/>
              <div style={{position:'absolute',top:0,bottom:0,left:'80%',width:2,background:'rgba(255,255,255,0.35)'}}/>
            </div>
            <div style={{display:'flex',justifyContent:'space-between'}}>
              <span style={{fontSize:9,color:C.muted}}>{totalJust} just. de {totalEsp} esperas ≥ 15min</span>
              <span style={{fontSize:9,color:C.muted}}>← 80%</span>
            </div>
          </div>
          <span style={{fontSize:15,fontWeight:800,color:metaColor,whiteSpace:'nowrap',minWidth:36,textAlign:'right'}}>{metaPct}%</span>
          <div style={{width:1,height:36,background:'rgba(255,255,255,0.07)',flexShrink:0}}/>
          {[{label:'Moderada',value:modCnt,sub:'15–30min',color:'#F59E0B'},{label:'Grave',value:grvCnt,sub:'31–1h29',color:'#F97316'},{label:'Crítica',value:critCnt,sub:'+1h30',color:'#F43F5E'}].map((k,i)=>(
            <div key={k.label} style={{display:'flex',flexDirection:'column',alignItems:'center',gap:2,minWidth:58,paddingLeft:i>0?14:0,borderLeft:i>0?'0.5px solid rgba(255,255,255,0.06)':'none'}}>
              <span style={{fontSize:9,fontWeight:700,color:k.color,textTransform:'uppercase',letterSpacing:'.07em'}}>{k.label}</span>
              <span style={{fontSize:20,fontWeight:800,color:k.color,lineHeight:1,letterSpacing:'-.5px'}}>{k.value.toLocaleString('pt-BR')}</span>
              <span style={{fontSize:8,color:C.muted}}>{k.sub}</span>
            </div>
          ))}
          <div style={{width:1,height:36,background:'rgba(255,255,255,0.07)',flexShrink:0}}/>
          <button onClick={()=>setJustModal(true)} style={{background:'rgba(245,158,11,0.1)',border:'0.5px solid rgba(245,158,11,0.3)',borderRadius:8,color:C.amber,fontSize:11,fontWeight:700,padding:'6px 12px',cursor:'pointer',whiteSpace:'nowrap'}}>+ Registrar</button>
        </div>
      </div>

      {/* FEED + MÉDICOS */}
      <div style={{display:'grid',gridTemplateColumns:'1fr 320px',gap:14,marginBottom:14}}>
        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'20px 22px'}}>
          {unidFilt&&(<div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12,padding:'8px 12px',background:'rgba(245,158,11,0.08)',border:'0.5px solid rgba(245,158,11,0.25)',borderRadius:9}}><div style={{width:6,height:6,borderRadius:'50%',background:C.amber,flexShrink:0}}/><span style={{fontSize:11,color:C.amber,flex:1,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Filtrando: {unidFilt}</span><button onClick={()=>setUnidFilt('')} style={{background:'transparent',border:'none',color:C.muted,cursor:'pointer',fontSize:12}}>✕</button></div>)}
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
            <div>
              <div style={{fontSize:13,fontWeight:700,color:C.text}}>Feed de Esperas por Hora</div>
              <div style={{fontSize:10.5,color:C.muted,marginTop:3}}>TEMPO_DE_ESPERA · hora via HR_REGISTRO_ESPERA · {unidFilt?'clique ✕ para limpar':'clique na unidade para filtrar'}</div>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:6,flexShrink:0}}>
              <select value={horaFilt} onChange={e=>{setHoraFilt(e.target.value);setHoraFiltFim('TODAS');setUnidFilt('')}} style={{background:'rgba(255,255,255,0.05)',border:'0.5px solid rgba(245,158,11,0.2)',borderRadius:8,color:C.text,fontSize:11,padding:'5px 8px',outline:'none',cursor:'pointer'}}>
                <option value="TODAS">Todas</option>
                {horasDisp.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}
              </select>
              <select
  value={classFiltro}
  onChange={e=>setClassFiltro(e.target.value)}
  style={{
    background:'rgba(255,255,255,0.05)',
    border:'0.5px solid rgba(245,158,11,0.2)',
    borderRadius:8,
    color:C.text,
    fontSize:11,
    padding:'5px 8px',
    outline:'none',
    cursor:'pointer'
  }}
>
  <option value="TODAS">Todas</option>
  <option value="MODERADA">Moderada</option>
  <option value="GRAVE">Grave</option>
  <option value="CRITICA">Crítica</option>
</select>
              {horaFilt!=='TODAS'&&(<><span style={{fontSize:11,color:C.muted}}>→</span><select value={horaFiltFim} onChange={e=>setHoraFiltFim(e.target.value)} style={{background:'rgba(255,255,255,0.05)',border:'0.5px solid rgba(245,158,11,0.2)',borderRadius:8,color:C.text,fontSize:11,padding:'5px 8px',outline:'none',cursor:'pointer'}}><option value="TODAS">{String(horaFilt).padStart(2,'0')}:00 só</option>{horasDispFim.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}:00</option>)}</select></>)}
            </div>
          </div>
          {feedList.length===0?(<div style={{textAlign:'center',padding:'32px 0',color:C.muted,fontSize:12}}>Sem esperas ≥ 15min no período/filtro selecionado.</div>):(
            <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:440,overflowY:'auto'}}>
              {feedList.map((item,i)=>{
                const cls=clsEspera(item.maxTempo),isCrit=item.maxTempo>=90,isGrv=item.maxTempo>=31&&item.maxTempo<90,isSel=unidFilt===item.nm_local
                return(<div key={i} onClick={()=>setUnidFilt(isSel?'':item.nm_local)} style={{display:'flex',alignItems:'center',gap:12,padding:'10px 14px',borderRadius:10,cursor:'pointer',background:isSel?`${cls.color}18`:isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':'rgba(255,255,255,0.02)',border:isSel?`1px solid ${cls.color}55`:`0.5px solid ${i<3?cls.border:'rgba(255,255,255,0.05)'}`,transition:'all .15s'}} onMouseEnter={e=>{if(!isSel)e.currentTarget.style.background=cls.bg}} onMouseLeave={e=>{if(!isSel)e.currentTarget.style.background=isCrit?'rgba(244,63,94,0.06)':isGrv?'rgba(249,115,22,0.04)':'rgba(255,255,255,0.02)'}}>
                  <div style={{width:8,height:8,borderRadius:'50%',background:cls.color,flexShrink:0,boxShadow:isCrit?`0 0 8px ${cls.color}`:'none'}}/>
                  <div style={{fontFamily:'monospace',fontSize:13,fontWeight:700,color:C.sub,flexShrink:0,minWidth:44}}>{item.horaStr}</div>
                  <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,color:C.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{item.nm_local}</div><div style={{fontSize:10,color:C.muted,marginTop:2}}>{[item.cidade,item.uf].filter(Boolean).join(' · ')}</div></div>
                  <div style={{textAlign:'center',flexShrink:0,minWidth:42}}><div style={{fontSize:12,fontWeight:700,color:'#0EA5E9'}}>{item.pac>0?item.pac:'—'}</div><div style={{fontSize:9,color:C.muted}}>pac.</div></div>
                  <div style={{fontSize:15,fontWeight:900,color:cls.color,flexShrink:0,minWidth:52,textAlign:'right'}}>{fmtMin(item.maxTempo)}</div>
                  <span style={{fontSize:9.5,fontWeight:700,padding:'3px 9px',borderRadius:20,background:cls.bg,color:cls.color,border:`0.5px solid ${cls.border}`,whiteSpace:'nowrap',flexShrink:0}}>{cls.label}</span>
                </div>)
              })}
            </div>
          )}
        </div>

        <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'18px 20px'}}>
          <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:14}}>Médicos — Falta e Atraso</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1px 1fr',gap:0,alignItems:'stretch'}}>
            <div style={{paddingRight:16,display:'flex',flexDirection:'column',gap:10}}>
              {[{label:'Faltas',value:faltasList.length,color:C.rose,pct:medFPct},{label:'Atrasos >31min',value:atrasosList.length,color:C.amber,pct:medAPct}].map(k=>(
                <div key={k.label}>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',marginBottom:5}}>
                    <div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:6,height:6,borderRadius:'50%',background:k.color}}/><span style={{fontSize:11,color:C.sub}}>{k.label}</span></div>
                    <span style={{fontSize:13,fontWeight:700,color:k.color}}>{k.value} <span style={{fontSize:9,color:C.muted,fontWeight:400}}>{k.pct}%</span></span>
                  </div>
                  <div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:4,overflow:'hidden'}}><div style={{height:'100%',background:k.color,width:`${k.pct}%`,borderRadius:3,transition:'width .6s'}}/></div>
                </div>
              ))}
              <div style={{display:'flex',alignItems:'center',gap:6,paddingTop:8,borderTop:'0.5px solid rgba(255,255,255,0.06)',marginTop:2}}>
                <span style={{fontSize:22,fontWeight:800,color:C.orange,lineHeight:1}}>{medTotalProb}</span>
                <span style={{fontSize:10,color:C.muted}}>total de ocorrências</span>
              </div>
            </div>
            <div style={{background:'rgba(255,255,255,0.07)',margin:'0 16px'}}/>
            <div style={{paddingLeft:16,display:'flex',flexDirection:'column',gap:8}}>
              <span style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.09em'}}>Classificação do Atraso</span>
              {statusAt.length===0&&<div style={{color:C.muted,fontSize:11}}>Nenhuma ocorrência.</div>}
              {statusAt.map(({k,v})=>{
                const cfg=getStatusCfg(k),pctAt=statusAt[0]?.v>0?Math.round(v/statusAt[0].v*100):0
                return(
                  <div key={k}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:3}}>
                      <span style={{fontSize:10,color:C.sub}}>{cfg.label}</span>
                      <div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontSize:12,fontWeight:700,color:cfg.color}}>{v}</span><span style={{fontSize:9,color:C.muted,fontWeight:400}}>{pctAt}%</span></div>
                    </div>
                    <div style={{background:'rgba(255,255,255,0.05)',borderRadius:3,height:5,overflow:'hidden'}}><div style={{height:'100%',background:`linear-gradient(90deg,${cfg.color},${cfg.color}88)`,width:`${pctAt}%`,borderRadius:3,transition:'width .6s'}}/></div>
                  </div>
                )
              })}
            </div>
          </div>
          {medTotalProb===0&&!unidFilt&&<div style={{color:C.muted,fontSize:11,textAlign:'center',padding:'8px 0',marginTop:10}}>Nenhuma ocorrência no período.</div>}

          {/* Lista de médicos quando unidade filtrada */}
          {unidFilt&&(docsFalta.length>0||docsAtraso.length>0)&&(
            <div style={{marginTop:12,paddingTop:12,borderTop:'0.5px solid rgba(255,255,255,0.06)'}}>
              <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.09em',marginBottom:8}}>Médicos · {unidFilt}</div>
              <div style={{display:'flex',flexDirection:'column',gap:5,maxHeight:140,overflowY:'auto'}}>
                {docsFalta.map(d=>(
                  <div key={'f'+d.nm} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0'}}>
                    <div style={{width:6,height:6,borderRadius:'50%',background:C.rose,flexShrink:0}}/>
                    <span style={{fontSize:10.5,color:C.text,flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.nm}</span>
                    <span style={{fontSize:9,fontWeight:700,padding:'1px 7px',borderRadius:20,background:'rgba(244,63,94,0.12)',color:C.rose,border:'0.5px solid rgba(244,63,94,0.3)',flexShrink:0}}>Falta</span>
                  </div>
                ))}
                {docsAtraso.map(d=>{
                  const cfg=getStatusCfg(d.status)
                  return(
                    <div key={'a'+d.nm} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0'}}>
                      <div style={{width:6,height:6,borderRadius:'50%',background:C.amber,flexShrink:0}}/>
                      <span style={{fontSize:10.5,color:C.text,flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.nm}</span>
                      <span style={{fontSize:9,fontWeight:700,padding:'1px 7px',borderRadius:20,background:`${cfg.color}18`,color:cfg.color,border:`0.5px solid ${cfg.color}30`,flexShrink:0,whiteSpace:'nowrap'}}>{cfg.label}</span>
                    </div>
                  )
                })}
              </div>
            </div>
          )}
          {unidFilt&&docsFalta.length===0&&docsAtraso.length===0&&<div style={{marginTop:10,fontSize:11,color:C.muted,textAlign:'center'}}>Sem ocorrências nesta unidade.</div>}
        </div>
      </div>

      {/* TENDÊNCIA */}
      <div style={{background:'rgba(255,255,255,0.025)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:14,padding:'20px 24px',marginBottom:14}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
          <div>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.13em'}}>{'TENDÊNCIA DE ESPERAS · '+trendRealCnt+' DIA'+(trendRealCnt!==1?'S':'')+' · TODOS OS ESTADOS'}</div>
            <div style={{fontSize:10.5,color:C.muted,marginTop:4}}>Esperas críticas ao longo do tempo</div>
          </div>
          <div style={{display:'flex',background:'rgba(255,255,255,0.04)',border:'0.5px solid rgba(255,255,255,0.08)',borderRadius:9,padding:3,gap:2}}>
            {[{key:'real',label:'Real'},{key:'proj',label:'Projeção'}].map(v=>(
              <button key={v.key} onClick={()=>setTrendView(v.key)} style={{padding:'5px 16px',borderRadius:6,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,transition:'all .15s',background:trendView===v.key?C.amber:'transparent',color:trendView===v.key?'#1a0800':C.muted}}>{v.label}</button>
            ))}
          </div>
        </div>

        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10,marginBottom:20}}>
          {(trendView==='real'?[
            {label:'Crítica hoje ≥90min',  value:String(trendLastReal?.crit||'—'),  color:C.rose},
            {label:'Grave hoje 31–89min',  value:String(trendLastReal?.grv||'—'),   color:C.orange},
            {label:'Moderada hoje 15–30min',value:String(trendLastReal?.mod||'—'),  color:C.amber},
            {label:'Pior dia (crítica)',   value:trendMaxDay?(trendMaxDay.crit+' em '+trendMaxDay.date.slice(5).replace('-','/')):'—',color:C.rose},
          ]:[
            {label:'Proj. crítica amanhã', value:String(projData[0]?.crit||'—'),    color:C.rose},
            {label:'Proj. grave amanhã',   value:String(projData[0]?.grv||'—'),     color:C.orange},
            {label:'Proj. moderada amanhã',value:String(projData[0]?.mod||'—'),     color:C.amber},
            {label:'Tendência crítica',    value:trendSlope>=0?'+'+trendSlope:String(trendSlope),color:trendSlope>0?C.rose:trendSlope<0?C.emerald:C.muted},
          ]).map((k,i)=>(
            <div key={k.label} style={{background:k.color+'14',border:'0.5px solid '+k.color+'22',borderRadius:10,padding:'11px 14px'}}>
              <div style={{fontSize:9,fontWeight:700,color:C.muted,textTransform:'uppercase',letterSpacing:'.08em',marginBottom:7}}>{k.label}</div>
              <div style={{fontSize:22,fontWeight:800,color:k.color,letterSpacing:'-.5px'}}>{k.value}</div>
            </div>
          ))}
        </div>

        {trendAllData.length===0?(
          <div style={{textAlign:'center',padding:'40px 0',color:C.muted,fontSize:12}}>Sem dados suficientes. Carregue uma planilha com múltiplos dias.</div>
        ):(
          <div>
            <div style={{display:'flex',gap:16,marginBottom:12,flexWrap:'wrap'}}>
              {[{color:C.rose,label:'Espera Crítica (real)',dashed:false},{color:C.orange,label:'Espera Grave (real)',dashed:false}].concat(trendView==='proj'?[{color:C.rose,label:'Projeção crítica',dashed:true}]:[]).map(l=>(
                <div key={l.label} style={{display:'flex',alignItems:'center',gap:5}}>
                  <div style={{width:16,height:3,borderRadius:2,background:l.dashed?'transparent':l.color,border:l.dashed?('1.5px dashed '+l.color):'none',opacity:l.dashed?0.7:1}}/>
                  <span style={{fontSize:10,color:C.muted}}>{l.label}</span>
                </div>
              ))}
            </div>
            {/* Tooltip */}
            {tTip&&(
              <div style={{position:'fixed',left:tTip.x+14,top:tTip.y-60,zIndex:999,pointerEvents:'none',
                background:'#0A0D16',border:'1px solid rgba(244,63,94,0.35)',borderRadius:10,
                padding:'10px 14px',boxShadow:'0 8px 32px rgba(0,0,0,0.6)',minWidth:140}}>
                <div style={{fontSize:10,fontWeight:700,color:C.muted,marginBottom:6,textTransform:'uppercase',letterSpacing:'.08em'}}>
                  {tTip.date?tTip.date.slice(5).replace('-','/'):'—'} {tTip.isProj?'· Projeção':''}
                </div>
                <div style={{display:'flex',flexDirection:'column',gap:5}}>
                  <div style={{display:'flex',justifyContent:'space-between',gap:16}}>
                    <span style={{fontSize:11,color:'#F43F5E'}}>● Crítica ≥90min</span>
                    <span style={{fontSize:13,fontWeight:800,color:'#F43F5E'}}>{tTip.crit||0}</span>
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',gap:16}}>
                    <span style={{fontSize:11,color:'#F97316'}}>● Grave 31–89min</span>
                    <span style={{fontSize:12,fontWeight:700,color:'#F97316'}}>{tTip.grv||0}</span>
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',gap:16}}>
                    <span style={{fontSize:11,color:'#F59E0B'}}>● Moderada 15–30min</span>
                    <span style={{fontSize:12,fontWeight:700,color:'#F59E0B'}}>{tTip.mod||0}</span>
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',gap:16,paddingTop:5,borderTop:'0.5px solid rgba(255,255,255,0.07)'}}>
                    <span style={{fontSize:10,color:C.muted}}>Total incidentes</span>
                    <span style={{fontSize:11,fontWeight:700,color:C.muted}}>{((tTip.crit||0)+(tTip.grv||0)+(tTip.mod||0))}</span>
                  </div>
                  {tTip.isProj&&<div style={{fontSize:9,color:C.amber,paddingTop:4,borderTop:'0.5px solid rgba(255,255,255,0.06)'}}>📊 Valor projetado</div>}
                </div>
              </div>
            )}
            <div style={{width:'100%',position:'relative'}}>
              <svg width="100%" viewBox={'0 0 '+VW+' '+VH} style={{display:'block',overflow:'visible',cursor:'crosshair'}}
                onMouseMove={e=>{
                  if(!trendAllData.length)return
                  const rc=e.currentTarget.getBoundingClientRect()
                  const mx=(e.clientX-rc.left)/rc.width*VW
                  let near=null,minD=Infinity
                  trendAllData.forEach((d,i)=>{
                    const xv=nTotal<=1?PL+CW/2:PL+i/(nTotal-1)*CW
                    const dist=Math.abs(mx-xv)
                    if(dist<minD){minD=dist;near={...d,idx:i}}
                  })
                  if(near&&minD<(CW/Math.max(nTotal-1,1))*0.8)setTTip({x:e.clientX,y:e.clientY,...near})
                  else setTTip(null)
                }}
                onMouseLeave={()=>setTTip(null)}>
                <defs>
                  <clipPath id="tClip">
                    <rect x={PL} y={PT-2} width={CW} height={CH+4}/>
                  </clipPath>
                </defs>

                {/* Y grid — single axis */}
                {[0,Math.round(trendMaxAll*0.33),Math.round(trendMaxAll*0.66),trendMaxAll].map((v,ti)=>(
                  <g key={ti}>
                    <line x1={PL} y1={tyA(v)} x2={VW-PR} y2={tyA(v)} stroke="rgba(255,255,255,0.05)" strokeWidth="0.5"/>
                    <text x={PL-7} y={tyA(v)+4} textAnchor="end" fontSize="9" fill="#475569">{v}</text>
                  </g>
                ))}

                {/* Separator real/proj */}
                {trendView==='proj'&&trendRealCnt>0&&nTotal>trendRealCnt&&(
                  <line x1={txP(trendRealCnt-1)} y1={PT} x2={txP(trendRealCnt-1)} y2={PT+CH} stroke="rgba(245,158,11,0.25)" strokeWidth="1" strokeDasharray="4,3"/>
                )}

                <g clipPath="url(#tClip)">
                  {/* Area fill crítica */}
                  {trendRealCnt>=2&&(
                    <path d={tSmooth(byDate,'crit',tyA,0)+' L'+txP(trendRealCnt-1).toFixed(1)+','+(PT+CH).toFixed(1)+' L'+txP(0).toFixed(1)+','+(PT+CH).toFixed(1)+' Z'} fill="rgba(244,63,94,0.10)"/>
                  )}
                  {/* Linha moderada — amber */}
                  {trendRealCnt>=2&&(
                    <path d={tSmooth(byDate,'mod',tyA,0)} fill="none" stroke="#F59E0B" strokeWidth="2" strokeOpacity={0.8}/>
                  )}
                  {/* Linha grave — orange */}
                  {trendRealCnt>=2&&(
                    <path d={tSmooth(byDate,'grv',tyA,0)} fill="none" stroke="#F97316" strokeWidth="2.5" strokeOpacity={0.9}/>
                  )}
                  {/* Linha crítica — rose, destaque */}
                  {trendRealCnt>=2&&(
                    <path d={tSmooth(byDate,'crit',tyA,0)} fill="none" stroke="#F43F5E" strokeWidth="3"/>
                  )}
                  {/* Proj moderada dashed */}
                  {trendView==='proj'&&projData.length>=1&&trendRealCnt>=1&&(
                    <path d={'M'+txP(trendRealCnt-1).toFixed(1)+','+tyA(trendLastReal?.mod||0).toFixed(1)+tSmooth(projData,'mod',tyA,trendRealCnt).slice(1)} fill="none" stroke="#F59E0B" strokeWidth="1.5" strokeOpacity={0.4} strokeDasharray="5,4"/>
                  )}
                  {/* Proj grave dashed */}
                  {trendView==='proj'&&projData.length>=1&&trendRealCnt>=1&&(
                    <path d={'M'+txP(trendRealCnt-1).toFixed(1)+','+tyA(trendLastReal?.grv||0).toFixed(1)+tSmooth(projData,'grv',tyA,trendRealCnt).slice(1)} fill="none" stroke="#F97316" strokeWidth="1.5" strokeOpacity={0.4} strokeDasharray="5,4"/>
                  )}
                  {/* Proj crítica dashed */}
                  {trendView==='proj'&&projData.length>=1&&trendRealCnt>=1&&(
                    <path d={'M'+txP(trendRealCnt-1).toFixed(1)+','+tyA(trendLastReal?.crit||0).toFixed(1)+tSmooth(projData,'crit',tyA,trendRealCnt).slice(1)} fill="none" stroke="#F43F5E" strokeWidth="2" strokeOpacity={0.45} strokeDasharray="6,4"/>
                  )}
                </g>

                {/* Dots moderada */}
                {byDate.map((d,i)=>(<circle key={'m'+i} cx={txP(i)} cy={tyA(d.mod||0)} r="3" fill="#F59E0B" stroke="#06080F" strokeWidth="1.5"/>))}
                {/* Dots grave */}
                {byDate.map((d,i)=>(<circle key={'g'+i} cx={txP(i)} cy={tyA(d.grv||0)} r="3.5" fill="#F97316" stroke="#06080F" strokeWidth="1.5"/>))}
                {/* Dots crítica — maiores */}
                {byDate.map((d,i)=>(<circle key={'c'+i} cx={txP(i)} cy={tyA(d.crit||0)} r="4" fill="#F43F5E" stroke="#06080F" strokeWidth="1.5"/>))}
                {/* Proj triangles crítica */}
                {trendView==='proj'&&projData.map((d,i)=>{
                  const cx=txP(i+trendRealCnt),cy=tyA(d.crit||0)
                  return(<polygon key={'t'+i} points={cx.toFixed(1)+','+(cy-7).toFixed(1)+' '+(cx-5).toFixed(1)+','+(cy+4).toFixed(1)+' '+(cx+5).toFixed(1)+','+(cy+4).toFixed(1)} fill="#F43F5E" fillOpacity={0.5}/>)
                })}
                {/* Hover highlight */}
                {tTip&&tTip.idx!=null&&(
                  <circle cx={nTotal<=1?PL+CW/2:PL+(tTip.idx)/(nTotal-1)*CW} cy={tyA(tTip.crit||0)} r="7" fill="none" stroke="#F43F5E" strokeWidth="2" strokeOpacity={0.8}/>
                )}
                {/* X labels */}
                {trendAllData.map((d,i)=>{
                  const show=nTotal<=10||i===0||i===nTotal-1||i%Math.ceil(nTotal/8)===0
                  return show?(<text key={'x'+i} x={txP(i)} y={VH-3} textAnchor="middle" fontSize="9" fill={d.isProj?'rgba(245,158,11,0.4)':'#334155'}>{d.date.slice(5).replace('-','/')}</text>):null
                })}
                <line x1={PL} y1={PT+CH} x2={VW-PR} y2={PT+CH} stroke="rgba(255,255,255,0.07)" strokeWidth="0.5"/>
            </svg>
            </div>
            {trendView==='proj'&&(
              <div style={{marginTop:12,padding:'9px 14px',background:'rgba(245,158,11,0.05)',border:'0.5px solid rgba(245,158,11,0.18)',borderRadius:8,fontSize:10.5,color:C.muted}}>
                <span style={{color:C.amber,fontWeight:700}}>Como a projeção é calculada — </span>
                {'Média dos últimos '+Math.min(byDate.length,3)+' dia'+(byDate.length!==1?'s':'')+' + tendência linear para os próximos 5 dias.'}
                {byDate.length<3&&<span style={{color:C.amber}}> Precisão aumenta com mais dias na base.</span>}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}

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
      all=all.concat(batch)
      if(batch.length<PAGE)break
      offset+=PAGE
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
      setStoreMsg(`${json.length.toLocaleString('pt-BR')} linhas — limpando banco…`)
      const delRes=await fetch('/api/save',{method:'DELETE'})
      if(!delRes.ok){setStoreMsg('⚠ Erro ao limpar banco');setStoring(false);return}
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
    await fetch('/api/save',{method:'DELETE'})
    setAgendas([]);setEspera([]);setStorageInfo({agendas:0,espera:0});setTimestamp('');setStoreMsg('');setStoring(false)
  }

  return(
    <div style={{minHeight:'100vh',background:C.bg,fontFamily:"'DM Sans','Segoe UI',sans-serif",color:C.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800;900&family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(245,158,11,0.2);border-radius:4px}
        select option{background:#0A0D16;color:#F1F5F9}
        input::placeholder{color:#475569}
        select{appearance:auto}
      `}</style>
      <div style={{background:'#070910',borderBottom:'1px solid rgba(245,158,11,0.1)',display:'flex',alignItems:'center',padding:'0 32px',height:52,position:'sticky',top:0,zIndex:100,backdropFilter:'blur(12px)'}}>
        <div style={{position:'absolute',top:-40,left:'50%',transform:'translateX(-50%)',width:500,height:80,background:'radial-gradient(ellipse,rgba(245,158,11,0.07),transparent 70%)',pointerEvents:'none'}}/>
        <div style={{display:'flex',alignItems:'center',gap:10,marginRight:32}}>
          <div style={{width:28,height:28,borderRadius:8,background:'linear-gradient(135deg,#F59E0B,#F97316)',display:'flex',alignItems:'center',justifyContent:'center'}}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
          </div>
          <div><span style={{fontSize:14,fontWeight:900,color:C.text,fontFamily:"'Syne',sans-serif",letterSpacing:'-.3px'}}>Monitor </span><span style={{fontSize:14,fontWeight:900,color:C.amber,fontFamily:"'Syne',sans-serif",letterSpacing:'-.3px'}}>Clínicas</span></div>
        </div>
        <div style={{display:'flex',height:'100%',gap:0}}>
          {[{key:'espera',label:'Fila de Espera',count:storageInfo.espera},{key:'agendas',label:'Agendas Médicas',count:storageInfo.agendas}].map(t=>(
            <button key={t.key} onClick={()=>setTab(t.key)} style={{padding:'0 20px',border:'none',background:'transparent',cursor:'pointer',fontSize:12,fontWeight:700,height:'100%',color:tab===t.key?C.amber:C.muted,borderBottom:tab===t.key?`2px solid ${C.amber}`:'2px solid transparent',transition:'all .2s',display:'flex',alignItems:'center',gap:7}}>
              {t.label}
              {t.count>0&&<span style={{fontSize:9,padding:'1px 6px',borderRadius:10,background:tab===t.key?`${C.amber}20`:'rgba(255,255,255,0.06)',color:tab===t.key?C.amber:C.muted,fontWeight:700}}>{t.count.toLocaleString('pt-BR')}</span>}
            </button>
          ))}
        </div>
        <div style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:14}}>
          <div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:6,height:6,borderRadius:'50%',background:C.emerald,boxShadow:`0 0 6px ${C.emerald}`}}/><span style={{fontSize:10,color:C.muted}}>Banco conectado</span></div>
          {timestamp&&<span style={{fontSize:10,color:C.muted}}>{timestamp}</span>}
          {storeMsg&&<span style={{fontSize:10,color:storeMsg.startsWith('☁')||storeMsg.startsWith('✓')?C.emerald:C.amber,fontWeight:600,maxWidth:220,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{storeMsg}</span>}
          {(storageInfo.agendas>0||storageInfo.espera>0)&&!storing&&(<button onClick={handleClear} style={{background:'transparent',border:'0.5px solid rgba(244,63,94,0.3)',borderRadius:8,color:C.rose,fontSize:11,padding:'5px 10px',cursor:'pointer'}}>🗑</button>)}
          <label style={{background:storing?'rgba(255,255,255,0.06)':'linear-gradient(135deg,#F59E0B,#F97316)',color:storing?C.muted:'#1a0800',fontWeight:800,fontSize:12,padding:'7px 16px',borderRadius:9,cursor:storing?'default':'pointer',transition:'all .2s',whiteSpace:'nowrap',fontFamily:"'Syne',sans-serif"}}>
            {storing?'Salvando…':'+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={storing||loadingDB}/>
          </label>
        </div>
      </div>
      <div style={{padding:'24px 32px'}}>
        {loadingDB?(
          <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',height:'calc(100vh - 120px)',gap:16}}>
            <div style={{fontSize:36,filter:'drop-shadow(0 0 20px #F59E0B88)'}}>⏳</div>
            <div style={{fontSize:16,color:C.sub,fontFamily:"'Syne',sans-serif",fontWeight:700}}>Conectando ao banco…</div>
            {storeMsg&&<div style={{fontSize:12,color:C.amber}}>{storeMsg}</div>}
          </div>
        ):(<>
          {tab==='agendas'&&<TabAgendas rows={agendas}/>}
          {tab==='espera'&&<TabEspera rows={espera}/>}
        </>)}
      </div>
    </div>
  )
}
