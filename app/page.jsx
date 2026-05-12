'use client'

// ─────────────────────────────────────────────────────────
//  SUPABASE — substitua pelas suas credenciais
// ─────────────────────────────────────────────────────────
const SUPABASE_URL      = 'https://fwdvzsywudpieqlqnxkp.supabase.co'
const SUPABASE_ANON_KEY = 'sb_publishable_x32NVeFMKLK9kLJfdunngg_GfxpTo1P'

import { useMemo, useState, useEffect, useCallback } from 'react'
import * as XLSX from 'xlsx'

// ─── Palette ──────────────────────────────────────────────
const T = {
  bg:      '#06080F', surface: '#0C1220', card:    '#0F1929',
  card2:   '#0D1525', border:  '#182840', accent:  '#00C6FF',
  accentB: '#0072FF', danger:  '#FF4D6A', warning: '#FFB340',
  success: '#00E5A0', text:    '#EDF2FF', muted:   '#4A6A88',
  sub:     '#7FA8C4',
}

// ─── Status config ────────────────────────────────────────
const STATUS_CFG = [
  { key: 'ATRASO',            bg: '#D97706', label: 'Atraso',         desc: '31 min ‹ atraso ‹ 45 min'    },
  { key: 'ATRASO CRÍTICO',    bg: '#EA580C', label: 'Atraso Crítico', desc: '46 min › atraso ‹ 1h30'      },
  { key: 'ATRASO GRAVE',      bg: '#DC2626', label: 'Atraso Grave',   desc: 'Atraso › 1h30'               },
  { key: 'Falta Médica',      bg: '#1D4ED8', label: 'Médico Faltou',  desc: 'Ausência registrada'         },
  { key: 'Remarcação Adm',    bg: '#7C3AED', label: 'Remarcação Adm', desc: 'Remarcação administrativa'   },
  { key: 'Remarcação Médico', bg: '#9333EA', label: 'Remanejamento',  desc: 'Remarcação pelo médico'      },
  { key: 'Remarcação médico', bg: '#9333EA', label: 'Remanejamento',  desc: 'Remarcação pelo médico'      },
  { key: 'SEM_PONTO',         bg: '#374151', label: 'Sem Ponto',      desc: 'HR_ENTRADA vazia'            },
  { key: 'OK',                bg: '#059669', label: 'Motivo (OK)',    desc: 'Atendimento dentro do prazo' },
]
const getCfg = (key) =>
  STATUS_CFG.find(s => s.key === key) || { bg: T.muted, label: key, desc: '' }

// ─── Helpers ─────────────────────────────────────────────
const parseHM = (v) => {
  if (!v && v !== 0) return 0
  if (typeof v === 'number') return v * 24
  const s = String(v).trim(), sign = s.startsWith('-') ? -1 : 1
  if (s.includes(':')) {
    const p = s.replace('-','').split(':').map(Number)
    return sign * ((p[0]||0) + (p[1]||0)/60)
  }
  return isNaN(parseFloat(s)) ? 0 : parseFloat(s)
}
const fmtH = (h) => {
  const a = Math.abs(h), hh = Math.floor(a), mm = Math.round((a-hh)*60)
  return `${hh}h${mm > 0 ? ` ${mm}m` : ''}`
}

// Serial Excel → 'YYYY-MM-DD' via UTC (sem shift de fuso)
const serialToDateStr = (v) => {
  if (!v && v !== 0) return ''
  let d
  if (typeof v === 'number') {
    d = new Date(Math.round((v - 25569) * 86400000))
  } else {
    const s = String(v).trim()
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
    d = m ? new Date(Date.UTC(+m[3],+m[2]-1,+m[1])) : new Date(s)
  }
  if (!d || isNaN(d)) return ''
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
}

const tdToHour = (v) => {
  if (!v && v !== 0) return -1
  if (typeof v === 'number') return Math.floor(v * 24)
  const m = String(v).match(/(\d+):(\d+)/)
  return m ? parseInt(m[1]) : -1
}

// Timestamp local dd/mm/yyyy HH:MM:SS
const nowBRTimestamp = () => {
  const d = new Date(), p = (n) => String(n).padStart(2,'0')
  return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`
}

// buildFilter — usa datas reais da base, não "hoje" do sistema
const buildFilter = (allDates, período) => {
  if (período === 'TODOS' || !allDates.length) return () => true
  const sorted  = [...allDates].sort()
  const maxDate = sorted[sorted.length - 1]

  // Hoje no UTC do sistema
  const now = new Date()
  const todayStr = `${now.getUTCFullYear()}-${String(now.getUTCMonth()+1).padStart(2,'0')}-${String(now.getUTCDate()).padStart(2,'0')}`

  if (período === 'HOJE')   return (ds) => ds === todayStr
  if (período === 'ONTEM')  return (ds) => ds === maxDate  // mais recente da base
  if (período === 'SEMANA') {
    const c = new Date(maxDate + 'T00:00:00Z')
    c.setUTCDate(c.getUTCDate() - 6)
    const cut = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
    return (ds) => ds >= cut && ds <= maxDate
  }
  if (período === 'MÊS') {
    const mesAno = maxDate.slice(0,7)
    return (ds) => ds.slice(0,7) === mesAno
  }
  return () => true
}

// ─── Supabase ─────────────────────────────────────────────
const SB = {
  headers: {
    'apikey':        SUPABASE_ANON_KEY,
    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
    'Content-Type':  'application/json',
  },

  // Extrai datas únicas de um array de linhas
  extractDates(rows, colData) {
    const dates = new Set()
    rows.forEach(r => {
      const ds = serialToDateStr(r[colData])
      if (ds) dates.add(ds)
    })
    return [...dates].sort()
  },

  // ── UPSERT POR DATAS
  // 1. Identifica quais datas estão no novo upload
  // 2. Apaga do Supabase apenas os chunks que cobrem ESSAS datas
  // 3. Insere os novos chunks com essas datas marcadas
  async upsert(rows, verifTs, colData) {
    const snapDates = this.extractDates(rows, colData)
    if (!snapDates.length) throw new Error('Nenhuma data identificada no arquivo')

    // Busca IDs dos chunks que cobrem qualquer uma dessas datas
    // snapshot_dates é um array, usamos o operador de sobreposição @>
    // Para cada data, apaga chunks que a contém
    for (const date of snapDates) {
      const delRes = await fetch(
        `${SUPABASE_URL}/rest/v1/hospital_dados?snapshot_dates=cs.{"${date}"}`,
        { method: 'DELETE', headers: { ...this.headers, 'Prefer': 'return=minimal' } }
      )
      // Ignora erros de "nenhum registro" (204 é ok)
    }

    // Insere novos chunks (1000 linhas cada)
    const CHUNK = 1000
    for (let i = 0; i < rows.length; i += CHUNK) {
      const body = {
        verif_ts:       verifTs,
        snapshot_dates: snapDates,   // todas as datas deste arquivo
        chunk_idx:      Math.floor(i / CHUNK),
        dados:          rows.slice(i, i + CHUNK),
      }
      const res = await fetch(`${SUPABASE_URL}/rest/v1/hospital_dados`, {
        method: 'POST',
        headers: { ...this.headers, 'Prefer': 'return=minimal' },
        body: JSON.stringify(body),
      })
      if (!res.ok) {
        const txt = await res.text()
        throw new Error(`Supabase POST ${res.status}: ${txt}`)
      }
    }

    return { snapDates, chunks: Math.ceil(rows.length / CHUNK) }
  },

  // ── LOAD: busca TODOS os chunks de TODAS as datas, ordena e concatena
  // Deduplica por chave composta para evitar dados repetidos entre uploads sobrepostos
  async loadAll() {
    // Busca todos os registros ordenados por uploaded_at + chunk_idx
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/hospital_dados?select=verif_ts,uploaded_at,snapshot_dates,chunk_idx,dados&order=uploaded_at.asc,chunk_idx.asc`,
      { headers: { ...this.headers, 'Prefer': 'return=representation' } }
    )
    if (!res.ok) throw new Error(`Supabase load error: ${res.status}`)
    const chunks = await res.json()
    if (!chunks?.length) return null

    // Combina todos os dados de todos os chunks em ordem
    const allRows = chunks.flatMap(c => c.dados || [])

    // Pega o verif_ts mais recente (último upload)
    const lastVerifTs = chunks[chunks.length - 1]?.verif_ts || ''
    const lastUpAt    = chunks[chunks.length - 1]?.uploaded_at || ''

    return { dados: allRows, verif_ts: lastVerifTs, uploaded_at: lastUpAt }
  },
}

// ─── Sub-components ───────────────────────────────────────
function SH({ children, style = {} }) {
  return (
    <div style={{ fontSize:10, fontWeight:700, textTransform:'uppercase',
      letterSpacing:'.12em', color:T.muted, marginBottom:14, ...style }}>{children}</div>
  )
}
function Card({ children, style = {} }) {
  return (
    <div style={{ background:T.card, border:`1px solid ${T.border}`,
      borderRadius:16, padding:20, ...style }}>{children}</div>
  )
}
function StatCard({ icon, label, value, sub, accent, note }) {
  return (
    <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14,
      padding:'18px 20px', display:'flex', flexDirection:'column', gap:8,
      position:'relative', overflow:'hidden' }}>
      <div style={{ position:'absolute', top:-20, right:-20, width:80, height:80,
        borderRadius:'50%', background:accent, opacity:.07 }} />
      <div style={{ display:'flex', alignItems:'center', gap:9 }}>
        <span style={{ fontSize:18 }}>{icon}</span>
        <span style={{ fontSize:10.5, color:T.muted, textTransform:'uppercase',
          letterSpacing:'.09em', fontWeight:700 }}>{label}</span>
      </div>
      <div style={{ fontSize:34, fontWeight:900, color:T.text, lineHeight:1 }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:T.sub }}>{sub}</div>}
      {note && (
        <div style={{ fontSize:10, color:T.muted, fontStyle:'italic',
          borderTop:`1px solid ${T.border}`, paddingTop:6, marginTop:2 }}>{note}</div>
      )}
    </div>
  )
}

function StatusCards({ breakdown, total }) {
  const ORDER = ['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE','Falta Médica',
    'Remarcação Adm','Remarcação Médico','Remarcação médico','SEM_PONTO','OK']
  const sorted = [...breakdown].sort((a,b) => {
    const ia = ORDER.indexOf(a.key), ib = ORDER.indexOf(b.key)
    return (ia<0?99:ia) - (ib<0?99:ib)
  })
  return (
    <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12 }}>
      {sorted.map(s => {
        const cfg = getCfg(s.key)
        const pct = total > 0 ? ((s.count/total)*100) : 0
        return (
          <div key={s.key} style={{
            background:T.card2, border:`1px solid ${cfg.bg}44`,
            borderLeft:`4px solid ${cfg.bg}`, borderRadius:12, padding:'16px 18px',
            display:'flex', flexDirection:'column', gap:8,
          }}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between' }}>
              <div style={{ display:'flex', alignItems:'center', gap:7 }}>
                <div style={{ width:9, height:9, borderRadius:3, background:cfg.bg, flexShrink:0 }} />
                <span style={{ fontSize:10, fontWeight:800, color:cfg.bg,
                  textTransform:'uppercase', letterSpacing:'.08em' }}>{cfg.label}</span>
              </div>
              <span style={{ fontSize:10, fontWeight:700, color:T.muted,
                background:T.border, borderRadius:99, padding:'2px 8px' }}>
                {pct.toFixed(1)}%
              </span>
            </div>
            <div style={{ fontSize:32, fontWeight:900, color:cfg.bg, lineHeight:1 }}>
              {s.count.toLocaleString('pt-BR')}
            </div>
            <div style={{ fontSize:10.5, color:T.sub }}>{cfg.desc}</div>
          </div>
        )
      })}
    </div>
  )
}

function MedTable({ rows, tipo }) {
  if (!rows.length) return (
    <div style={{ color:T.muted, fontSize:13, padding:'12px 0' }}>
      Nenhum médico {tipo === 'atraso' ? 'em atraso' : 'sem ponto'}.
    </div>
  )
  return (
    <div style={{ overflowY:'auto', maxHeight:360 }}>
      <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
        <thead style={{ position:'sticky', top:0, background:T.card }}>
          <tr>
            {['#','Médico','Unidade','Status','Agendas','Pacientes'].map(h => (
              <th key={h} style={{ padding:'8px 10px', textAlign:'left',
                borderBottom:`1px solid ${T.border}`, color:T.muted, fontWeight:700,
                fontSize:10, textTransform:'uppercase', letterSpacing:'.07em',
                whiteSpace:'nowrap' }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.slice(0,40).map((m,i) => {
            const cfg = getCfg(m.status)
            return (
              <tr key={`${m.nome}${i}`}
                style={{ borderBottom:`1px solid ${T.border}`, transition:'background .1s' }}
                onMouseEnter={e => e.currentTarget.style.background='#0e1b2c'}
                onMouseLeave={e => e.currentTarget.style.background='transparent'}>
                <td style={{ padding:'8px 10px', color:T.muted, fontSize:10 }}>{i+1}</td>
                <td style={{ padding:'8px 10px', fontWeight:600, color:T.text,
                  maxWidth:180, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{m.nome}</td>
                <td style={{ padding:'8px 10px', color:T.sub, fontSize:11,
                  maxWidth:150, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{m.unid}</td>
                <td style={{ padding:'8px 10px' }}>
                  <span style={{ fontSize:10, fontWeight:700, padding:'2px 9px',
                    borderRadius:99, background:cfg.bg+'22', color:cfg.bg, whiteSpace:'nowrap' }}>
                    {cfg.label}
                  </span>
                </td>
                <td style={{ padding:'8px 10px', color:T.accent, fontWeight:700,
                  textAlign:'center', fontSize:13 }}>{m.agendas}</td>
                <td style={{ padding:'8px 10px', color:T.warning, fontWeight:700,
                  textAlign:'center', fontSize:13 }}>{m.pacts}</td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

const PERIODOS = [
  { key:'TODOS',  label:'Todos'  },
  { key:'HOJE',   label:'Hoje'   },
  { key:'ONTEM',  label:'Ontem'  },
  { key:'SEMANA', label:'Semana' },
  { key:'MÊS',    label:'Mês'    },
]
function PeriodoSelector({ value, onChange, infoLabel }) {
  return (
    <div style={{ display:'flex', alignItems:'center', gap:8 }}>
      <div style={{ display:'flex', gap:3, background:T.card,
        border:`1px solid ${T.border}`, borderRadius:10, padding:4 }}>
        {PERIODOS.map(p => (
          <button key={p.key} onClick={() => onChange(p.key)} style={{
            padding:'6px 13px', borderRadius:7, border:'none',
            fontSize:12, fontWeight:700, cursor:'pointer', transition:'all .15s',
            background: value===p.key
              ? `linear-gradient(135deg,${T.accent},${T.accentB})` : 'transparent',
            color: value===p.key ? '#000' : T.muted,
          }}>{p.label}</button>
        ))}
      </div>
      {infoLabel && <span style={{ fontSize:11, color:T.muted, whiteSpace:'nowrap' }}>{infoLabel}</span>}
    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
export default function Home() {
  const [dados,      setDados]      = useState([])
  const [verifTs,    setVerifTs]    = useState('')
  const [período,    setPeriodo]    = useState('TODOS')
  const [horaFilt,   setHoraFilt]   = useState('TODAS')
  const [uf,         setUf]         = useState('TODOS')
  const [statusFilt, setStatusFilt] = useState('TODOS')
  const [search,     setSearch]     = useState('')
  const [loading,    setLoading]    = useState(false)
  const [syncing,    setSyncing]    = useState(false)
  const [sbStatus,   setSbStatus]   = useState('idle') // idle|ok|error
  const [sbInfo,     setSbInfo]     = useState('')      // info text

  // ── Carrega do Supabase na montagem
  useEffect(() => {
    const load = async () => {
      setSyncing(true)
      try {
        const rec = await SB.loadAll()
        if (rec?.dados?.length) {
          setDados(rec.dados)
          setVerifTs(rec.verif_ts || '')
          setSbStatus('ok')
          setSbInfo(`${rec.dados.length.toLocaleString('pt-BR')} linhas`)
        }
      } catch(e) {
        console.error('Supabase load:', e)
        setSbStatus('error')
        setSbInfo(String(e.message).slice(0,60))
      } finally { setSyncing(false) }
    }
    load()
  }, [])

  // ── Detecção de colunas
  const cols = useMemo(() => {
    if (!dados.length) return {}
    const k = Object.keys(dados[0])
    const find = (...terms) =>
      k.find(c => terms.some(t => c.trim() === t)) ||
      k.find(c => terms.some(t => c.toLowerCase().includes(t.toLowerCase()))) || ''
    return {
      unidade:   find('NM_LOCAL','UNIDADE'),
      medico:    find('NM_MEDICO'),
      esp:       find('DS_ESPECIALIDADE','ESPECIALIDADE'),
      status:    find('STATUS'),
      espera:    find('TEMPO_DE_ESPERA'),
      qtPacts:   find('QT_PACIENTES_AGUARDANDO',' QT_PACIENTES_AGUARDANDO'),
      uf:        find('UF'),
      data:      find('DATA_AGENDA','DATA'),
      hrInicio:  find('HR_INICIO'),
      hrEntrada: find('HR_ENTRADA'),
    }
  }, [dados])

  // ── Upload
  const handleUpload = useCallback(async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setLoading(true)
    const ts = nowBRTimestamp()   // timestamp do momento do clique
    setVerifTs(ts)

    try {
      const buf  = await file.arrayBuffer()
      const wb   = XLSX.read(buf, { type:'buffer' })
      const ws   = wb.Sheets[wb.SheetNames[0]]
      const json = XLSX.utils.sheet_to_json(ws, { range:3, defval:'' })

      // Identifica a coluna de data antes do upsert
      const colData = (() => {
        if (!json.length) return ''
        const k = Object.keys(json[0])
        return k.find(c => c.trim() === 'DATA_AGENDA') ||
               k.find(c => c.toLowerCase().includes('data')) || ''
      })()

      // Upsert no Supabase: apaga só as datas do novo arquivo, insere novo
      setSyncing(true)
      setSbInfo('Atualizando Supabase…')
      try {
        const { snapDates, chunks } = await SB.upsert(json, ts, colData)

        // Recarrega tudo do Supabase para ter a visão completa acumulada
        const rec = await SB.loadAll()
        if (rec?.dados?.length) {
          setDados(rec.dados)
          setSbStatus('ok')
          setSbInfo(`${rec.dados.length.toLocaleString('pt-BR')} linhas · datas: ${snapDates.map(d=>d.split('-').reverse().join('/')).join(', ')}`)
        }
      } catch(e) {
        console.error('Supabase upsert:', e)
        // Mesmo com erro no Supabase, mostra os dados locais
        setDados(json)
        setSbStatus('error')
        setSbInfo(`Erro: ${String(e.message).slice(0,50)}`)
      } finally { setSyncing(false) }

      setPeriodo('TODOS'); setHoraFilt('TODAS')
      setUf('TODOS'); setStatusFilt('TODOS'); setSearch('')

    } finally { setLoading(false) }
  }, [])

  // ── Enriquece linhas
  const dadosRich = useMemo(() => dados.map(d => {
    const hrEnt    = d[cols.hrEntrada]
    const semPonto = !hrEnt || String(hrEnt).trim() === '' || String(hrEnt).trim() === 'NaT'
    return {
      ...d,
      _dateStr:    serialToDateStr(d[cols.data]),
      _hora:       tdToHour(d[cols.hrInicio]),
      _semPonto:   semPonto,
      _statusNorm: semPonto ? 'SEM_PONTO' : String(d[cols.status]||'').trim(),
    }
  }), [dados, cols])

  const allDates = useMemo(() =>
    [...new Set(dadosRich.map(d => d._dateStr).filter(Boolean))].sort(), [dadosRich])

  const horasDisp = useMemo(() =>
    [...new Set(dadosRich.filter(d => d._hora >= 0).map(d => d._hora))].sort((a,b)=>a-b),
    [dadosRich])

  const periodoFn = useMemo(() => buildFilter(allDates, período), [allDates, período])

  const períodoLabel = useMemo(() => {
    if (!allDates.length) return ''
    const sorted  = [...allDates].sort()
    const max = sorted[sorted.length-1], min = sorted[0]
    const fmt = s => s.split('-').reverse().join('/')
    const now = new Date()
    const todayStr = `${now.getUTCFullYear()}-${String(now.getUTCMonth()+1).padStart(2,'0')}-${String(now.getUTCDate()).padStart(2,'0')}`

    if (período === 'TODOS')  return min===max ? fmt(max) : `${fmt(min)} → ${fmt(max)}`
    if (período === 'HOJE')   return `Hoje · ${fmt(todayStr)}`
    if (período === 'ONTEM')  return `Ontem · ${fmt(max)}`
    if (período === 'SEMANA') {
      const c = new Date(max+'T00:00:00Z'); c.setUTCDate(c.getUTCDate()-6)
      const cs = `${c.getUTCFullYear()}-${String(c.getUTCMonth()+1).padStart(2,'0')}-${String(c.getUTCDate()).padStart(2,'0')}`
      const sm = sorted.find(d=>d>=cs)||cs
      return `${fmt(sm)} → ${fmt(max)}`
    }
    if (período === 'MÊS') {
      const ma = max.slice(0,7), sm = sorted.find(d=>d.slice(0,7)===ma)||max
      const [y,m] = max.split('-')
      const nome = new Date(+y,+m-1).toLocaleString('pt-BR',{month:'long'})
      return `${nome.charAt(0).toUpperCase()+nome.slice(1)} · ${fmt(sm)} → ${fmt(max)}`
    }
    return ''
  }, [allDates, período])

  const filtered = useMemo(() => {
    let r = dadosRich.filter(d => periodoFn(d._dateStr))
    if (horaFilt !== 'TODAS')   r = r.filter(d => d._hora === Number(horaFilt))
    if (uf !== 'TODOS')         r = r.filter(d => String(d[cols.uf]||'').trim() === uf)
    if (statusFilt !== 'TODOS') r = r.filter(d => d._statusNorm === statusFilt)
    if (search) r = r.filter(d =>
      [d[cols.unidade],d[cols.medico],d[cols.esp]]
        .some(v => String(v||'').toLowerCase().includes(search.toLowerCase())))
    return r
  }, [dadosRich, periodoFn, horaFilt, uf, statusFilt, search, cols])

  const ufs = useMemo(() =>
    [...new Set(filtered.map(d=>String(d[cols.uf]||'').trim()).filter(Boolean))].sort(),
    [filtered,cols])
  const statuses = useMemo(() =>
    [...new Set(dadosRich.filter(d=>periodoFn(d._dateStr)).map(d=>d._statusNorm).filter(Boolean))].sort(),
    [dadosRich,periodoFn])

  // KPIs
  const totalReg   = filtered.length
  const totalUnid  = new Set(filtered.map(d=>d[cols.unidade]).filter(Boolean)).size
  const totalMed   = new Set(filtered.map(d=>String(d[cols.medico]||'').trim()).filter(Boolean)).size
  const emAtraso   = filtered.filter(d=>['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm)).length
  const semPontoN  = filtered.filter(d=>d._semPonto).length
  const taxaAtras  = totalReg > 0 ? ((emAtraso/totalReg)*100).toFixed(1) : 0
  const totalPacs  = filtered.reduce((a,d)=>a+(Number(d[cols.qtPacts])||0),0)
  const mediaEsp   = totalReg > 0 ? filtered.reduce((a,d)=>a+parseHM(d[cols.espera]),0)/totalReg : 0

  const statusBreak = useMemo(() => {
    const m = {}
    filtered.forEach(d => { const s=d._statusNorm||'OK'; m[s]=(m[s]||0)+1 })
    return Object.entries(m).map(([key,count]) => ({ key, count }))
  }, [filtered])

  const medicosAtraso = useMemo(() => {
    const m = {}
    filtered.filter(d=>['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm))
      .forEach(d => {
        const nome=String(d[cols.medico]||'').trim(), unid=String(d[cols.unidade]||'').trim()
        const status=d._statusNorm
        if (!nome) return
        if (!m[nome]) m[nome]={nome,unid,status,pacts:0,agendasSet:new Set()}
        m[nome].agendasSet.add(`${d._dateStr}||${unid}`)
        m[nome].pacts+=Number(d[cols.qtPacts])||0
        const ord=['ATRASO GRAVE','ATRASO CRÍTICO','ATRASO']
        if(ord.indexOf(status)<ord.indexOf(m[nome].status)) m[nome].status=status
      })
    return Object.values(m).map(x=>({...x,agendas:x.agendasSet.size}))
      .sort((a,b)=>{const o=['ATRASO GRAVE','ATRASO CRÍTICO','ATRASO'];return o.indexOf(a.status)-o.indexOf(b.status)||b.pacts-a.pacts})
  }, [filtered,cols])

  const medicosSemPonto = useMemo(() => {
    const m = {}
    filtered.filter(d=>d._semPonto).forEach(d => {
      const nome=String(d[cols.medico]||'').trim(), unid=String(d[cols.unidade]||'').trim()
      if (!nome) return
      if (!m[nome]) m[nome]={nome,unid,status:d._statusNorm,pacts:0,agendasSet:new Set()}
      m[nome].agendasSet.add(`${d._dateStr}||${unid}`)
      m[nome].pacts+=Number(d[cols.qtPacts])||0
    })
    return Object.values(m).map(x=>({...x,agendas:x.agendasSet.size})).sort((a,b)=>b.pacts-a.pacts)
  }, [filtered,cols])

  const topUnidades = useMemo(() => {
    const m = {}
    filtered.filter(d=>['ATRASO','ATRASO CRÍTICO','ATRASO GRAVE'].includes(d._statusNorm))
      .forEach(d=>{const u=d[cols.unidade]||'Sem Unidade';m[u]=(m[u]||0)+1})
    return Object.entries(m).map(([nome,total])=>({nome,total})).sort((a,b)=>b.total-a.total).slice(0,10)
  }, [filtered,cols])

  const hasData = dados.length > 0

  return (
    <div style={{ background:T.bg, minHeight:'100vh',
      fontFamily:"'DM Sans','Segoe UI',sans-serif", color:T.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;background:${T.bg}}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
        select option{background:${T.surface};color:${T.text}}
        .ubtn:hover{opacity:.9;transform:translateY(-1px)}
        input::placeholder{color:${T.muted}}
      `}</style>

      {/* ── Topbar ── */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`,
        padding:'0 36px', height:62, display:'flex', alignItems:'center',
        justifyContent:'space-between', position:'sticky', top:0, zIndex:100 }}>
        <div style={{ display:'flex', alignItems:'center', gap:12 }}>
          <div style={{ width:32, height:32, borderRadius:8, fontSize:16,
            background:`linear-gradient(135deg,${T.accent},${T.accentB})`,
            display:'flex', alignItems:'center', justifyContent:'center' }}>🏥</div>
          <div>
            <div style={{ fontSize:14, fontWeight:700 }}>Monitor Hospitalar</div>
            <div style={{ fontSize:11, color:T.muted, display:'flex', alignItems:'center', gap:8 }}>
              <span>Atrasos Médicos · Operacional</span>
              {verifTs && <span style={{ color:T.accent }}>· Atualizado: {verifTs}</span>}
              {syncing  && <span style={{ color:T.warning }}>· Sincronizando…</span>}
              {!syncing && sbStatus==='ok'    && <span style={{ color:T.success }}>· ☁ Supabase ✓{sbInfo ? ` (${sbInfo})` : ''}</span>}
              {!syncing && sbStatus==='error' && <span style={{ color:T.danger  }}>· ☁ Erro{sbInfo ? `: ${sbInfo}` : ''}</span>}
            </div>
          </div>
        </div>

        <div style={{ display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
          {hasData && (
            <PeriodoSelector value={período} onChange={setPeriodo}
              infoLabel={`${totalReg.toLocaleString('pt-BR')} reg. · ${períodoLabel}`} />
          )}
          <label className="ubtn" style={{
            background:`linear-gradient(135deg,${T.accent},${T.accentB})`,
            color:'#000', fontWeight:700, fontSize:13,
            padding:'8px 18px', borderRadius:9, cursor:'pointer', transition:'all .2s' }}>
            {loading ? 'Lendo…' : syncing ? 'Salvando…' : '+ Carregar Planilha'}
            <input type="file" accept=".xlsx,.xls" style={{ display:'none' }} onChange={handleUpload} />
          </label>
        </div>
      </div>

      <div style={{ padding:'24px 36px' }}>
        {!hasData && syncing && (
          <div style={{ minHeight:'calc(100vh - 130px)', display:'flex', flexDirection:'column',
            alignItems:'center', justifyContent:'center', gap:14 }}>
            <div style={{ fontSize:48 }}>⏳</div>
            <div style={{ fontSize:18, fontWeight:700, color:T.warning }}>Carregando dados do Supabase…</div>
          </div>
        )}
        {!hasData && !syncing && (
          <div style={{ minHeight:'calc(100vh - 130px)', display:'flex', flexDirection:'column',
            alignItems:'center', justifyContent:'center', gap:14 }}>
            <div style={{ fontSize:52 }}>📋</div>
            <div style={{ fontSize:20, fontWeight:700 }}>Nenhuma planilha carregada</div>
            <div style={{ color:T.muted, fontSize:13 }}>Clique em "+ Carregar Planilha" para começar</div>
          </div>
        )}

        {hasData && (<>
          {/* Filtros */}
          <div style={{ display:'flex', gap:10, marginBottom:20, flexWrap:'wrap', alignItems:'center' }}>
            <input value={search} onChange={e=>setSearch(e.target.value)}
              placeholder="Buscar unidade, médico, especialidade…"
              style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:9,
                color:T.text, fontSize:13, padding:'8px 13px', outline:'none', width:270 }} />

            <select value={horaFilt} onChange={e=>setHoraFilt(e.target.value)} style={{
              background:T.card, border:`1px solid ${T.accent}55`, borderRadius:9,
              color:T.accent, fontSize:13, fontWeight:700, padding:'8px 12px',
              outline:'none', cursor:'pointer' }}>
              <option value="TODAS">⏰ Todas as horas</option>
              {horasDisp.map(h=><option key={h} value={h}>{String(h).padStart(2,'0')}h</option>)}
            </select>

            <select value={uf} onChange={e=>setUf(e.target.value)} style={{
              background:T.card, border:`1px solid ${T.border}`, borderRadius:9,
              color:T.text, fontSize:13, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
              <option value="TODOS">Todos os Estados</option>
              {ufs.map(u=><option key={u}>{u}</option>)}
            </select>

            <select value={statusFilt} onChange={e=>setStatusFilt(e.target.value)} style={{
              background:T.card, border:`1px solid ${T.border}`, borderRadius:9,
              color:T.text, fontSize:13, padding:'8px 12px', outline:'none', cursor:'pointer' }}>
              <option value="TODOS">Todos os Status</option>
              {statuses.map(s=><option key={s} value={s}>{getCfg(s).label}</option>)}
            </select>

            {(uf!=='TODOS'||statusFilt!=='TODOS'||search||horaFilt!=='TODAS') && (
              <button onClick={()=>{setUf('TODOS');setStatusFilt('TODOS');setSearch('');setHoraFilt('TODAS')}}
                style={{ background:'transparent', border:`1px solid ${T.border}`, borderRadius:9,
                  color:T.muted, fontSize:13, padding:'8px 13px', cursor:'pointer' }}>✕ Limpar</button>
            )}
          </div>

          {/* KPIs */}
          <div style={{ display:'grid', gridTemplateColumns:'repeat(5,1fr)', gap:12, marginBottom:18 }}>
            <StatCard icon="🩺" label="Total Registros"
              value={totalReg.toLocaleString('pt-BR')}
              sub={`${totalUnid} unidades · ${totalMed} médicos`} accent={T.accent} />
            <StatCard icon="⚠️" label="Em Atraso"
              value={emAtraso.toLocaleString('pt-BR')}
              sub={`${taxaAtras}% do total`} accent={T.danger} />
            <StatCard icon="🚫" label="Sem Ponto"
              value={semPontoN.toLocaleString('pt-BR')}
              sub="HR_ENTRADA vazia" accent={T.warning} />
            <StatCard icon="⏱️" label="Espera Média"
              value={fmtH(mediaEsp)} sub="por atendimento" accent="#818CF8"
              note={mediaEsp === 0 ? 'Sem registros de espera neste filtro — HR_REGISTRO_ESPERA ainda não disponível' : undefined} />
            <StatCard icon="👥" label="Pacientes na Fila"
              value={totalPacs.toLocaleString('pt-BR')} sub="aguardando agora" accent={T.success}
              note={totalPacs === 0 ? 'Sem fila registrada neste filtro — QT_PACIENTES_AGUARDANDO vazio' : undefined} />
          </div>

          {/* Status */}
          <Card style={{ marginBottom:18 }}>
            <SH>📊 Distribuição de Status — {períodoLabel}{horaFilt!=='TODAS'?` · ${String(horaFilt).padStart(2,'0')}h`:''}</SH>
            <StatusCards breakdown={statusBreak} total={totalReg} />
            {/* Nota explicativa sobre status ausentes */}
            {statusBreak.length > 0 && (() => {
              const keys = statusBreak.map(s => s.key)
              const ausentes = ['ATRASO GRAVE','Falta Médica','Remarcação Adm','Remarcação Médico','SEM_PONTO']
                .filter(k => !keys.includes(k))
              if (!ausentes.length) return null
              return (
                <div style={{ marginTop:14, padding:'10px 14px',
                  background:T.border+'55', borderRadius:9,
                  fontSize:11, color:T.muted, display:'flex', alignItems:'center', gap:8 }}>
                  <span>ℹ️</span>
                  <span>
                    Status sem ocorrências neste filtro (normal): {' '}
                    <strong style={{ color:T.sub }}>
                      {ausentes.map(k => getCfg(k).label).join(', ')}
                    </strong>
                    . Os dados são reais — não há registros desses status para o período/hora selecionado.
                  </span>
                </div>
              )
            })()}
          </Card>

          {/* Tabelas */}
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16, marginBottom:16 }}>
            <Card>
              <SH>🚨 Médicos em Atraso — agendas e pacientes ({medicosAtraso.length})</SH>
              <MedTable rows={medicosAtraso} tipo="atraso" />
            </Card>
            <Card>
              <SH>🔴 Médicos Sem Ponto — impacto na fila ({medicosSemPonto.length})</SH>
              <MedTable rows={medicosSemPonto} tipo="semponto" />
            </Card>
          </div>

          {/* Top Unidades */}
          <Card style={{ marginBottom:18 }}>
            <SH>🏥 Unidades com Mais Atrasos</SH>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'0 48px' }}>
              {topUnidades.map((u,i) => (
                <div key={u.nome} style={{ marginBottom:13 }}>
                  <div style={{ display:'flex', justifyContent:'space-between', marginBottom:5, gap:8 }}>
                    <div style={{ display:'flex', alignItems:'center', gap:8, minWidth:0 }}>
                      <span style={{ fontSize:10, fontWeight:800, minWidth:22, flexShrink:0,
                        color:i<3?T.danger:T.muted }}>#{i+1}</span>
                      <span style={{ fontSize:12.5, color:T.text, overflow:'hidden',
                        textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{u.nome}</span>
                    </div>
                    <span style={{ fontSize:13, fontWeight:700, flexShrink:0,
                      color:i===0?T.danger:i===1?'#FF7A00':T.warning }}>{u.total} atrasos</span>
                  </div>
                  <div style={{ background:T.border, borderRadius:99, height:6, overflow:'hidden' }}>
                    <div style={{ height:'100%', borderRadius:99,
                      background:i===0?T.danger:i===1?'#FF7A00':T.warning,
                      width:`${(u.total/(topUnidades[0]?.total||1))*100}%`, transition:'width .6s' }} />
                  </div>
                </div>
              ))}
            </div>
          </Card>

          {/* Footer */}
          <div style={{ textAlign:'center', color:T.muted, fontSize:11, paddingBottom:20 }}>
            Monitor Hospitalar · {new Date().toLocaleDateString('pt-BR')}
            {período!=='TODOS' && ` · ${períodoLabel}`}
            {horaFilt!=='TODAS' && ` · Hora ${String(horaFilt).padStart(2,'0')}h`}
            {verifTs && ` · Último upload: ${verifTs}`}
          </div>
        </>)}
      </div>
    </div>
  )
}
