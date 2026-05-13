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
  if (v === '' || v === null || v === undefined) return 0
  if (typeof v === 'number') return v * 24 * 60
  const s = String(v).trim()
  if (s.includes(':')) {
    // formato "HH:MM:SS" ou "H:MM"
    const parts = s.split(':').map(Number)
    return (parts[0]||0)*60 + (parts[1]||0)
  }
  // string numérica (fração de dia vinda como string)
  const n = Number(s)
  if (!isNaN(n)) return n * 24 * 60
  return 0
}
const fmtMin = (m) => {
  const mm=Math.round(Math.abs(m))
  if (mm < 60) return `${mm}min`
  return `${Math.floor(mm/60)}h${mm%60>0?` ${mm%60}m`:''}`
}
const serialToDateStr = (v) => {
  if (v === null || v === undefined || v === '') return ''
  // Número: serial Excel (dias desde 31/12/1899, com bug Lotus +1 para seriais > 60)
  if (typeof v === 'number') {
    const offsetDays = v > 60 ? 25568 : 25569
    const d = new Date(Math.round((v - offsetDays) * 86400 * 1000))
    if (!d || isNaN(d)) return ''
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
  }
  const s = String(v).trim()
  // String numérica ex: "46154" → tratar como serial
  if (/^\d{4,6}$/.test(s)) {
    const n = Number(s)
    const offsetDays = n > 60 ? 25568 : 25569
    const d = new Date(Math.round((n - offsetDays) * 86400 * 1000))
    if (!d || isNaN(d)) return ''
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
  }
  // String ISO "2026-05-13" ou "2026-05-13T00:00:00" → extrair direto sem conversão UTC
  const isoM = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (isoM) return `${isoM[1]}-${isoM[2]}-${isoM[3]}`
  // String "dd/mm/yyyy"
  const brM = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if (brM) return `${brM[3]}-${brM[2].padStart(2,'0')}-${brM[1].padStart(2,'0')}`
  // Fallback: tentar Date parse mas usar local date
  const d = new Date(s)
  if (!d || isNaN(d)) return ''
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
}
const todayStr = () => {
  const n = new Date()
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}`
}
// Converte DATA_AGENDA (qualquer formato) para serial Excel numérico
const toSerial = (v) => {
  if (v === null || v === undefined || v === '') return null
  if (typeof v === 'number' && v > 1000) return v  // já é serial Excel
  const s = String(v).trim()
  // string numérica "46154"
  if (/^\d{4,6}$/.test(s)) return Number(s)
  // ISO "2026-05-13..." → converter para serial
  const isoM = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (isoM) {
    const utcDays = Math.floor(Date.UTC(+isoM[1], +isoM[2]-1, +isoM[3]) / 86400000)
    return utcDays + 25568
  }
  // "dd/mm/yyyy"
  const brM = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if (brM) {
    const utcDays = Math.floor(Date.UTC(+brM[3], +brM[2]-1, +brM[1]) / 86400000)
    return utcDays + 25568
  }
  return null
}
// Serial de hoje baseado na data local do sistema
const todaySerial = () => {
  const n = new Date()
  return Math.floor(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / 86400000) + 25568
}
// Serial de uma data relativa (ontem, X dias atrás)
const offsetSerial = (serial, days) => serial + days
const buildFilter = (allSerials, período, dateFrom, dateTo) => {
  if (!allSerials.length) return () => true
  const sorted    = [...allSerials].sort((a,b)=>a-b)
  const maxSerial = sorted[sorted.length - 1]
  if (período === 'HOJE') {
    const s = todaySerial()
    return (serial) => serial === s
  }
  if (período === 'ONTEM') {
    const s = todaySerial() - 1
    return (serial) => serial === s
  }
  if (período === 'SEMANA') {
    const cut = maxSerial - 6
    return (serial) => serial >= cut && serial <= maxSerial
  }
  if (período === 'MES') {
    // mesmo mês do maxSerial
    const maxDate = new Date((maxSerial - 25568) * 86400000)
    const y = maxDate.getUTCFullYear(), m = maxDate.getUTCMonth()
    const firstOfMonth = Math.floor(Date.UTC(y,m,1)/86400000) + 25568
    const firstOfNext  = Math.floor(Date.UTC(y,m+1,1)/86400000) + 25568
    return (serial) => serial >= firstOfMonth && serial < firstOfNext
  }
  if (período === 'ANO') {
    const maxDate = new Date((maxSerial - 25568) * 86400000)
    const y = maxDate.getUTCFullYear()
    const firstOfYear = Math.floor(Date.UTC(y,0,1)/86400000) + 25568
    const firstOfNext = Math.floor(Date.UTC(y+1,0,1)/86400000) + 25568
    return (serial) => serial >= firstOfYear && serial < firstOfNext
  }
  if (período === 'PERIODO') {
    const from = dateFrom ? (Math.floor(Date.UTC(...dateFrom.split('-').map((v,i)=>i===1?+v-1:+v))/86400000)+25568) : sorted[0]
    const to   = dateTo   ? (Math.floor(Date.UTC(...dateTo.split('-').map((v,i)=>i===1?+v-1:+v))/86400000)+25568) : maxSerial
    return (serial) => serial >= from && serial <= to
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

// ─── ALTERADO: Unidades com Maior Espera + Pacientes Aguardando ──────────────
function MaiorEsperaCard({ rows, cols, horaFilt }) {
  const items = useMemo(() => {
    // ── Lógica correta de snapshot ────────────────────────────────────────
    // Cada linha tem _hrReg (hora de HR_REGISTRO_ESPERA) e _dateStr (data da agenda).
    //
    // Estratégia:
    // 1. Agrupar por DATA + UNIDADE
    // 2. Por unidade+data: achar o snapshot correto (hrAlvo exato se filtrado, ou max disponível)
    // 3. Pegar só as linhas desse snapshot, com pac > 0
    // 4. Agregar esperaMax e pacAguardando
    // 5. Ordenar por esperaMax desc

    const hrAlvo = horaFilt !== 'TODAS' ? Number(horaFilt) : null

    // 1ª passagem: achar o snapshot alvo por (data+unidade)
    // chave: `${dateStr}||${unid}`
    const keyAlvo = {}
    rows.forEach(d => {
      const unid  = String(d[cols.unidade]||'').trim() || 'Sem Unidade'
      const hrReg = d._hrReg
      const ds    = d._dateStr || ''
      if (hrReg === null || hrReg === undefined) return
      const key = ds + '||' + unid
      if (hrAlvo !== null) {
        // filtro ativo: exatamente aquela hora
        if (hrReg === hrAlvo) keyAlvo[key] = hrAlvo
      } else {
        // sem filtro: máximo disponível por data+unidade
        if (keyAlvo[key] === undefined || hrReg > keyAlvo[key]) {
          keyAlvo[key] = hrReg
        }
      }
    })

    // 2ª passagem: agregar apenas registros do snapshot correto, com pac > 0
    // resultado agrupado por UNIDADE (não por data+unidade) para mostrar ranking geral
    const m = {}
    rows.forEach(d => {
      const unid  = String(d[cols.unidade]||'').trim() || 'Sem Unidade'
      const hrReg = d._hrReg
      const ds    = d._dateStr || ''
      const pac   = Number(d[cols.qtPacts]) || 0
      const esp   = parseEsperaMin(d[cols.espera])
      if (hrReg === null || hrReg === undefined) return
      const key = ds + '||' + unid
      if (keyAlvo[key] === undefined) return
      if (hrReg !== keyAlvo[key]) return
      if (pac <= 0) return   // só conta se tem pacientes aguardando

      if (!m[unid]) m[unid] = { unid, esperaMax: 0, pacAguardando: 0, registros: 0, hrSnap: hrReg }
      if (esp > m[unid].esperaMax) m[unid].esperaMax = esp
      m[unid].pacAguardando += pac
      m[unid].registros     += 1
    })

    return Object.values(m)
      .filter(x => x.esperaMax > 0 && x.pacAguardando > 0)
      .sort((a, b) => b.esperaMax - a.esperaMax)
      .slice(0, 10)
  }, [rows, cols, horaFilt])

  // Debug: logar primeiras linhas para diagnóstico
  if (rows.length > 0) {
    const sample = rows[0]
    console.log('[DEBUG row0]', {
      DATA_AGENDA: sample[cols.data],
      _dateSerial: sample._dateSerial,
      _dateStr: sample._dateStr,
      _hrReg: sample._hrReg,
      HR_REGISTRO_ESPERA: sample[cols.hrRegistroEspera],
      TEMPO_DE_ESPERA: sample[cols.espera],
      QT_PACTS: sample[cols.qtPacts],
      totalRows: rows.length,
    })
  }

  if (!items.length) return (
    <div style={{ color:T.muted, fontSize:13 }}>Sem dados de espera nos filtros atuais.</div>
  )

  const maxMin = items[0]?.esperaMax || 1
  const maxPac = Math.max(...items.map(x => x.pacAguardando), 1)

  return (
    <div>
      {items.map((r, i) => {
        const barEspPct = (r.esperaMax / maxMin) * 100
        const barPacPct = (r.pacAguardando / maxPac) * 100
        const espColor  = i < 3 ? T.danger : T.warning
        const pacColor  = i < 3 ? '#FF7A50' : T.accent

        return (
          <div key={r.unid} style={{ marginBottom: 18 }}>
            <div style={{ display:'flex', justifyContent:'space-between',
              alignItems:'flex-start', marginBottom: 6, gap: 8 }}>
              <div style={{ display:'flex', alignItems:'flex-start', gap: 8, minWidth: 0 }}>
                <span style={{ fontSize:11, fontWeight:700,
                  color: i < 3 ? T.danger : T.muted,
                  minWidth:22, flexShrink:0, marginTop:1 }}>#{i+1}</span>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize:13, color:T.text, fontWeight:500,
                    overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                    {r.unid}
                  </div>
                  <div style={{ display:'flex', alignItems:'center', gap: 10, marginTop: 3 }}>
                    <span style={{ fontSize:10, color:T.muted }}>
                      snapshot {String(r.hrSnap).padStart(2,'0')}h · {r.registros} {r.registros === 1 ? 'registro' : 'registros'}
                    </span>
                    <span style={{ fontSize:10, color: pacColor, fontWeight:700 }}>
                      {r.pacAguardando.toLocaleString('pt-BR')} pac. aguardando
                    </span>
                  </div>
                </div>
              </div>
              <span style={{ fontSize:13, fontWeight:700, color: espColor, flexShrink:0 }}>
                {fmtMin(r.esperaMax)}
              </span>
            </div>

            <div style={{ marginBottom: 4 }}>
              <div style={{ background:T.border, borderRadius:6, height:6, overflow:'hidden' }}>
                <div style={{ height:'100%', borderRadius:6, background: espColor,
                  width:`${barEspPct}%`, transition:'width .6s ease' }} />
              </div>
            </div>

            <div style={{ display:'flex', alignItems:'center', gap: 6 }}>
              <span style={{ fontSize:9, color: T.muted, minWidth: 80 }}>Pac. aguardando</span>
              <div style={{ flex:1, background:T.border, borderRadius:4, height:4, overflow:'hidden' }}>
                <div style={{ height:'100%', borderRadius:4, background: pacColor,
                  width:`${barPacPct}%`, transition:'width .6s ease', opacity:.8 }} />
              </div>
              <span style={{ fontSize:10, color: pacColor, fontWeight:600, minWidth:28, textAlign:'right' }}>
                {r.pacAguardando}
              </span>
            </div>
          </div>
        )
      })}
    </div>
  )
}

function RecorrenciaAtrasoCard({ rows, cols }) {
  const items = useMemo(() => {
    const m = {}
    rows.forEach(d => {
      const unid   = String(d[cols.unidade]||'').trim() || 'Sem Unidade'
      const status = String(d[cols.status]||'').trim()
      const dateS  = d._dateStr || ''
      const isAtraso = status.toUpperCase().includes('ATRASO')
      if (!m[unid]) m[unid] = {
        unid,
        totalAtrasos: 0,
        totalRegistros: 0,
        diasComAtraso: new Set(),
        atrasoGrave: 0,
        atrasoCritico: 0,
        atrasoNormal: 0,
      }
      m[unid].totalRegistros += 1
      if (isAtraso) {
        m[unid].totalAtrasos += 1
        if (dateS) m[unid].diasComAtraso.add(dateS)
        const su = status.toUpperCase()
        if (su.includes('GRAVE'))    m[unid].atrasoGrave    += 1
        else if (su.includes('CRÍT') || su.includes('CRIT')) m[unid].atrasoCritico += 1
        else                          m[unid].atrasoNormal   += 1
      }
    })

    return Object.values(m)
      .filter(x => x.totalAtrasos > 0)
      .map(x => ({
        ...x,
        diasComAtraso: x.diasComAtraso.size,
        taxaAtraso: x.totalRegistros > 0 ? (x.totalAtrasos / x.totalRegistros) * 100 : 0,
      }))
      // Ordenar por dias com atraso desc, depois por total
      .sort((a, b) => b.diasComAtraso - a.diasComAtraso || b.totalAtrasos - a.totalAtrasos)
      .slice(0, 10)
  }, [rows, cols])

  if (!items.length) return (
    <div style={{ color:T.muted, fontSize:13 }}>Nenhum atraso registrado no período.</div>
  )

  const maxDias    = items[0]?.diasComAtraso || 1
  const maxAtrasos = Math.max(...items.map(x => x.totalAtrasos), 1)

  return (
    <div style={{ overflowX:'auto' }}>
      <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12.5 }}>
        <thead>
          <tr style={{ borderBottom:`1px solid ${T.border}` }}>
            {['#','Unidade','Dias c/ Atraso','Total Atrasos','Taxa','Composição'].map(h => (
              <th key={h} style={{ padding:'9px 12px', textAlign:'left', color:T.muted,
                fontWeight:600, fontSize:10.5, textTransform:'uppercase',
                letterSpacing:'.06em', whiteSpace:'nowrap' }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {items.map((r, i) => {
            const diasPct    = (r.diasComAtraso / maxDias) * 100
            const atrasosPct = (r.totalAtrasos / maxAtrasos) * 100
            const rowColor   = i < 3 ? T.danger : i < 6 ? T.warning : T.muted
            // Composição: proporção dos tipos
            const total = r.atrasoGrave + r.atrasoCritico + r.atrasoNormal || 1
            const graveW    = (r.atrasoGrave    / total) * 100
            const criticoW  = (r.atrasoCritico  / total) * 100
            const normalW   = (r.atrasoNormal   / total) * 100

            return (
              <tr key={r.unid} className="row-table"
                style={{ borderBottom:`1px solid ${T.border}22`, transition:'background .15s' }}>
                <td style={{ padding:'10px 12px', color: rowColor, fontWeight:700, fontSize:12 }}>
                  #{i+1}
                </td>
                <td style={{ padding:'10px 12px', minWidth:180, maxWidth:260 }}>
                  <div style={{ fontWeight:600, color:T.text, fontSize:12.5,
                    overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                    {r.unid}
                  </div>
                  <div style={{ fontSize:10, color:T.muted, marginTop:2 }}>
                    {r.totalRegistros} registros totais
                  </div>
                </td>
                {/* Dias com atraso + mini-barra */}
                <td style={{ padding:'10px 12px', minWidth:130 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                    <span style={{ fontWeight:800, color: rowColor, fontSize:14, minWidth:20 }}>
                      {r.diasComAtraso}
                    </span>
                    <div style={{ flex:1, background:T.border, borderRadius:4, height:5, overflow:'hidden', minWidth:60 }}>
                      <div style={{ height:'100%', borderRadius:4, background: rowColor,
                        width:`${diasPct}%`, transition:'width .6s ease' }} />
                    </div>
                  </div>
                  <div style={{ fontSize:10, color:T.muted, marginTop:2 }}>
                    {r.diasComAtraso === 1 ? 'dia único' : 'dias distintos'}
                  </div>
                </td>
                {/* Total atrasos */}
                <td style={{ padding:'10px 12px', minWidth:120 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                    <span style={{ fontWeight:700, color: rowColor, fontSize:13, minWidth:24 }}>
                      {r.totalAtrasos}
                    </span>
                    <div style={{ flex:1, background:T.border, borderRadius:4, height:5, overflow:'hidden', minWidth:50 }}>
                      <div style={{ height:'100%', borderRadius:4, background: rowColor, opacity:.7,
                        width:`${atrasosPct}%`, transition:'width .6s ease' }} />
                    </div>
                  </div>
                </td>
                {/* Taxa */}
                <td style={{ padding:'10px 12px' }}>
                  <span style={{
                    fontSize:12, fontWeight:700, padding:'3px 9px', borderRadius:20,
                    background: r.taxaAtraso >= 50 ? T.danger+'22'
                              : r.taxaAtraso >= 20 ? T.warning+'22' : T.muted+'22',
                    color: r.taxaAtraso >= 50 ? T.danger
                         : r.taxaAtraso >= 20 ? T.warning : T.muted,
                  }}>
                    {r.taxaAtraso.toFixed(1)}%
                  </span>
                </td>
                {/* Composição de severidade */}
                <td style={{ padding:'10px 12px', minWidth:140 }}>
                  <div style={{ display:'flex', height:8, borderRadius:4, overflow:'hidden', gap:1 }}>
                    {r.atrasoNormal > 0 && (
                      <div title={`Normal: ${r.atrasoNormal}`}
                        style={{ width:`${normalW}%`, background: '#F59E0B', borderRadius:'4px 0 0 4px' }} />
                    )}
                    {r.atrasoCritico > 0 && (
                      <div title={`Crítico: ${r.atrasoCritico}`}
                        style={{ width:`${criticoW}%`, background: '#EF4444' }} />
                    )}
                    {r.atrasoGrave > 0 && (
                      <div title={`Grave: ${r.atrasoGrave}`}
                        style={{ width:`${graveW}%`, background: '#7C3AED', borderRadius:'0 4px 4px 0' }} />
                    )}
                  </div>
                  <div style={{ display:'flex', gap:8, marginTop:4 }}>
                    {r.atrasoNormal   > 0 && <span style={{ fontSize:9, color:'#F59E0B' }}>■ {r.atrasoNormal} normal</span>}
                    {r.atrasoCritico  > 0 && <span style={{ fontSize:9, color:'#EF4444' }}>■ {r.atrasoCritico} crít.</span>}
                    {r.atrasoGrave    > 0 && <span style={{ fontSize:9, color:'#7C3AED' }}>■ {r.atrasoGrave} grave</span>}
                  </div>
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
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

const rowDateStr = (r) => {
  const v = r['DATA_AGENDA'] ?? r[Object.keys(r).find(k => k.includes('DATA')) || ''] ?? ''
  return serialToDateStr(v)
}
const rowSerial = (r) => {
  const v = r['DATA_AGENDA'] ?? r[Object.keys(r).find(k => k.includes('DATA')) || ''] ?? ''
  return toSerial(v)
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

  useEffect(() => {
    const loadFromSupabase = async () => {
      try {
        setStoreStatus('Carregando dados…')
        let allRows = []
        let lastVerifTs = ''
        const res = await sbFetch(
          `hospital_dados?select=dados,verif_ts&order=id.asc`,
          { headers: { 'Range': '0-9999', 'Range-Unit': 'items' } }
        )
        if (!res.ok) {
          const txt = await res.text()
          setStoreStatus(`Erro ${res.status}: ${txt.slice(0,120)}`)
          setInitLoading(false)
          return
        }
        const allItems = await res.json()
        if (Array.isArray(allItems)) {
          for (const item of allItems) {
            try {
              const d = item.dados
              if (!d) continue
              const parsed = typeof d === 'string' ? JSON.parse(d) : d
              if (Array.isArray(parsed))                     allRows = allRows.concat(parsed)
              else if (parsed && typeof parsed === 'object') allRows.push(parsed)
            } catch(e) {}
            if (item.verif_ts) lastVerifTs = item.verif_ts
          }
        }
        setStoreStatus(`Processando ${allRows.length} registros…`)
        if (allRows.length > 0) {
          if (lastVerifTs) setTimestamp(lastVerifTs)
          const diasSet = new Set(allRows.map(rowDateStr).filter(Boolean))
          setDados(allRows)
          setStorageInfo({ dias: diasSet.size, total: allRows.length })
          setPeriodo('MES')
          setStoreStatus(`☁ ${diasSet.size} dias · ${allRows.length.toLocaleString('pt-BR')} registros`)
          setTimeout(() => setStoreStatus(''), 4000)
        } else {
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

  const saveToSupabase = useCallback(async (newRows, ts) => {
    setStoring(true)
    try {
      const MAX_ROWS_PER_INSERT = 2000
      setStoreStatus('Limpando dados antigos…')
      await sbFetch('hospital_dados?id=gt.0', { method: 'DELETE' })
      const byDate = {}
      newRows.forEach(r => {
        const ds = rowDateStr(r)
        if (!ds) return
        if (!byDate[ds]) byDate[ds] = []
        byDate[ds].push(r)
      })
      const dates = Object.keys(byDate).sort()
      let savedDays = 0
      for (const date of dates) {
        savedDays++
        const dayRows = byDate[date]
        const chunks = Math.ceil(dayRows.length / MAX_ROWS_PER_INSERT)
        for (let ci = 0; ci < chunks; ci++) {
          const slice = dayRows.slice(ci * MAX_ROWS_PER_INSERT, (ci + 1) * MAX_ROWS_PER_INSERT)
          setStoreStatus(`Salvando ${date} parte ${ci+1}/${chunks} (dia ${savedDays}/${dates.length})…`)
          const res = await sbFetch('hospital_dados', {
            method: 'POST',
            headers: { 'Prefer': 'return=minimal' },
            body: JSON.stringify({ verif_ts: ts, dados: slice, snapshot_dates: [date] }),
          })
          if (!res.ok) {
            const txt = await res.text()
            setStoreStatus(`⚠ Erro ${date} parte ${ci+1}: ${txt.slice(0,120)}`)
            setStoring(false)
            return
          }
        }
      }
      const diasSet = new Set(dates)
      setStorageInfo({ dias: diasSet.size, total: newRows.length })
      setStoreStatus('✓ Salvo no Supabase')
      setTimeout(() => setStoreStatus(''), 3000)
    } catch (e) {
      console.error('save error:', e)
      setStoreStatus(`⚠ Erro: ${e.message}`)
    }
    setStoring(false)
  }, [])

  const clearStorage = async () => {
    if (!confirm('Apagar TODOS os dados do Supabase? Isso não pode ser desfeito.')) return
    setStoring(true)
    setStoreStatus('Apagando…')
    try { await sbFetch('hospital_dados?id=gt.0', { method: 'DELETE' }) } catch(e) {}
    setDados([]); setStorageInfo({ dias:0, total:0 }); setTimestamp(''); setStoreStatus(''); setStoring(false)
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
    setDados(json); setInitLoading(false); setLoading(false)
    saveToSupabase(json, ts)
  }

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

  const dadosComData = useMemo(() => {
    if (!dados.length) return []
    return dados.map(d => {
      // _hora: hora da consulta (HR_INICIO) — filtro geral de hora
      const hrV = d[cols.hrInicio]
      let hora = 99
      if (typeof hrV === 'number') hora = Math.floor(hrV * 24)
      else if (hrV && typeof hrV === 'string') {
        const m = hrV.match(/(\d{1,2}):/)
        if (m) hora = parseInt(m[1])
      }
      // _hrReg: hora exata do snapshot de espera (HR_REGISTRO_ESPERA)
      // Suporta: número fracionário (0.4583), string numérica ('0.4583'), string hora ('11:00:00')
      const hrRV = d[cols.hrRegistroEspera]
      let hrReg = null
      if (hrRV !== '' && hrRV !== null && hrRV !== undefined) {
        const n = Number(hrRV)
        if (!isNaN(n) && typeof hrRV !== 'string') {
          // número nativo (fração de dia)
          hrReg = Math.round(n * 24)
        } else if (typeof hrRV === 'string') {
          const mStr = hrRV.match(/^(\d{1,2}):/)
          if (mStr) {
            // string formato "HH:MM:SS"
            hrReg = parseInt(mStr[1])
          } else if (!isNaN(n) && n !== 0) {
            // string numérica "0.4583..."
            hrReg = Math.round(n * 24)
          }
        }
      }
      const _serial = toSerial(d[cols.data])
      return { ...d, _dateSerial: _serial, _dateStr: serialToDateStr(d[cols.data]), _hora: hora, _hrReg: hrReg }
    })
  }, [dados, cols])

  const allSerials = useMemo(() => {
    const serials = [...new Set(dadosComData.map(d=>d._dateSerial).filter(s=>s!=null))].sort((a,b)=>a-b)
    if (dadosComData.length > 0) {
      const s0 = dadosComData[0]
      const ts = todaySerial()
      console.log('[allSerials debug] todaySerial=' + ts + ' serialsArray=' + JSON.stringify(serials) + ' período=' + período + ' row0_DATA_AGENDA=' + s0[cols.data] + ' type=' + typeof s0[cols.data] + ' _dateSerial=' + s0._dateSerial)
    }
    return serials
  }, [dadosComData])

  // allDates: para exibição (períodoLabel) — converter seriais para strings
  const allDates = useMemo(() =>
    allSerials.map(s => serialToDateStr(s)).filter(Boolean).sort(),
    [allSerials])

  const periodoFn = useMemo(() => buildFilter(allSerials, período, dateFrom, dateTo), [allSerials, período, dateFrom, dateTo])

  const dadosPorPeriodo = useMemo(() =>
    dadosComData.filter(d => d._dateSerial != null && periodoFn(d._dateSerial)),
    [dadosComData, periodoFn])

  const horasDisp = useMemo(() => {
    // Horas dos snapshots de espera (HR_REGISTRO_ESPERA) disponíveis no período
    const hs = [...new Set(dadosPorPeriodo.map(d=>d._hrReg).filter(h=>h!==null&&h>=0&&h<=23))].sort((a,b)=>a-b)
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
    if (período==='PERIODO' && dateFrom && dateTo) return `${fmt(dateFrom)} → ${fmt(dateTo)}`
    return ''
  }, [allDates, período, dateFrom, dateTo])

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
            background: storing ? T.border : `linear-gradient(135deg,${T.accent},${T.accentB})`,
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

            {/* ── ALTERADO: Maior Espera + Pacientes Aguardando ── */}
            <Card style={{ marginBottom:16 }}>
              <SH>⏳ Unidades com Maior Tempo de Espera · Pacientes Aguardando</SH>
              <MaiorEsperaCard rows={dadosPorPeriodo} cols={cols} horaFilt={horaFilt} />
            </Card>

            {/* ── NOVO: Recorrência de Atraso ── */}
            <Card style={{ marginBottom:16 }}>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
                <SH style={{ marginBottom:0 }}>🔁 Unidades que Mais Atrasaram — Recorrência</SH>
                <div style={{ display:'flex', gap:12, fontSize:10 }}>
                  <span style={{ color:'#F59E0B' }}>■ Normal 31–45min</span>
                  <span style={{ color:'#EF4444' }}>■ Crítico 46min–1h30</span>
                  <span style={{ color:'#7C3AED' }}>■ Grave +1h30</span>
                </div>
              </div>
              <RecorrenciaAtrasoCard rows={filtered} cols={cols} />
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
