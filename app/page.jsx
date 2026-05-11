'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

// ─── Palette ──────────────────────────────────────────────
const T = {
  bg: '#080C14',
  surface: '#0E1623',
  card: '#111A2B',
  border: '#1C2D45',
  accent: '#00C6FF',
  accentB: '#0072FF',
  danger: '#FF4D6A',
  warning: '#FFB340',
  success: '#00E5A0',
  text: '#EDF2FF',
  muted: '#5A7799',
  sub: '#8FADC8',
}

// ─── Helpers ──────────────────────────────────────────────
const parseHM = (v) => {
  if (!v && v !== 0) return 0

  if (typeof v === 'number') return v * 24

  const s = String(v).trim()

  if (s.includes(':')) {
    const [h, m] = s.split(':').map(Number)
    return (h || 0) + (m || 0) / 60
  }

  const n = parseFloat(s.replace(',', '.'))

  return isNaN(n) ? 0 : n
}

const fmtH = (h) => {
  const abs = Math.abs(h)
  const hh = Math.floor(abs)
  const mm = Math.round((abs - hh) * 60)

  return `${hh}h${mm > 0 ? ` ${mm}m` : ''}`
}

const STATUS_COLOR = (s = '') => {
  const u = s.toUpperCase()

  if (u.includes('CRÍTICO')) return T.danger
  if (u.includes('GRAVE')) return '#FF7A00'
  if (u.includes('ATRASO')) return T.warning
  if (u.includes('FALTA')) return '#B060FF'
  if (u.includes('REMARCA')) return T.muted

  return T.success
}

const STATUS_LABEL = (s = '') => {
  const u = s.toUpperCase()

  if (u.includes('CRÍTICO')) return 'Crítico'
  if (u.includes('GRAVE')) return 'Grave'
  if (u.includes('ATRASO')) return 'Atraso'
  if (u.includes('FALTA')) return 'Falta'
  if (u.includes('REMARCA')) return 'Remarcação'

  return 'OK'
}

// ─── Parseia datas ────────────────────────────────────────
const parseDate = (v) => {

  if (!v) return null

  if (typeof v === 'number') {

    const d = new Date((v - 25569) * 86400 * 1000)

    return isNaN(d) ? null : d
  }

  const s = String(v).trim()

  if (!s) return null

  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)

  if (m1) {
    return new Date(+m1[3], +m1[2] - 1, +m1[1])
  }

  const d = new Date(s)

  return isNaN(d) ? null : d
}

// ─── Filtro período corrigido ─────────────────────────────
const dentroDoPeríodo = (date, período, baseDatas = []) => {

  if (!date || período === 'TODOS') return true

  const datasValidas = baseDatas.filter(Boolean)

  const dataReferencia = datasValidas.length
    ? new Date(Math.max(...datasValidas.map(d => d.getTime())))
    : new Date()

  if (período === 'DIA') {

    return (
      date.getDate() === dataReferencia.getDate() &&
      date.getMonth() === dataReferencia.getMonth() &&
      date.getFullYear() === dataReferencia.getFullYear()
    )

  }

  if (período === 'SEMANA') {

    const inicioSemana = new Date(dataReferencia)

    inicioSemana.setDate(dataReferencia.getDate() - 6)

    inicioSemana.setHours(0, 0, 0, 0)

    return date >= inicioSemana && date <= dataReferencia

  }

  if (período === 'MÊS') {

    return (
      date.getMonth() === dataReferencia.getMonth() &&
      date.getFullYear() === dataReferencia.getFullYear()
    )

  }

  return true
}

// ─── Horizontal Bar ───────────────────────────────────────
function HBar({ label, value, max, color, unit = '', rank }) {

  const pct = max > 0
    ? (value / max) * 100
    : 0

  return (
    <div style={{ marginBottom: 14 }}>

      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        marginBottom: 5
      }}>

        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: 8
        }}>

          {rank != null && (
            <span style={{
              fontSize: 11,
              fontWeight: 700,
              color: rank < 3 ? T.danger : T.muted
            }}>
              #{rank + 1}
            </span>
          )}

          <span style={{
            fontSize: 13,
            color: T.text,
            fontWeight: 500
          }}>
            {label}
          </span>

        </div>

        <span style={{
          fontSize: 13,
          fontWeight: 700,
          color
        }}>
          {value}{unit}
        </span>

      </div>

      <div style={{
        background: T.border,
        borderRadius: 6,
        height: 6,
        overflow: 'hidden'
      }}>

        <div style={{
          height: '100%',
          borderRadius: 6,
          background: color,
          width: `${pct}%`,
        }} />

      </div>

    </div>
  )
}

// ─── Card ─────────────────────────────────────────────────
function Card({ children, style = {} }) {

  return (
    <div style={{
      background: T.card,
      border: `1px solid ${T.border}`,
      borderRadius: 16,
      padding: 24,
      ...style
    }}>
      {children}
    </div>
  )
}

// ─── Section Header ───────────────────────────────────────
function SectionHeader({ children }) {

  return (
    <h2 style={{
      fontSize: 14,
      fontWeight: 700,
      textTransform: 'uppercase',
      letterSpacing: '.12em',
      color: T.sub,
      marginBottom: 18,
      marginTop: 0
    }}>
      {children}
    </h2>
  )
}

// ─── Period Selector ──────────────────────────────────────
const PERIODOS = [
  { key: 'TODOS', label: 'Todos' },
  { key: 'DIA', label: 'Dia' },
  { key: 'SEMANA', label: 'Semana' },
  { key: 'MÊS', label: 'Mês' },
]

function PeriodoSelector({ value, onChange }) {

  return (
    <div style={{
      display: 'flex',
      gap: 4,
      background: T.card,
      border: `1px solid ${T.border}`,
      borderRadius: 10,
      padding: 4
    }}>

      {PERIODOS.map(p => (

        <button
          key={p.key}
          onClick={() => onChange(p.key)}
          style={{
            padding: '6px 14px',
            borderRadius: 7,
            border: 'none',
            fontSize: 12,
            fontWeight: 700,
            cursor: 'pointer',
            background:
              value === p.key
                ? `linear-gradient(135deg,${T.accent},${T.accentB})`
                : 'transparent',
            color:
              value === p.key
                ? '#000'
                : T.muted,
          }}
        >
          {p.label}
        </button>

      ))}

    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
export default function Home() {

  const [dados, setDados] = useState([])
  const [uf, setUf] = useState('TODOS')
  const [status, setStatus] = useState('TODOS')
  const [search, setSearch] = useState('')
  const [loading, setLoading] = useState(false)
  const [período, setPeriodo] = useState('TODOS')

  // ─── Upload ─────────────────────────────────────────────
  const handleUpload = async (e) => {

    const file = e.target.files[0]

    if (!file) return

    setLoading(true)

    const buf = await file.arrayBuffer()

    const wb = XLSX.read(buf, {
      type: 'buffer'
    })

    const ws = wb.Sheets[wb.SheetNames[0]]

    const json = XLSX.utils.sheet_to_json(ws, {
      range: 3,
      defval: ''
    })

    setDados(json)

    setLoading(false)
  }

  // ─── Colunas ────────────────────────────────────────────
  const cols = useMemo(() => {

    if (!dados.length) return {}

    const k = Object.keys(dados[0])

    const find = (...terms) =>
      k.find(c =>
        terms.some(t =>
          c.toLowerCase().includes(t.toLowerCase())
        )
      ) || ''

    return {

      unidade: find('NM_LOCAL', 'UNIDADE'),

      medico: find(
        'NOME_MEDICO',
        'NM_PRESTADOR',
        'MEDICO_NOME',
        'MEDICO'
      ),

      esp: find('ESPECIALIDADE'),

      status: find('STATUS'),

      espera: find('TEMPO_DE_ESPERA'),

      qtPacts: find('QT_PACIENTES_AGUARDANDO'),

      uf: find('UF'),

      cidade: find('CIDADE'),

      data: find(
        'HR_INICIO',
        'HR_ENTRADA',
        'DT_',
        'DATA'
      ),
    }

  }, [dados])

  // ─── Período ────────────────────────────────────────────
  const dadosPorPeriodo = useMemo(() => {

    if (período === 'TODOS' || !cols.data) {
      return dados
    }

    const todasDatas = dados
      .map(d => parseDate(d[cols.data]))
      .filter(Boolean)

    return dados.filter(d => {

      const dt = parseDate(d[cols.data])

      return dentroDoPeríodo(
        dt,
        período,
        todasDatas
      )

    })

  }, [dados, cols.data, período])

  // ─── Filters ────────────────────────────────────────────
  const filtered = useMemo(() => {

    let r = dadosPorPeriodo

    if (uf !== 'TODOS') {
      r = r.filter(
        d => String(d[cols.uf] || '').trim() === uf
      )
    }

    if (status !== 'TODOS') {
      r = r.filter(
        d => String(d[cols.status] || '').trim() === status
      )
    }

    if (search) {

      r = r.filter(d =>

        [
          d[cols.unidade],
          d[cols.medico],
          d[cols.esp],
          d[cols.cidade]
        ]

          .some(v =>
            String(v || '')
              .toLowerCase()
              .includes(search.toLowerCase())
          )
      )
    }

    return r

  }, [
    dadosPorPeriodo,
    uf,
    status,
    search,
    cols
  ])

  // ─── KPIs ───────────────────────────────────────────────
  const totalRegistros = filtered.length

  const totalUnidades = new Set(
    filtered.map(d => d[cols.unidade]).filter(Boolean)
  ).size

  const totalMedicos = new Set(
    filtered.map(d => d[cols.medico]).filter(Boolean)
  ).size

  // ─── Médicos corrigido ──────────────────────────────────
  const topMedicos = useMemo(() => {

    const m = {}

    filtered.forEach(d => {

      const nome = String(
        d[cols.medico] || 'Sem Médico'
      ).trim()

      const data = parseDate(d[cols.data])

      const dia = data
        ? data.toLocaleDateString('pt-BR')
        : 'Sem Data'

      const key = `${nome}_${dia}`

      const espera = parseHM(
        d[cols.espera]
      )

      if (!m[key]) {

        m[key] = {
          med: nome,
          total: 0,
          agendas: new Set(),
          dias: new Set(),
        }

      }

      m[key].total += espera

      const agenda =
        d['AGENDA'] ||
        d['NM_AGENDA'] ||
        d['DS_AGENDA'] ||
        'Agenda'

      m[key].agendas.add(agenda)

      m[key].dias.add(dia)

    })

    return Object.values(m)

      .map(x => ({
        med: x.med,
        total: x.total,
        agendas: x.agendas.size,
      }))

      .sort((a, b) => b.total - a.total)

      .slice(0, 10)

  }, [filtered, cols])

  // ─── Impacto fila corrigido ─────────────────────────────
  const impactoFila = useMemo(() => {

    const m = {}

    filtered.forEach(d => {

      const s = String(
        d[cols.status] || ''
      ).toUpperCase()

      if (!s.includes('ATRASO')) return

      const unidade =
        d[cols.unidade] || 'Sem Unidade'

      const medico = String(
        d[cols.medico] || 'Sem Médico'
      ).trim()

      if (!m[unidade]) {

        m[unidade] = {
          unidade,
          medicos: new Set(),
          pacientes: 0,
          agendas: 0,
        }

      }

      m[unidade].medicos.add(medico)

      m[unidade].pacientes +=
        Number(d[cols.qtPacts]) || 0

      m[unidade].agendas++

    })

    return Object.values(m)

      .map(x => ({
        unidade: x.unidade,
        medicos: x.medicos.size,
        pacientes: x.pacientes,
        agendas: x.agendas,
      }))

      .sort((a, b) => b.pacientes - a.pacientes)

      .slice(0, 8)

  }, [filtered, cols])

  return (
    <div style={{
      background: T.bg,
      minHeight: '100vh',
      color: T.text,
      padding: 40
    }}>

      <PeriodoSelector
        value={período}
        onChange={setPeriodo}
      />

      <div style={{ marginTop: 20 }}>

        <label style={{
          background: T.accent,
          color: '#000',
          padding: '10px 20px',
          borderRadius: 10,
          cursor: 'pointer',
          fontWeight: 700
        }}>

          {loading ? 'Carregando...' : '+ Carregar Planilha'}

          <input
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleUpload}
          />

        </label>

      </div>

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(3,1fr)',
        gap: 16,
        marginTop: 30
      }}>

        <Card>
          <SectionHeader>Total Registros</SectionHeader>

          <div style={{
            fontSize: 42,
            fontWeight: 800
          }}>
            {totalRegistros}
          </div>
        </Card>

        <Card>
          <SectionHeader>Total Unidades</SectionHeader>

          <div style={{
            fontSize: 42,
            fontWeight: 800
          }}>
            {totalUnidades}
          </div>
        </Card>

        <Card>
          <SectionHeader>Total Médicos</SectionHeader>

          <div style={{
            fontSize: 42,
            fontWeight: 800
          }}>
            {totalMedicos}
          </div>
        </Card>

      </div>

      <Card style={{ marginTop: 30 }}>

        <SectionHeader>
          👨‍⚕️ Médicos com Maior Tempo de Espera
        </SectionHeader>

        {topMedicos.map((m, i) => (

          <HBar
            key={i}
            rank={i}
            label={`${m.med} (${m.agendas} agendas)`}
            value={parseFloat(m.total.toFixed(1))}
            max={topMedicos[0]?.total || 1}
            color={i < 3 ? T.danger : T.warning}
            unit="h"
          />

        ))}

      </Card>

      <Card style={{ marginTop: 30 }}>

        <SectionHeader>
          📊 Impacto na Fila
        </SectionHeader>

        <table style={{
          width: '100%',
          borderCollapse: 'collapse'
        }}>

          <thead>

            <tr>

              {[
                'Unidade',
                'Médicos em Atraso',
                'Agendas',
                'Pacientes'
              ].map(h => (

                <th
                  key={h}
                  style={{
                    textAlign: 'left',
                    padding: 12,
                    color: T.sub
                  }}
                >
                  {h}
                </th>

              ))}

            </tr>

          </thead>

          <tbody>

            {impactoFila.map((r, i) => (

              <tr key={i}>

                <td style={{ padding: 12 }}>
                  {r.unidade}
                </td>

                <td style={{
                  padding: 12,
                  color: T.danger,
                  fontWeight: 700
                }}>
                  {r.medicos}
                </td>

                <td style={{
                  padding: 12,
                  color: T.accent,
                  fontWeight: 700
                }}>
                  {r.agendas}
                </td>

                <td style={{
                  padding: 12,
                  color: T.warning,
                  fontWeight: 700
                }}>
                  {r.pacientes}
                </td>

              </tr>

            ))}

          </tbody>

        </table>

      </Card>

    </div>
  )
}
