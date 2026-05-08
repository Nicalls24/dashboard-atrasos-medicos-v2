'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Cell,
} from 'recharts'

const COLORS = [
  '#d946ef',
  '#c026d3',
  '#a855f7',
  '#9333ea',
  '#7e22ce',
]

export default function Home() {

  const [rows, setRows] = useState([])
  const [unidadeFiltro, setUnidadeFiltro] = useState('TODAS')

  const handleUpload = (e) => {

    const file = e.target.files[0]

    if (!file) return

    const reader = new FileReader()

    reader.onload = (evt) => {

      const data = evt.target.result

      const workbook = XLSX.read(data, {
        type: 'binary',
      })

      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const json =
        XLSX.utils.sheet_to_json(sheet)

      setRows(json)
    }

    reader.readAsBinaryString(file)
  }

  const unidades = useMemo(() => {

    const lista = rows
      .map(
        (r) =>
          r.NM_LOCAL ||
          r.UNIDADE ||
          ''
      )
      .filter(Boolean)

    return [...new Set(lista)]

  }, [rows])

  const dadosFiltrados = useMemo(() => {

    return rows.filter((r) => {

      const unidade =
        r.NM_LOCAL ||
        r.UNIDADE ||
        ''

      if (
        unidadeFiltro === 'TODAS'
      )
        return true

      return unidade === unidadeFiltro
    })

  }, [rows, unidadeFiltro])

  const parseTempo = (valor) => {

    if (!valor) return 0

    if (typeof valor === 'number')
      return valor

    const texto =
      valor.toString()

    const numeros =
      texto.match(/\d+/g)

    if (!numeros) return 0

    if (numeros.length >= 2) {

      const h = parseInt(
        numeros[0]
      )

      const m = parseInt(
        numeros[1]
      )

      return h + m / 60
    }

    return parseFloat(
      numeros[0]
    )
  }

  const totalPacientes =
    dadosFiltrados.length

  const totalUnidades =
    new Set(
      dadosFiltrados.map(
        (r) =>
          r.NM_LOCAL ||
          'SEM UNIDADE'
      )
    ).size

  const totalMedicos =
    new Set(
      dadosFiltrados.map(
        (r) =>
          r.NM_PRESTADOR ||
          r.MEDICO ||
          'SEM MÉDICO'
      )
    ).size

  const mediaEspera = useMemo(() => {

    const tempos =
      dadosFiltrados.map((r) =>
        parseTempo(
          r.TEMPO_ESPERA ||
            r.TEMPO ||
            r.ESPERA
        )
      )

    if (!tempos.length)
      return 0

    const soma = tempos.reduce(
      (a, b) => a + b,
      0
    )

    return (
      soma / tempos.length
    ).toFixed(1)

  }, [dadosFiltrados])

  const rankingPacientes =
    useMemo(() => {

      const mapa = {}

      dadosFiltrados.forEach(
        (r) => {

          const unidade =
            r.NM_LOCAL ||
            'SEM UNIDADE'

          mapa[unidade] =
            (mapa[unidade] || 0) +
            1
        }
      )

      return Object.entries(
        mapa
      )
        .map(([name, value]) => ({
          name,
          value,
        }))
        .sort(
          (a, b) =>
            b.value - a.value
        )
        .slice(0, 10)

    }, [dadosFiltrados])

  const statusData = useMemo(() => {

    const mapa = {}

    dadosFiltrados.forEach(
      (r) => {

        const status =
          r.STATUS ||
          r.STATUS_MEDICO ||
          r.STATUS_PONTO ||
          'OK'

        mapa[status] =
          (mapa[status] || 0) + 1
      }
    )

    return Object.entries(
      mapa
    ).map(([name, value]) => ({
      name,
      value,
    }))

  }, [dadosFiltrados])

  const medicosAtraso =
    useMemo(() => {

      const mapa = {}

      dadosFiltrados.forEach(
        (r) => {

          const medico =
            r.NM_PRESTADOR ||
            r.MEDICO ||
            'SEM MÉDICO'

          const unidade =
            r.NM_LOCAL ||
            'SEM UNIDADE'

          const tempo =
            parseTempo(
              r.TEMPO_ESPERA ||
                r.TEMPO ||
                r.ESPERA
            )

          if (
            !mapa[medico]
          ) {

            mapa[medico] = {
              medico,
              unidade,
              tempo,
            }

          } else {

            if (
              tempo >
              mapa[medico].tempo
            ) {

              mapa[medico].tempo =
                tempo
            }
          }
        }
      )

      return Object.values(
        mapa
      )
        .sort(
          (a, b) =>
            b.tempo - a.tempo
        )
        .slice(0, 10)

    }, [dadosFiltrados])

  const topEspecialidades =
    useMemo(() => {

      const mapa = {}

      dadosFiltrados.forEach(
        (r) => {

          const esp =
            r.ESPECIALIDADE ||
            'SEM ESPECIALIDADE'

          const tempo =
            parseTempo(
              r.TEMPO_ESPERA ||
                r.TEMPO ||
                r.ESPERA
            )

          mapa[esp] =
            (mapa[esp] || 0) +
            tempo
        }
      )

      return Object.entries(
        mapa
      )
        .map(([name, value]) => ({
          name,
          value:
            value.toFixed(1),
        }))
        .sort(
          (a, b) =>
            b.value - a.value
        )
        .slice(0, 3)

    }, [dadosFiltrados])

  const impactoFila =
    useMemo(() => {

      const mapa = {}

      dadosFiltrados.forEach(
        (r) => {

          const status = (
            r.STATUS ||
            ''
          ).toUpperCase()

          if (
            !status.includes(
              'ATRASO'
            )
          )
            return

          const unidade =
            r.NM_LOCAL ||
            'SEM UNIDADE'

          if (
            !mapa[unidade]
          ) {

            mapa[unidade] = {
              unidade,
              medicos:
                new Set(),
              pacientes: 0,
            }
          }

          mapa[
            unidade
          ].medicos.add(
            r.NM_PRESTADOR ||
              r.MEDICO ||
              'SEM MÉDICO'
          )

          mapa[
            unidade
          ].pacientes += 1
        }
      )

      return Object.values(
        mapa
      )
        .map((r) => ({
          unidade:
            r.unidade,
          qtdMedicos:
            r.medicos.size,
          pacientes:
            r.pacientes,
        }))
        .sort(
          (a, b) =>
            b.qtdMedicos -
            a.qtdMedicos
        )
        .slice(0, 10)

    }, [dadosFiltrados])

  return (

    <main style={styles.page}>

      <div style={styles.header}>

        <div>

          <h1 style={styles.title}>
            Dashboard Atrasos Médicos
          </h1>

          <p style={styles.subtitle}>
            Monitoramento operacional hospitalar
          </p>

        </div>

        <label style={styles.uploadButton}>

          Carregar Planilha

          <input
            type='file'
            accept='.xlsx,.xls'
            onChange={handleUpload}
            style={{
              display: 'none',
            }}
          />

        </label>

      </div>

      <div style={styles.filterRow}>

        <select
          value={unidadeFiltro}
          onChange={(e) =>
            setUnidadeFiltro(
              e.target.value
            )
          }
          style={styles.select}
        >

          <option value='TODAS'>
            TODAS
          </option>

          {unidades.map(
            (u) => (

              <option
                key={u}
                value={u}
              >

                {u}

              </option>
            )
          )}

        </select>

        <button
          onClick={() =>
            setUnidadeFiltro(
              'TODAS'
            )
          }
          style={styles.resetButton}
        >

          Redefinir

        </button>

      </div>

      <div style={styles.cards}>

        <div style={styles.card}>

          <h3>Total Pacientes</h3>

          <h2>
            {totalPacientes}
          </h2>

        </div>

        <div style={styles.card}>

          <h3>Unidades</h3>

          <h2>
            {totalUnidades}
          </h2>

        </div>

        <div style={styles.card}>

          <h3>Médicos</h3>

          <h2>
            {totalMedicos}
          </h2>

        </div>

        <div style={styles.card}>

          <h3>
            Tempo Médio Espera
          </h3>

          <h2>
            {mediaEspera}h
          </h2>

        </div>

        <div style={styles.card}>

          <h3>
            Top Especialidades Críticas
          </h3>

          {topEspecialidades.map(
            (e, i) => (

              <div
                key={i}
                style={{
                  marginTop: 14,
                }}
              >

                <div
                  style={{
                    display:
                      'flex',
                    justifyContent:
                      'space-between',
                    marginBottom: 4,
                  }}
                >

                  <span>
                    {e.name}
                  </span>

                  <strong>
                    {e.value}h
                  </strong>

                </div>

                <div
                  style={{
                    width: '100%',
                    height: 10,
                    borderRadius: 999,
                    background:
                      '#312e81',
                  }}
                >

                  <div
                    style={{
                      width: '80%',
                      height:
                        '100%',
                      borderRadius: 999,
                      background:
                        COLORS[
                          i
                        ],
                    }}
                  />

                </div>

              </div>
            )
          )}

        </div>

      </div>

      <div style={styles.grid}>

        <div style={styles.chartCard}>

          <h2>
            Ranking Crítico de Pacientes Aguardando
          </h2>

          <ResponsiveContainer
            width='100%'
            height={420}
          >

            <BarChart
              data={
                rankingPacientes
              }
            >

              <CartesianGrid
                strokeDasharray='3 3'
                stroke='#4c1d95'
              />

              <XAxis
                dataKey='name'
                stroke='#ddd6fe'
              />

              <YAxis stroke='#ddd6fe' />

              <Tooltip />

              <Bar
                dataKey='value'
                radius={[
                  10, 10, 0, 0,
                ]}
              >

                {rankingPacientes.map(
                  (
                    entry,
                    index
                  ) => (

                    <Cell
                      key={index}
                      fill={
                        COLORS[
                          index %
                            COLORS.length
                        ]
                      }
                    />

                  )
                )}

              </Bar>

            </BarChart>

          </ResponsiveContainer>

        </div>

        <div style={styles.chartCard}>

          <h2>
            Status de Pontos Médicos
          </h2>

          <ResponsiveContainer
            width='100%'
            height={420}
          >

            <BarChart
              data={statusData}
              layout='vertical'
              margin={{
                left: 60,
              }}
            >

              <CartesianGrid
                strokeDasharray='3 3'
                stroke='#4c1d95'
              />

              <XAxis
                type='number'
                stroke='#ddd6fe'
              />

              <YAxis
                dataKey='name'
                type='category'
                stroke='#ddd6fe'
                width={160}
              />

              <Tooltip />

              <Bar
                dataKey='value'
                radius={[
                  0, 12, 12, 0,
                ]}
              >

                {statusData.map(
                  (
                    entry,
                    index
                  ) => (

                    <Cell
                      key={index}
                      fill={
                        COLORS[
                          index %
                            COLORS.length
                        ]
                      }
                    />

                  )
                )}

              </Bar>

            </BarChart>

          </ResponsiveContainer>

        </div>

      </div>

      <div style={styles.fullCard}>

        <h2>
          Médicos com Maior Tempo de Atraso
        </h2>

        <ResponsiveContainer
          width='100%'
          height={420}
        >

          <BarChart
            data={medicosAtraso}
            layout='vertical'
            margin={{
              left: 120,
            }}
          >

            <CartesianGrid
              strokeDasharray='3 3'
              stroke='#4c1d95'
            />

            <XAxis
              type='number'
              stroke='#ddd6fe'
            />

            <YAxis
              type='category'
              dataKey='medico'
              stroke='#ddd6fe'
              width={220}
            />

            <Tooltip />

            <Bar
              dataKey='tempo'
              radius={[
                0, 10, 10, 0,
              ]}
            >

              {medicosAtraso.map(
                (
                  entry,
                  index
                ) => (

                  <Cell
                    key={index}
                    fill={
                      COLORS[
                        index %
                          COLORS.length
                      ]
                    }
                  />

                )
              )}

            </Bar>

          </BarChart>

        </ResponsiveContainer>

      </div>

      <div style={styles.fullCard}>

        <h2>
          Impacto na Fila de Espera
        </h2>

        <table
          style={styles.table}
        >

          <thead>

            <tr>

              <th>
                Unidade
              </th>

              <th>
                Qtd Médicos
              </th>

              <th>
                Pacientes
              </th>

            </tr>

          </thead>

          <tbody>

            {impactoFila.map(
              (r, i) => (

                <tr key={i}>

                  <td>
                    {
                      r.unidade
                    }
                  </td>

                  <td>
                    {
                      r.qtdMedicos
                    }
                  </td>

                  <td>
                    {
                      r.pacientes
                    }
                  </td>

                </tr>
              )
            )}

          </tbody>

        </table>

      </div>

    </main>
  )
}

const styles = {

  page: {
    minHeight: '100vh',
    background:
      'linear-gradient(180deg,#070b24,#1f1b4b)',
    padding: '28px',
    color: '#fff',
    fontFamily:
      'Arial, sans-serif',
  },

  header: {
    display: 'flex',
    justifyContent:
      'space-between',
    alignItems: 'center',
    marginBottom: '30px',
  },

  title: {
    fontSize: '56px',
    marginBottom: 10,
  },

  subtitle: {
    color: '#c4b5fd',
    fontSize: 24,
  },

  uploadButton: {
    background:
      'linear-gradient(90deg,#9333ea,#d946ef)',
    padding:
      '18px 28px',
    borderRadius: 16,
    cursor: 'pointer',
    fontWeight: 'bold',
    fontSize: 18,
  },

  filterRow: {
    display: 'flex',
    gap: 16,
    marginBottom: 30,
  },

  select: {
    padding:
      '14px 18px',
    borderRadius: 12,
    border: 'none',
    fontSize: 16,
    minWidth: 220,
  },

  resetButton: {
    background:
      '#9333ea',
    border: 'none',
    color: '#fff',
    padding:
      '14px 24px',
    borderRadius: 12,
    cursor: 'pointer',
    fontWeight: 'bold',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(5,1fr)',
    gap: 20,
    marginBottom: 30,
  },

  card: {
    background:
      'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.08)',
    borderRadius: 24,
    padding: 28,
    minHeight: 180,
  },

  grid: {
    display: 'grid',
    gridTemplateColumns:
      '1fr 1fr',
    gap: 24,
    marginBottom: 30,
  },

  chartCard: {
    background:
      'rgba(255,255,255,0.05)',
    borderRadius: 24,
    padding: 20,
  },

  fullCard: {
    background:
      'rgba(255,255,255,0.05)',
    borderRadius: 24,
    padding: 20,
    marginBottom: 30,
  },

  table: {
    width: '100%',
    borderCollapse:
      'collapse',
  },
}
