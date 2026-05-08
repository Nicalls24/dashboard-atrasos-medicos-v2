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

export default function Home() {
  const [dados, setDados] = useState([])
  const [unidadeFiltro, setUnidadeFiltro] =
    useState('TODAS')

  const cores = [
    '#c84df2',
    '#8000db',
    '#921c9e',
    '#6d3fc2',
    '#7b30f2',
  ]

  const handleUpload = (e) => {
    const arquivo = e.target.files[0]

    if (!arquivo) return

    const reader = new FileReader()

    reader.onload = (evento) => {
      const data = evento.target.result

      const workbook = XLSX.read(data, {
        type: 'binary',
      })

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const json = XLSX.utils.sheet_to_json(
        worksheet,
        {
          range: 3,
          defval: '',
        }
      )

      setDados(json)
    }

    reader.readAsBinaryString(arquivo)
  }

  const unidades = useMemo(() => {
    const lista = dados
      .map((item) => item['NM_FILIAL'])
      .filter(Boolean)

    return ['TODAS', ...new Set(lista)]
  }, [dados])

  const dadosFiltrados = useMemo(() => {
    if (unidadeFiltro === 'TODAS')
      return dados

    return dados.filter(
      (item) =>
        item['NM_FILIAL'] === unidadeFiltro
    )
  }, [dados, unidadeFiltro])

  const pacientesAguardando =
    dadosFiltrados.reduce(
      (acc, item) =>
        acc +
        Number(
          item[
            ' QT_PACIENTES_AGUARDANDO'
          ] || 0
        ),
      0
    )

  const medicosAtrasados =
    dadosFiltrados.filter(
      (item) =>
        String(item['ATRASO'])
          .toUpperCase()
          .includes('SIM')
    ).length

  const percentualAtraso =
    dadosFiltrados.length > 0
      ? (
          (medicosAtrasados /
            dadosFiltrados.length) *
          100
        ).toFixed(1)
      : 0

  const maiorTempoEspera = useMemo(() => {
    const tempos = dadosFiltrados
      .map((item) => {
        const valor =
          item['TEMPO_DE_ESPERA']

        if (!valor) return 0

        if (typeof valor === 'number') {
          return valor * 24 * 60
        }

        if (typeof valor === 'string') {
          const partes = valor.split(':')

          if (partes.length >= 2) {
            const horas =
              Number(partes[0]) || 0

            const minutos =
              Number(partes[1]) || 0

            return horas * 60 + minutos
          }
        }

        return 0
      })
      .filter((v) => v > 0)

    if (tempos.length === 0)
      return 0

    return Math.max(...tempos)
  }, [dadosFiltrados])

  const tempoEsperaMedio = useMemo(() => {
    const tempos = dadosFiltrados
      .map((item) => {
        const valor =
          item['TEMPO_DE_ESPERA']

        if (!valor) return 0

        if (
          typeof valor === 'number'
        ) {
          return valor * 24 * 60
        }

        if (
          typeof valor === 'string'
        ) {
          const partes =
            valor.split(':')

          if (partes.length >= 2) {
            const horas =
              Number(partes[0]) || 0

            const minutos =
              Number(partes[1]) || 0

            return horas * 60 + minutos
          }
        }

        return 0
      })
      .filter((v) => v > 0)

    if (tempos.length === 0)
      return 0

    const media =
      tempos.reduce(
        (a, b) => a + b,
        0
      ) / tempos.length

    return Math.round(media)
  }, [dadosFiltrados])

  const rankingCritico =
    Object.values(
      dadosFiltrados.reduce(
        (acc, item) => {
          const unidade =
            item['NM_FILIAL'] ||
            'SEM UNIDADE'

          if (!acc[unidade]) {
            acc[unidade] = {
              unidade,
              total: 0,
            }
          }

          acc[unidade].total +=
            Number(
              item[
                ' QT_PACIENTES_AGUARDANDO'
              ] || 0
            )

          return acc
        },
        {}
      )
    )
      .sort(
        (a, b) =>
          b.total - a.total
      )
      .slice(0, 10)

  const statusMedico =
    Object.values(
      dadosFiltrados.reduce(
        (acc, item) => {
          const status =
            item['STATUS'] ||
            'SEM STATUS'

          if (!acc[status]) {
            acc[status] = {
              motivo: status,
              quantidade: 0,
            }
          }

          acc[status].quantidade += 1

          return acc
        },
        {}
      )
    )

  const tabelaCritica = Object.values(
    dadosFiltrados.reduce(
      (acc, item) => {
        const status = String(
          item['STATUS']
        ).toUpperCase()

        if (
          !status.includes('ATRASO')
        )
          return acc

        const medico =
          item['NM_MEDICO'] ||
          'SEM MÉDICO'

        const unidade =
          item['NM_FILIAL'] ||
          'SEM UNIDADE'

        const chave =
          medico + unidade

        if (!acc[chave]) {
          acc[chave] = {
            unidade,
            medico,
            pacientes: 0,
            status,
          }
        }

        acc[chave].pacientes +=
          Number(
            item[
              ' QT_PACIENTES_AGUARDANDO'
            ] || 0
          )

        return acc
      },
      {}
    )
  )

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <div>
          <h1 style={styles.title}>
            Dashboard Atrasos Médicos
          </h1>

          <p style={styles.subtitle}>
            Gestão operacional hospitalar
          </p>
        </div>

        <label style={styles.uploadButton}>
          Carregar Planilhas

          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUpload}
            style={{
              display: 'none',
            }}
          />
        </label>
      </div>

      <div style={styles.filters}>
        <select
          value={unidadeFiltro}
          onChange={(e) =>
            setUnidadeFiltro(
              e.target.value
            )
          }
          style={styles.select}
        >
          {unidades.map(
            (unidade, index) => (
              <option key={index}>
                {unidade}
              </option>
            )
          )}
        </select>

        <button
          style={styles.resetButton}
          onClick={() =>
            setUnidadeFiltro(
              'TODAS'
            )
          }
        >
          Redefinir
        </button>
      </div>

      <div style={styles.cards}>
        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Médicos em Atraso
          </h3>

          <p style={styles.cardValue}>
            {medicosAtrasados}
          </p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            % Em Atraso
          </h3>

          <p style={styles.cardValue}>
            {percentualAtraso}%
          </p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Pacientes Aguardando
          </h3>

          <p style={styles.cardValue}>
            {
              pacientesAguardando
            }
          </p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Tempo Médio Espera
          </h3>

          <p style={styles.cardValue}>
            {tempoEsperaMedio}{' '}
            min
          </p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Maior Tempo de Espera
          </h3>

          <p style={styles.cardValue}>
            {maiorTempoEspera} min
          </p>
        </div>
      </div>

      <div style={styles.grid}>
        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>
            Ranking Crítico de
            Pacientes Aguardando
          </h2>

          <ResponsiveContainer
            width="100%"
            height={350}
          >
            <BarChart
              data={
                rankingCritico
              }
            >
              <CartesianGrid stroke="#35204d" />

              <XAxis
                dataKey="unidade"
                stroke="#d7c6ff"
              />

              <YAxis stroke="#d7c6ff" />

              <Tooltip />

              <Bar
                dataKey="total"
                fill="#c84df2"
                radius={[
                  10, 10, 0, 0,
                ]}
              />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>
            Status de Pontos
            Médicos
          </h2>

          <ResponsiveContainer
            width="100%"
            height={350}
          >
            <BarChart data={statusMedico}>
              <CartesianGrid stroke="#35204d" />

              <XAxis
                dataKey="motivo"
                stroke="#d7c6ff"
              />

              <YAxis stroke="#d7c6ff" />

              <Tooltip />

              <Bar
                dataKey="quantidade"
                radius={[
                  10, 10, 0, 0,
                ]}
              >
                {statusMedico.map(
                  (
                    entry,
                    index
                  ) => (
                    <Cell
                      key={index}
                      fill={
                        cores[
                          index %
                            cores.length
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

      <div style={styles.tableContainer}>
        <h2 style={styles.chartTitle}>
          Impacto na Fila de
          Espera
        </h2>

        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>
                Unidade
              </th>

              <th style={styles.th}>
                Médico
              </th>

              <th style={styles.th}>
                Pacientes
              </th>

              <th style={styles.th}>
                Status
              </th>
            </tr>
          </thead>

          <tbody>
            {tabelaCritica
              .sort(
                (a, b) =>
                  b.pacientes -
                  a.pacientes
              )
              .slice(0, 20)
              .map(
                (
                  item,
                  index
                ) => (
                  <tr key={index}>
                    <td
                      style={
                        styles.td
                      }
                    >
                      {
                        item.unidade
                      }
                    </td>

                    <td
                      style={
                        styles.td
                      }
                    >
                      {
                        item.medico
                      }
                    </td>

                    <td
                      style={
                        styles.td
                      }
                    >
                      {
                        item.pacientes
                      }
                    </td>

                    <td
                      style={
                        styles.td
                      }
                    >
                      {
                        item.status
                      }
                    </td>
                  </tr>
                )
              )}
          </tbody>
        </table>
      </div>
    </div>
  )
}

const styles = {
  page: {
    minHeight: '100vh',
    background:
      'linear-gradient(180deg,#060816,#0d1230)',
    padding: '30px',
    color: '#fff',
    fontFamily: 'Arial',
  },

  header: {
    display: 'flex',
    justifyContent:
      'space-between',
    alignItems: 'center',
    marginBottom: '30px',
    flexWrap: 'wrap',
  },

  title: {
    fontSize: '46px',
    marginBottom: '10px',
  },

  subtitle: {
    color: '#b8a7d9',
    fontSize: '18px',
  },

  uploadButton: {
    background:
      'linear-gradient(135deg,#8000db,#c84df2)',
    padding: '14px 28px',
    borderRadius: '14px',
    cursor: 'pointer',
    fontWeight: 'bold',
    color: '#fff',
    boxShadow:
      '0 0 20px rgba(128,0,219,0.4)',
  },

  filters: {
    display: 'flex',
    gap: '15px',
    marginBottom: '30px',
  },

  select: {
    padding: '12px',
    borderRadius: '12px',
    minWidth: '250px',
    border: 'none',
  },

  resetButton: {
    background:
      'linear-gradient(135deg,#6d3fc2,#7b30f2)',
    border: 'none',
    color: '#fff',
    padding: '12px 24px',
    borderRadius: '12px',
    fontWeight: 'bold',
    cursor: 'pointer',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit,minmax(220px,1fr))',
    gap: '20px',
    marginBottom: '30px',
  },

  card: {
    background:
      'rgba(255,255,255,0.04)',
    backdropFilter: 'blur(12px)',
    border:
      '1px solid rgba(255,255,255,0.08)',
    borderRadius: '24px',
    padding: '30px',
  },

  cardTitle: {
    color: '#d7c6ff',
    marginBottom: '25px',
    fontSize: '20px',
  },

  cardValue: {
    fontSize: '48px',
    fontWeight: 'bold',
  },

  grid: {
    display: 'grid',
    gridTemplateColumns:
      '1fr 1fr',
    gap: '20px',
    marginBottom: '30px',
  },

  chartBox: {
    background:
      'rgba(255,255,255,0.04)',
    borderRadius: '24px',
    padding: '20px',
    border:
      '1px solid rgba(255,255,255,0.08)',
  },

  chartTitle: {
    fontSize: '30px',
    marginBottom: '20px',
  },

  tableContainer: {
    background:
      'rgba(255,255,255,0.04)',
    borderRadius: '24px',
    padding: '20px',
    border:
      '1px solid rgba(255,255,255,0.08)',
  },

  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },

  th: {
    background:
      'rgba(255,255,255,0.08)',
    padding: '16px',
    textAlign: 'left',
  },

  td: {
    padding: '16px',
    borderBottom:
      '1px solid rgba(255,255,255,0.08)',
  },
}
