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
  PieChart,
  Pie,
  Cell,
} from 'recharts'

export default function Home() {
  const [dados, setDados] = useState([])
  const [unidadeFiltro, setUnidadeFiltro] = useState('TODAS')

  const cores = ['#ef4444', '#f59e0b', '#38bdf8', '#22c55e']

  const handleUpload = (e) => {
    const arquivo = e.target.files[0]

    if (!arquivo) return

    const reader = new FileReader()

    reader.onload = (evento) => {
      const data = evento.target.result

      const workbook = XLSX.read(data, {
        type: 'binary',
      })

      const sheetName = workbook.SheetNames[0]

      const worksheet = workbook.Sheets[sheetName]

      const json = XLSX.utils.sheet_to_json(worksheet, {
        range: 3,
        defval: '',
      })

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
    if (unidadeFiltro === 'TODAS') {
      return dados
    }

    return dados.filter(
      (item) => item['NM_FILIAL'] === unidadeFiltro
    )
  }, [dados, unidadeFiltro])

  const totalRegistros = dadosFiltrados.length

  const medicosAtrasados = dadosFiltrados.filter(
    (item) => item['ATRASO'] === 'SIM'
  ).length

  const percentualAtraso =
    totalRegistros > 0
      ? ((medicosAtrasados / totalRegistros) * 100).toFixed(1)
      : 0

  const pacientesAguardando = dadosFiltrados.reduce(
    (acc, item) =>
      acc + Number(item[' QT_PACIENTES_AGUARDANDO'] || 0),
    0
  )

  const tempoEsperaMedio = (
    dadosFiltrados.reduce((acc, item) => {
      const tempo = item['TEMPO_DE_ESPERA']

      if (!tempo) return acc

      const segundos =
        tempo?.h * 3600 +
        tempo?.m * 60 +
        tempo?.s

      return acc + segundos
    }, 0) / (dadosFiltrados.length || 1)
  )

  const tempoMedioMinutos = Math.round(
    tempoEsperaMedio / 60
  )

  const topCriticos = Object.values(
    dadosFiltrados.reduce((acc, item) => {
      const unidade = item['NM_FILIAL'] || 'SEM UNIDADE'

      if (!acc[unidade]) {
        acc[unidade] = {
          unidade,
          pacientes: 0,
          atrasos: 0,
        }
      }

      acc[unidade].pacientes += Number(
        item[' QT_PACIENTES_AGUARDANDO'] || 0
      )

      if (item['ATRASO'] === 'SIM') {
        acc[unidade].atrasos += 1
      }

      return acc
    }, {})
  )
    .map((item) => ({
      ...item,
      impacto:
        item.pacientes + item.atrasos * 2,
    }))
    .sort((a, b) => b.impacto - a.impacto)
    .slice(0, 10)

  const statusData = Object.values(
    dadosFiltrados.reduce((acc, item) => {
      const status = item['STATUS'] || 'SEM STATUS'

      if (!acc[status]) {
        acc[status] = {
          name: status,
          value: 0,
        }
      }

      acc[status].value += 1

      return acc
    }, {})
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

        <div style={styles.uploadBox}>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUpload}
          />
        </div>
      </div>

      <div style={styles.filters}>
        <select
          value={unidadeFiltro}
          onChange={(e) =>
            setUnidadeFiltro(e.target.value)
          }
          style={styles.select}
        >
          {unidades.map((unidade, index) => (
            <option key={index}>
              {unidade}
            </option>
          ))}
        </select>

        <button
          style={styles.button}
          onClick={() => setUnidadeFiltro('TODAS')}
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
            {pacientesAguardando}
          </p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Tempo Médio Espera
          </h3>

          <p style={styles.cardValue}>
            {tempoMedioMinutos} min
          </p>
        </div>
      </div>

      <div style={styles.grid}>
        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>
            Ranking Crítico
          </h2>

          <ResponsiveContainer
            width="100%"
            height={350}
          >
            <BarChart data={topCriticos}>
              <CartesianGrid stroke="#16385f" />

              <XAxis
                dataKey="unidade"
                stroke="#9fb3c8"
              />

              <YAxis stroke="#9fb3c8" />

              <Tooltip />

              <Bar
                dataKey="impacto"
                fill="#ef4444"
                radius={[8, 8, 0, 0]}
              />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>
            Status Operacional
          </h2>

          <ResponsiveContainer
            width="100%"
            height={350}
          >
            <PieChart>
              <Pie
                data={statusData}
                dataKey="value"
                nameKey="name"
                outerRadius={120}
                label
              >
                {statusData.map(
                  (entry, index) => (
                    <Cell
                      key={index}
                      fill={
                        cores[
                          index % cores.length
                        ]
                      }
                    />
                  )
                )}
              </Pie>

              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={styles.tableContainer}>
        <h2 style={styles.chartTitle}>
          Impacto na Fila de Espera
        </h2>

        <div style={styles.tableWrapper}>
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
                  Especialidade
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
              {dadosFiltrados
                .sort(
                  (a, b) =>
                    Number(
                      b[
                        ' QT_PACIENTES_AGUARDANDO'
                      ] || 0
                    ) -
                    Number(
                      a[
                        ' QT_PACIENTES_AGUARDANDO'
                      ] || 0
                    )
                )
                .slice(0, 20)
                .map((item, index) => (
                  <tr key={index}>
                    <td style={styles.td}>
                      {item['NM_FILIAL']}
                    </td>

                    <td style={styles.td}>
                      {item['NM_MEDICO']}
                    </td>

                    <td style={styles.td}>
                      {
                        item[
                          'DS_ESPECIALIDADE'
                        ]
                      }
                    </td>

                    <td style={styles.td}>
                      {
                        item[
                          ' QT_PACIENTES_AGUARDANDO'
                        ]
                      }
                    </td>

                    <td style={styles.td}>
                      {item['STATUS']}
                    </td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

const styles = {
  page: {
    minHeight: '100vh',
    background: '#061529',
    padding: '30px',
    color: '#fff',
    fontFamily: 'Arial',
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '30px',
    flexWrap: 'wrap',
    gap: '20px',
  },

  title: {
    fontSize: '42px',
    marginBottom: '10px',
  },

  subtitle: {
    color: '#9fb3c8',
    fontSize: '18px',
  },

  uploadBox: {
    background: '#0d223d',
    padding: '20px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  filters: {
    display: 'flex',
    gap: '15px',
    marginBottom: '30px',
    flexWrap: 'wrap',
  },

  select: {
    padding: '12px',
    borderRadius: '10px',
    border: 'none',
    minWidth: '250px',
  },

  button: {
    padding: '12px 20px',
    borderRadius: '10px',
    border: 'none',
    background: '#38bdf8',
    color: '#fff',
    fontWeight: 'bold',
    cursor: 'pointer',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit, minmax(220px, 1fr))',
    gap: '20px',
    marginBottom: '30px',
  },

  card: {
    background: '#0d223d',
    padding: '25px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  cardTitle: {
    color: '#8fb3d9',
    marginBottom: '15px',
  },

  cardValue: {
    fontSize: '42px',
    fontWeight: 'bold',
  },

  grid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '20px',
    marginBottom: '30px',
  },

  chartBox: {
    background: '#0d223d',
    padding: '20px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  chartTitle: {
    marginBottom: '20px',
    fontSize: '28px',
  },

  tableContainer: {
    background: '#0d223d',
    padding: '20px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  tableWrapper: {
    overflowX: 'auto',
  },

  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },

  th: {
    background: '#102944',
    padding: '16px',
    textAlign: 'left',
  },

  td: {
    padding: '16px',
    borderBottom: '1px solid #16385f',
  },
}
