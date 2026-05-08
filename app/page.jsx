'use client'

import { useState } from 'react'
import * as XLSX from 'xlsx'
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from 'recharts'

export default function Home() {
  const [dados, setDados] = useState([])
  const [unidadeFiltro, setUnidadeFiltro] = useState('TODAS')

  const cores = ['#38bdf8', '#22c55e', '#f59e0b', '#ef4444']

  const handleUpload = (e) => {
    const arquivo = e.target.files[0]

    if (!arquivo) return

    const reader = new FileReader()

    reader.onload = (evt) => {
      const data = evt.target.result

      const workbook = XLSX.read(data, { type: 'binary' })

      const sheetName = workbook.SheetNames[0]

      const worksheet = workbook.Sheets[sheetName]

      const json = XLSX.utils.sheet_to_json(worksheet)

      setDados(json)
    }

    reader.readAsBinaryString(arquivo)
  }

  const unidades = [
    'TODAS',
    ...new Set(dados.map((item) => item.NM_FILIAL).filter(Boolean)),
  ]

  const dadosFiltrados =
    unidadeFiltro === 'TODAS'
      ? dados
      : dados.filter((item) => item.NM_FILIAL === unidadeFiltro)

  const totalPacientes = dadosFiltrados.length

  const topUnidades = Object.values(
    dadosFiltrados.reduce((acc, item) => {
      const unidade = item.NM_FILIAL || 'Sem unidade'

      if (!acc[unidade]) {
        acc[unidade] = {
          unidade,
          total: 0,
        }
      }

      acc[unidade].total += 1

      return acc
    }, {})
  )

  const statusData = Object.values(
    dadosFiltrados.reduce((acc, item) => {
      const status = item.STATUS || 'Sem Status'

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

  const redefinirFiltros = () => {
    setUnidadeFiltro('TODAS')
  }

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <div>
          <h1 style={styles.title}>Dashboard Atrasos Médicos</h1>

          <p style={styles.subtitle}>
            Monitoramento operacional hospitalar
          </p>
        </div>

        <div style={styles.uploadContainer}>
          <input type="file" accept=".xlsx, .xls" onChange={handleUpload} />
        </div>
      </div>

      <div style={styles.filters}>
        <select
          value={unidadeFiltro}
          onChange={(e) => setUnidadeFiltro(e.target.value)}
          style={styles.select}
        >
          {unidades.map((unidade, index) => (
            <option key={index}>{unidade}</option>
          ))}
        </select>

        <button onClick={redefinirFiltros} style={styles.button}>
          Redefinir
        </button>
      </div>

      <div style={styles.cards}>
        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Total Registros</h3>
          <p style={styles.cardValue}>{totalPacientes}</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Unidades</h3>
          <p style={styles.cardValue}>{unidades.length - 1}</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Status Diferentes</h3>
          <p style={styles.cardValue}>{statusData.length}</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Médicos</h3>
          <p style={styles.cardValue}>
            {
              new Set(
                dadosFiltrados.map((item) => item.NM_MEDICO).filter(Boolean)
              ).size
            }
          </p>
        </div>
      </div>

      <div style={styles.grid}>
        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>Top Unidades</h2>

          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={topUnidades}>
              <XAxis dataKey="unidade" stroke="#fff" />
              <YAxis stroke="#fff" />
              <Tooltip />
              <Bar dataKey="total" fill="#38bdf8" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>Distribuição Status</h2>

          <ResponsiveContainer width="100%" height={320}>
            <PieChart>
              <Pie
                data={statusData}
                dataKey="value"
                nameKey="name"
                outerRadius={120}
                label
              >
                {statusData.map((entry, index) => (
                  <Cell
                    key={`cell-${index}`}
                    fill={cores[index % cores.length]}
                  />
                ))}
              </Pie>
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={styles.tableContainer}>
        <h2 style={styles.chartTitle}>Dados Operacionais</h2>

        <div style={styles.tableWrapper}>
          <table style={styles.table}>
            <thead>
              <tr>
                <th style={styles.th}>Unidade</th>
                <th style={styles.th}>Médico</th>
                <th style={styles.th}>Especialidade</th>
                <th style={styles.th}>Status</th>
              </tr>
            </thead>

            <tbody>
              {dadosFiltrados.slice(0, 20).map((item, index) => (
                <tr key={index}>
                  <td style={styles.td}>{item.NM_FILIAL}</td>
                  <td style={styles.td}>{item.NM_MEDICO}</td>
                  <td style={styles.td}>{item.DS_ESPECIALIDADE}</td>
                  <td style={styles.td}>{item.STATUS}</td>
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

  uploadContainer: {
    background: '#0d223d',
    padding: '20px',
    borderRadius: '16px',
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
    minWidth: '220px',
  },

  button: {
    padding: '12px 18px',
    borderRadius: '10px',
    border: 'none',
    background: '#38bdf8',
    color: '#fff',
    fontWeight: 'bold',
    cursor: 'pointer',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
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
    fontSize: '38px',
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
    padding: '14px',
    textAlign: 'left',
  },

  td: {
    padding: '14px',
    borderBottom: '1px solid #16385f',
  },
}
