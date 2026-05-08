'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  FunnelChart,
  Funnel,
  LabelList,
  Legend
} from 'recharts'

export default function Home() {

  const [data, setData] = useState([])
  const [selectedUnit, setSelectedUnit] = useState('TODAS')

  function handleFile(event) {

    const file = event.target.files[0]

    if (!file) return

    const reader = new FileReader()

    reader.onload = (e) => {

      const workbook = XLSX.read(
        e.target.result,
        { type: 'binary' }
      )

      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const json =
        XLSX.utils.sheet_to_json(sheet)

      setData(json)
    }

    reader.readAsBinaryString(file)
  }

  const unidades = useMemo(() => {

    const lista = data.map(item =>
      item.NM_LOCAL ||
      item.UNIDADE ||
      item.Local
    )

    return [
      'TODAS',
      ...new Set(lista.filter(Boolean))
    ]

  }, [data])

  const filteredData = useMemo(() => {

    if (selectedUnit === 'TODAS') {
      return data
    }

    return data.filter(item => {

      const unidade =
        item.NM_LOCAL ||
        item.UNIDADE ||
        item.Local

      return unidade === selectedUnit
    })

  }, [data, selectedUnit])

  function extrairMinutos(valor) {

    if (!valor) return 0

    const texto = String(valor)

    const matchHoras =
      texto.match(/(\d+)\s*h/i)

    const matchMin =
      texto.match(/(\d+)\s*m/i)

    let total = 0

    if (matchHoras) {
      total += Number(matchHoras[1]) * 60
    }

    if (matchMin) {
      total += Number(matchMin[1])
    }

    if (!matchHoras && !matchMin) {

      const numero = parseFloat(
        texto.replace(',', '.')
      )

      if (!isNaN(numero)) {
        total = numero
      }
    }

    return total
  }

  const totalPacientes =
    filteredData.length

  const unidadesCount =
    new Set(
      filteredData.map(item =>
        item.NM_LOCAL ||
        item.UNIDADE ||
        item.Local
      )
    ).size

  const medicosCount =
    new Set(
      filteredData.map(item =>
        item.NM_MEDICO ||
        item.MEDICO
      )
    ).size

  const tempos = filteredData.map(item =>
    extrairMinutos(
      item.TEMPO_ESPERA ||
      item.TEMPO ||
      item.ESPERA
    )
  )

  const tempoMedio =
    tempos.length
      ? tempos.reduce((a, b) => a + b, 0) / tempos.length
      : 0

  const maiorTempo =
    tempos.length
      ? Math.max(...tempos)
      : 0

  const statusMap = {}

  filteredData.forEach(item => {

    const status =
      item.STATUS ||
      item.Status ||
      'SEM STATUS'

    statusMap[status] =
      (statusMap[status] || 0) + 1
  })

  const statusData =
    Object.entries(statusMap)
      .map(([name, value]) => ({
        name,
        value
      }))

  const rankingMap = {}

  filteredData.forEach(item => {

    const unidade =
      item.NM_LOCAL ||
      item.UNIDADE ||
      item.Local ||
      'SEM UNIDADE'

    const status =
      String(
        item.STATUS ||
        item.Status ||
        ''
      ).toUpperCase()

    if (!rankingMap[unidade]) {

      rankingMap[unidade] = {
        name: unidade,
        value: 0,
        critico: 0
      }
    }

    rankingMap[unidade].value += 1

    if (
      status.includes('CRITICO') ||
      status.includes('CRÍTICO')
    ) {
      rankingMap[unidade].critico += 1
    }
  })

  const rankingData =
    Object.values(rankingMap)
      .sort((a, b) => b.value - a.value)
      .slice(0, 10)

  const impactoFila = [...filteredData]
    .reduce((acc, item) => {

      const unidade =
        item.NM_LOCAL ||
        item.UNIDADE ||
        item.Local ||
        'SEM UNIDADE'

      const status =
        String(
          item.STATUS ||
          item.Status ||
          ''
        ).toUpperCase()

      if (status.includes('ATRASO')) {

        const existente =
          acc.find(
            x => x.unidade === unidade
          )

        if (existente) {

          existente.pacientes += 1

          existente.medicos.add(
            item.NM_MEDICO ||
            'SEM MÉDICO'
          )

        } else {

          acc.push({
            unidade,
            pacientes: 1,
            medicos: new Set([
              item.NM_MEDICO ||
              'SEM MÉDICO'
            ])
          })
        }
      }

      return acc

    }, [])

    .map(item => ({
      unidade: item.unidade,
      pacientes: item.pacientes,
      medicos: item.medicos.size
    }))

    .sort((a, b) => b.medicos - a.medicos)

    .slice(0, 10)

  return (

    <div style={styles.container}>

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
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFile}
            style={{ display: 'none' }}
          />

        </label>

      </div>

      <div style={styles.filterContainer}>

        <select
          value={selectedUnit}
          onChange={(e) =>
            setSelectedUnit(e.target.value)
          }
          style={styles.select}
        >

          {unidades.map((item, index) => (

            <option
              key={index}
              value={item}
            >
              {item}
            </option>

          ))}

        </select>

        <button
          style={styles.resetButton}
          onClick={() =>
            setSelectedUnit('TODAS')
          }
        >
          Redefinir
        </button>

      </div>

      <div style={styles.cards}>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Total Pacientes
          </h3>

          <h2 style={styles.cardValue}>
            {totalPacientes}
          </h2>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Unidades
          </h3>

          <h2 style={styles.cardValue}>
            {unidadesCount}
          </h2>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Médicos
          </h3>

          <h2 style={styles.cardValue}>
            {medicosCount}
          </h2>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Tempo Médio Espera
          </h3>

          <h2 style={styles.cardValue}>
            {(tempoMedio / 60).toFixed(1)}h
          </h2>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>
            Maior Tempo de Espera
          </h3>

          <h2 style={styles.cardValue}>
            {(maiorTempo / 60).toFixed(1)}h
          </h2>
        </div>

      </div>

      <div style={styles.chartGrid}>

        <div style={styles.chartCard}>

          <h2 style={styles.chartTitle}>
            Ranking Crítico de Pacientes Aguardando
          </h2>

          <ResponsiveContainer
            width="100%"
            height={420}
          >

            <BarChart data={rankingData}>

              <CartesianGrid
                strokeDasharray="3 3"
                stroke="#6d28d9"
              />

              <XAxis
                dataKey="name"
                stroke="#ddd6fe"
              />

              <YAxis stroke="#ddd6fe" />

              <Tooltip />

              <Legend />

              <Bar
                dataKey="value"
                stackId="a"
                fill="#d946ef"
                radius={[8, 8, 0, 0]}
                name="Pacientes"
              />

              <Bar
                dataKey="critico"
                stackId="a"
                fill="#7e22ce"
                radius={[8, 8, 0, 0]}
                name="Críticos"
              />

            </BarChart>

          </ResponsiveContainer>

        </div>

        <div style={styles.chartCard}>

          <h2 style={styles.chartTitle}>
            Status de Pontos Médicos
          </h2>

          <ResponsiveContainer
            width="100%"
            height={420}
          >

            <FunnelChart>

              <Tooltip />

              <Funnel
                dataKey="value"
                data={statusData}
                isAnimationActive
              >

                <LabelList
                  position="right"
                  fill="#ffffff"
                  stroke="none"
                  dataKey="name"
                />

              </Funnel>

            </FunnelChart>

          </ResponsiveContainer>

        </div>

      </div>

      <div style={styles.tableCard}>

        <h2 style={styles.chartTitle}>
          Impacto na Fila de Espera
        </h2>

        <table style={styles.table}>

          <thead>

            <tr>

              <th style={styles.th}>
                Unidade
              </th>

              <th style={styles.th}>
                Qtd Médicos
              </th>

              <th style={styles.th}>
                Pacientes
              </th>

            </tr>

          </thead>

          <tbody>

            {impactoFila.map((item, index) => (

              <tr key={index}>

                <td style={styles.td}>
                  {item.unidade}
                </td>

                <td style={styles.td}>
                  {item.medicos}
                </td>

                <td style={styles.td}>
                  {item.pacientes}
                </td>

              </tr>

            ))}

          </tbody>

        </table>

      </div>

    </div>
  )
}

const styles = {

  container: {
    minHeight: '100vh',
    background:
      'linear-gradient(135deg,#070b1f,#111133,#1e1b4b)',
    padding: '24px',
    color: '#fff',
    fontFamily: 'Arial'
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '30px',
    flexWrap: 'wrap',
    gap: '20px'
  },

  title: {
    fontSize: '56px',
    marginBottom: '10px'
  },

  subtitle: {
    fontSize: '24px',
    color: '#c4b5fd'
  },

  uploadButton: {
    background:
      'linear-gradient(135deg,#c026d3,#7e22ce)',
    padding: '16px 28px',
    borderRadius: '14px',
    cursor: 'pointer',
    fontWeight: 'bold',
    fontSize: '16px',
    boxShadow:
      '0 0 20px rgba(192,38,211,0.4)'
  },

  filterContainer: {
    display: 'flex',
    gap: '16px',
    marginBottom: '30px'
  },

  select: {
    padding: '14px',
    borderRadius: '10px',
    border: 'none',
    minWidth: '260px',
    fontSize: '16px'
  },

  resetButton: {
    background:
      'linear-gradient(135deg,#38bdf8,#2563eb)',
    border: 'none',
    borderRadius: '10px',
    padding: '14px 24px',
    color: '#fff',
    fontWeight: 'bold',
    cursor: 'pointer'
  },

  cards: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit,minmax(240px,1fr))',
    gap: '20px',
    marginBottom: '30px'
  },

  card: {
    background:
      'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.1)',
    borderRadius: '20px',
    padding: '30px',
    backdropFilter: 'blur(10px)'
  },

  cardTitle: {
    color: '#c4b5fd',
    fontSize: '22px',
    marginBottom: '18px'
  },

  cardValue: {
    fontSize: '52px',
    fontWeight: 'bold'
  },

  chartGrid: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit,minmax(500px,1fr))',
    gap: '24px',
    marginBottom: '30px'
  },

  chartCard: {
    background:
      'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.1)',
    borderRadius: '24px',
    padding: '20px'
  },

  chartTitle: {
    fontSize: '26px',
    marginBottom: '20px'
  },

  tableCard: {
    background:
      'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.1)',
    borderRadius: '24px',
    padding: '20px'
  },

  table: {
    width: '100%',
    borderCollapse: 'collapse'
  },

  th: {
    textAlign: 'left',
    padding: '16px',
    background: '#312e81',
    color: '#fff',
    fontSize: '18px'
  },

  td: {
    padding: '16px',
    borderBottom:
      '1px solid rgba(255,255,255,0.08)',
    fontSize: '16px'
  }
}
