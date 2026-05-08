'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

import {
  ResponsiveContainer,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip
} from 'recharts'

export default function Home() {

  const [data, setData] = useState([])
  const [selectedUnit, setSelectedUnit] = useState('TODAS')

  function handleFile(event) {

    const file = event.target.files[0]

    if (!file) return

    const reader = new FileReader()

    reader.onload = (e) => {

      const workbook = XLSX.read(e.target.result, {
        type: 'binary'
      })

      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true
      })

      const headers = rows[3]

      const body = rows
        .slice(4)
        .filter(row => row.length > 0)

      const json = body.map(row => {

        const obj = {}

        headers.forEach((header, index) => {
          obj[header] = row[index]
        })

        return obj
      })

      setData(json)
    }

    reader.readAsBinaryString(file)
  }

  function tempoParaMinutos(valor) {

    if (!valor) return 0

    if (typeof valor === 'number') {
      return valor * 24 * 60
    }

    const texto = String(valor)

    const partes = texto.split(':')

    if (partes.length >= 2) {

      const horas = Number(partes[0])
      const minutos = Number(partes[1])

      return (horas * 60) + minutos
    }

    return 0
  }

  const filteredData = useMemo(() => {

    if (selectedUnit === 'TODAS') {
      return data
    }

    return data.filter(item =>
      item['NM_LOCAL'] === selectedUnit
    )

  }, [data, selectedUnit])

  const unidades = useMemo(() => {

    return [
      'TODAS',
      ...new Set(
        data
          .map(item => item['NM_LOCAL'])
          .filter(Boolean)
      )
    ]

  }, [data])

  const totalPacientes = filteredData.length

  const totalUnidades = new Set(
    filteredData.map(item => item['NM_LOCAL'])
  ).size

  const totalMedicos = new Set(
    filteredData.map(item => item['NM_MEDICO'])
  ).size

  const tempos = filteredData.map(item =>
    tempoParaMinutos(item['TEMPO_DE_ESPERA'])
  )

  const tempoMedio = tempos.length
    ? tempos.reduce((a, b) => a + b, 0) / tempos.length
    : 0

  const especialidades = [...filteredData]

    .reduce((acc, item) => {

      const nome =
        item['DS_ESPECIALIDADE'] ||
        'SEM ESPECIALIDADE'

      const tempo = tempoParaMinutos(
        item['TEMPO DE ATRASO']
      )

      const existente =
        acc.find(x => x.nome === nome)

      if (existente) {

        existente.tempo += tempo
        existente.total += 1

      } else {

        acc.push({
          nome,
          tempo,
          total: 1
        })
      }

      return acc

    }, [])

    .map(item => ({
      nome: item.nome,
      media: Number(
        (
          (item.tempo / item.total) / 60
        ).toFixed(1)
      )
    }))

    .sort((a, b) => b.media - a.media)

    .slice(0, 3)

  const rankingUnidades = [...filteredData]

    .reduce((acc, item) => {

      const unidade =
        item['NM_LOCAL'] || 'SEM UNIDADE'

      const qtd = Number(
        item[' QT_PACIENTES_AGUARDANDO'] || 0
      )

      const existente =
        acc.find(x => x.name === unidade)

      if (existente) {
        existente.value += qtd
      }

      else {

        acc.push({
          name: unidade,
          value: qtd
        })
      }

      return acc

    }, [])

    .sort((a, b) => b.value - a.value)

    .slice(0, 10)

  const statusData = [...filteredData]

    .reduce((acc, item) => {

      const status =
        item['STATUS'] || 'SEM STATUS'

      const existente =
        acc.find(x => x.name === status)

      if (existente) {
        existente.value += 1
      }

      else {

        acc.push({
          name: status,
          value: 1
        })
      }

      return acc

    }, [])

  const medicosAtraso = [...filteredData]

    .reduce((acc, item) => {

      const medico =
        item['NM_MEDICO'] || 'SEM MÉDICO'

      const unidade =
        item['NM_LOCAL'] || 'SEM UNIDADE'

      const tempo = tempoParaMinutos(
        item['TEMPO DE ATRASO']
      )

      const existente =
        acc.find(x => x.medico === medico)

      if (existente) {

        if (tempo > existente.tempo) {
          existente.tempo = tempo
        }

      } else {

        acc.push({
          medico,
          unidade,
          tempo
        })
      }

      return acc

    }, [])

    .sort((a, b) => b.tempo - a.tempo)

    .slice(0, 10)

    .map(item => ({
      ...item,
      horas: Number(
        (item.tempo / 60).toFixed(1)
      )
    }))

  const impactoFila = [...filteredData]

    .filter(item => {

      const status = String(
        item['STATUS']
      ).toUpperCase()

      return status.includes('ATRASO')
    })

    .reduce((acc, item) => {

      const unidade =
        item['NM_LOCAL'] || 'SEM UNIDADE'

      const pacientes = Number(
        item[' QT_PACIENTES_AGUARDANDO'] || 0
      )

      const medico =
        item['NM_MEDICO'] || 'SEM MÉDICO'

      const existente =
        acc.find(x => x.unidade === unidade)

      if (existente) {

        existente.pacientes += pacientes
        existente.medicos.add(medico)

      } else {

        acc.push({
          unidade,
          pacientes,
          medicos: new Set([medico])
        })
      }

      return acc

    }, [])

    .map(item => ({
      unidade: item.unidade,
      pacientes: item.pacientes,
      qtdMedicos: item.medicos.size
    }))

    .sort((a, b) => b.pacientes - a.pacientes)

    .slice(0, 10)

  const COLORS = [
    '#d946ef',
    '#a855f7',
    '#9333ea',
    '#7e22ce',
    '#6d28d9'
  ]

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
            type='file'
            accept='.xlsx,.xls'
            onChange={handleFile}
            style={{ display: 'none' }}
          />

        </label>

      </div>

      <div style={styles.filterArea}>

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

      </div>

      <div style={styles.cards}>

        <div style={styles.card}>
          <h3>Total Pacientes</h3>
          <h1>{totalPacientes}</h1>
        </div>

        <div style={styles.card}>
          <h3>Unidades</h3>
          <h1>{totalUnidades}</h1>
        </div>

        <div style={styles.card}>
          <h3>Médicos</h3>
          <h1>{totalMedicos}</h1>
        </div>

        <div style={styles.card}>
          <h3>Tempo Médio Espera</h3>
          <h1>
            {(tempoMedio / 60).toFixed(1)}h
          </h1>
        </div>

        <div style={styles.card}>

          <h3>Top Especialidades Críticas</h3>

          {especialidades.map((item, index) => (

            <div
              key={index}
              style={{ marginTop: '20px' }}
            >

              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between'
                }}
              >

                <span>{item.nome}</span>

                <strong>
                  {item.media}h
                </strong>

              </div>

              <div style={styles.progress}>
                <div
                  style={{
                    ...styles.progressBar,
                    width: `${item.media * 8}%`
                  }}
                />
              </div>

            </div>

          ))}

        </div>

      </div>

      <div style={styles.grid}>

        <div style={styles.chartCard}>

          <h2>
            Ranking Crítico de Pacientes Aguardando
          </h2>

          <ResponsiveContainer
            width='100%'
            height={400}
          >

            <BarChart data={rankingUnidades}>

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
                fill='#c026d3'
                radius={[10, 10, 0, 0]}
              />

            </BarChart>

          </ResponsiveContainer>

        </div>

        <div style={styles.chartCard}>

          <h2>Status de Pontos Médicos</h2>

          <ResponsiveContainer
            width='100%'
            height={400}
          >

            <PieChart>

              <Pie
                data={statusData}
                dataKey='value'
                nameKey='name'
                outerRadius={130}
                label
              >

                {statusData.map((entry, index) => (

                  <Cell
                    key={index}
                    fill={
                      COLORS[
                        index % COLORS.length
                      ]
                    }
                  />

                ))}

              </Pie>

              <Tooltip />

            </PieChart>

          </ResponsiveContainer>

        </div>

      </div>

      <div style={styles.chartCard}>

        <h2>
          Médicos com Maior Tempo de Atraso
        </h2>

        <ResponsiveContainer
          width='100%'
          height={450}
        >

          <BarChart
            data={medicosAtraso}
            layout='vertical'
            margin={{ left: 200 }}
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
              dataKey='medico'
              type='category'
              width={250}
              stroke='#ddd6fe'
            />

            <Tooltip />

            <Bar
              dataKey='horas'
              fill='#d946ef'
              radius={[0, 10, 10, 0]}
            />

          </BarChart>

        </ResponsiveContainer>

      </div>

      <div style={styles.chartCard}>

        <h2>
          Impacto na Fila de Espera
        </h2>

        <table style={styles.table}>

          <thead>

            <tr>
              <th>Unidade</th>
              <th>Qtd Médicos</th>
              <th>Pacientes</th>
            </tr>

          </thead>

          <tbody>

            {impactoFila.map((item, index) => (

              <tr key={index}>
                <td>{item.unidade}</td>
                <td>{item.qtdMedicos}</td>
                <td>{item.pacientes}</td>
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
      'linear-gradient(135deg,#050816,#15153b,#221b5e)',
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
    fontSize: '52px'
  },

  subtitle: {
    color: '#c4b5fd',
    marginTop: '10px',
    fontSize: '22px'
  },

  uploadButton: {
    background:
      'linear-gradient(135deg,#d946ef,#7e22ce)',
    padding: '16px 26px',
    borderRadius: '14px',
    cursor: 'pointer',
    fontWeight: 'bold'
  },

  filterArea: {
    marginBottom: '30px'
  },

  select: {
    padding: '14px',
    borderRadius: '10px',
    minWidth: '250px',
    border: 'none'
  },

  cards: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit,minmax(240px,1fr))',
    gap: '20px',
    marginBottom: '30px'
  },

  card: {
    background: 'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.08)',
    borderRadius: '24px',
    padding: '30px'
  },

  progress: {
    height: '10px',
    background: '#312e81',
    borderRadius: '999px',
    marginTop: '8px'
  },

  progressBar: {
    height: '10px',
    background:
      'linear-gradient(90deg,#d946ef,#9333ea)',
    borderRadius: '999px'
  },

  grid: {
    display: 'grid',
    gridTemplateColumns:
      'repeat(auto-fit,minmax(500px,1fr))',
    gap: '20px',
    marginBottom: '30px'
  },

  chartCard: {
    background: 'rgba(255,255,255,0.05)',
    border:
      '1px solid rgba(255,255,255,0.08)',
    borderRadius: '24px',
    padding: '20px',
    marginBottom: '30px'
  },

  table: {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '20px'
  }
}
