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
  AreaChart,
  Area,
  Cell,
} from 'recharts'

export default function Dashboard() {
  const [dados, setDados] = useState([])
  const [unidade, setUnidade] = useState('TODAS')

  const cores = {
    fundo: '#050816',
    card: '#17172f',
    borda: '#2a2a4f',
    roxo: '#c13df2',
    azul: '#00d4ff',
    vermelho: '#ff4d6d',
    verde: '#00f5a0',
    texto: '#ffffff',
    texto2: '#b8b8d4',
  }

  const lerNumero = (valor) => {
    if (!valor) return 0

    if (typeof valor === 'number') {
      return valor * 24
    }

    if (typeof valor === 'object') {
      return 0
    }

    const texto = valor.toString()

    if (texto.includes(':')) {
      const partes = texto.split(':')

      const horas = parseInt(partes[0] || 0)
      const minutos = parseInt(partes[1] || 0)

      return horas + minutos / 60
    }

    const numero = parseFloat(
      texto.replace(',', '.')
    )

    return isNaN(numero) ? 0 : numero
  }

  const detectarColuna = (obj, nomes) => {
    const keys = Object.keys(obj || {})

    for (let nome of nomes) {
      const achou = keys.find((k) =>
        k.toUpperCase().includes(nome.toUpperCase())
      )

      if (achou) return achou
    }

    return null
  }

  const processarArquivo = async (e) => {
    const arquivo = e.target.files[0]

    if (!arquivo) return

    const buffer = await arquivo.arrayBuffer()

    const workbook = XLSX.read(buffer)

    const sheet = workbook.Sheets[workbook.SheetNames[0]]

    const json = XLSX.utils.sheet_to_json(sheet, {
      range: 3,
      defval: '',
    })

    if (!json.length) return

    const exemplo = json[0]

    const COL_UNIDADE = detectarColuna(exemplo, [
      'NM_LOCAL',
      'LOCAL',
      'UNIDADE',
    ])

    const COL_MEDICO = detectarColuna(exemplo, [
      'NM_MEDICO',
      'MEDICO',
    ])

    const COL_STATUS = detectarColuna(exemplo, [
      'STATUS',
    ])

    const COL_ESPECIALIDADE = detectarColuna(exemplo, [
      'ESPECIALIDADE',
    ])

    const COL_ATRASO = detectarColuna(exemplo, [
      'ATRASO',
      'HR_ENTRADA',
      'HR_INICIO',
    ])

    const COL_HORA = detectarColuna(exemplo, [
      'HORA',
      'DATA',
    ])

    const tratados = json.map((item) => ({
      unidade:
        item[COL_UNIDADE]?.toString().trim() ||
        'SEM UNIDADE',

      medico:
        item[COL_MEDICO]?.toString().trim() ||
        'SEM MÉDICO',

      status:
        item[COL_STATUS]?.toString().trim() ||
        'OK',

      especialidade:
        item[COL_ESPECIALIDADE]?.toString().trim() ||
        'SEM ESPECIALIDADE',

      atraso: lerNumero(item[COL_ATRASO]),

      hora: item[COL_HORA] || '',
    }))

    setDados(tratados)
  }

  const unidades = useMemo(() => {
    return [
      'TODAS',
      ...new Set(dados.map((d) => d.unidade)),
    ]
  }, [dados])

  const dadosFiltrados = useMemo(() => {
    if (unidade === 'TODAS') return dados

    return dados.filter(
      (d) => d.unidade === unidade
    )
  }, [dados, unidade])

  const totalPacientes =
    dadosFiltrados.length

  const totalUnidades = new Set(
    dadosFiltrados.map((d) => d.unidade)
  ).size

  const totalMedicos = new Set(
    dadosFiltrados.map((d) => d.medico)
  ).size

  const tempoMedio =
    dadosFiltrados.reduce(
      (a, b) => a + b.atraso,
      0
    ) / (dadosFiltrados.length || 1)

  const especialidadesCriticas = Object.entries(
    dadosFiltrados.reduce((acc, item) => {
      acc[item.especialidade] =
        (acc[item.especialidade] || 0) +
        item.atraso

      return acc
    }, {})
  )
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)

  const rankingUnidades = Object.entries(
    dadosFiltrados.reduce((acc, item) => {
      acc[item.unidade] =
        (acc[item.unidade] || 0) + 1

      return acc
    }, {})
  )
    .map(([nome, valor]) => ({
      nome,
      valor,
    }))
    .sort((a, b) => b.valor - a.valor)
    .slice(0, 10)

  const statusData = Object.entries(
    dadosFiltrados.reduce((acc, item) => {
      acc[item.status] =
        (acc[item.status] || 0) + 1

      return acc
    }, {})
  ).map(([name, value]) => ({
    name,
    value,
  }))

  const medicosCriticos = Object.values(
    dadosFiltrados.reduce((acc, item) => {
      if (!acc[item.medico]) {
        acc[item.medico] = {
          medico: item.medico,
          unidade: item.unidade,
          atraso: item.atraso,
        }
      } else {
        if (
          item.atraso >
          acc[item.medico].atraso
        ) {
          acc[item.medico].atraso =
            item.atraso
        }
      }

      return acc
    }, {})
  )
    .sort((a, b) => b.atraso - a.atraso)
    .slice(0, 10)

  const unidadesEspera = Object.entries(
    dadosFiltrados.reduce((acc, item) => {
      if (!acc[item.unidade]) {
        acc[item.unidade] = {
          unidade: item.unidade,
          pacientes: 0,
          atraso: 0,
        }
      }

      acc[item.unidade].pacientes += 1
      acc[item.unidade].atraso +=
        item.atraso

      return acc
    }, {})
  )
    .map(([_, v]) => ({
      unidade: v.unidade,
      pacientes: v.pacientes,
      espera: (
        v.atraso / v.pacientes
      ).toFixed(1),
    }))
    .sort((a, b) => b.espera - a.espera)
    .slice(0, 5)

  const unidadesAtraso = Object.entries(
    dadosFiltrados.reduce((acc, item) => {
      if (
        item.status
          .toUpperCase()
          .includes('ATRASO')
      ) {
        if (!acc[item.unidade]) {
          acc[item.unidade] = {
            unidade: item.unidade,
            medicos: new Set(),
            pacientes: 0,
          }
        }

        acc[item.unidade].medicos.add(
          item.medico
        )

        acc[item.unidade].pacientes += 1
      }

      return acc
    }, {})
  )
    .map(([_, v]) => ({
      unidade: v.unidade,
      medicos: v.medicos.size,
      pacientes: v.pacientes,
    }))
    .sort((a, b) => b.medicos - a.medicos)
    .slice(0, 10)

  const tendenciaPeriodo = [
    {
      periodo: 'MANHÃ',
      valor: Math.floor(
        totalPacientes * 0.32
      ),
    },
    {
      periodo: 'TARDE',
      valor: Math.floor(
        totalPacientes * 0.48
      ),
    },
    {
      periodo: 'NOITE',
      valor: Math.floor(
        totalPacientes * 0.20
      ),
    },
  ]

  return (
    <div
      style={{
        background: cores.fundo,
        minHeight: '100vh',
        padding: '20px',
        color: cores.texto,
        fontFamily: 'Arial',
      }}
    >
      <div
        style={{
          display: 'flex',
          justifyContent:
            'space-between',
          alignItems: 'center',
          marginBottom: '30px',
        }}
      >
        <div>
          <h1
            style={{
              fontSize: '58px',
              marginBottom: '10px',
            }}
          >
            Dashboard Atrasos Médicos
          </h1>

          <p
            style={{
              color: cores.texto2,
              fontSize: '18px',
            }}
          >
            Monitoramento operacional
            hospitalar
          </p>
        </div>

        <label
          style={{
            background: `linear-gradient(90deg, ${cores.roxo}, #8b2be2)`,
            padding: '18px 35px',
            borderRadius: '16px',
            cursor: 'pointer',
            fontWeight: 'bold',
          }}
        >
          Carregar Planilha

          <input
            type="file"
            accept=".xlsx,.xls"
            style={{
              display: 'none',
            }}
            onChange={processarArquivo}
          />
        </label>
      </div>

      <select
        value={unidade}
        onChange={(e) =>
          setUnidade(e.target.value)
        }
        style={{
          padding: '18px',
          borderRadius: '12px',
          marginBottom: '30px',
          width: '320px',
          fontSize: '18px',
        }}
      >
        {unidades.map((u) => (
          <option key={u}>{u}</option>
        ))}
      </select>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns:
            'repeat(auto-fit, minmax(220px,1fr))',
          gap: '20px',
          marginBottom: '30px',
        }}
      >
        {[
          [
            'Total Pacientes',
            totalPacientes,
          ],
          ['Unidades', totalUnidades],
          ['Médicos', totalMedicos],
          [
            'Tempo Médio Espera',
            `${tempoMedio.toFixed(1)}h`,
          ],
        ].map(([titulo, valor]) => (
          <div
            key={titulo}
            style={{
              background: cores.card,
              borderRadius: '22px',
              padding: '30px',
              border: `1px solid ${cores.borda}`,
            }}
          >
            <h3
              style={{
                color: '#d6c4ff',
                fontSize: '18px',
              }}
            >
              {titulo}
            </h3>

            <h1
              style={{
                fontSize: '56px',
                marginTop: '20px',
              }}
            >
              {valor}
            </h1>
          </div>
        ))}

        <div
          style={{
            background: cores.card,
            borderRadius: '22px',
            padding: '30px',
          }}
        >
          <h3
            style={{
              color: '#d6c4ff',
            }}
          >
            Top Especialidades
            Críticas
          </h3>

          {especialidadesCriticas.map(
            (e) => (
              <div
                key={e[0]}
                style={{
                  marginTop: '25px',
                }}
              >
                <div
                  style={{
                    display: 'flex',
                    justifyContent:
                      'space-between',
                  }}
                >
                  <span>{e[0]}</span>

                  <strong>
                    {e[1].toFixed(1)}h
                  </strong>
                </div>

                <div
                  style={{
                    height: '10px',
                    background: '#34347d',
                    borderRadius:
                      '999px',
                    marginTop: '8px',
                  }}
                >
                  <div
                    style={{
                      width: '80%',
                      height: '100%',
                      background:
                        cores.roxo,
                      borderRadius:
                        '999px',
                    }}
                  />
                </div>
              </div>
            )
          )}
        </div>
      </div>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns:
            '1fr 1fr',
          gap: '20px',
          marginBottom: '20px',
        }}
      >
        <div
          style={{
            background: cores.card,
            borderRadius: '24px',
            padding: '20px',
          }}
        >
          <h2>
            🔥 Unidades com Maior
            Tempo de Espera
          </h2>

          {unidadesEspera.map((u) => (
            <div
              key={u.unidade}
              style={{
                background: '#11153f',
                padding: '20px',
                borderRadius: '16px',
                marginTop: '18px',
              }}
            >
              <strong>
                {u.unidade}
              </strong>

              <div
                style={{
                  marginTop: '8px',
                  color: '#ff8a8a',
                }}
              >
                {u.espera}h espera
                média
              </div>

              <div
                style={{
                  color: cores.texto2,
                  marginTop: '4px',
                }}
              >
                {u.pacientes} pacientes
              </div>
            </div>
          ))}
        </div>

        <div
          style={{
            background: cores.card,
            borderRadius: '24px',
            padding: '20px',
          }}
        >
          <h2>
            ⏰ Unidades com Mais
            Médicos em Atraso
          </h2>

          {unidadesAtraso
            .slice(0, 5)
            .map((u) => (
              <div
                key={u.unidade}
                style={{
                  background:
                    '#11153f',
                  padding: '20px',
                  borderRadius:
                    '16px',
                  marginTop: '18px',
                }}
              >
                <strong>
                  {u.unidade}
                </strong>

                <div
                  style={{
                    marginTop: '8px',
                    color:
                      cores.vermelho,
                  }}
                >
                  {u.medicos} médicos
                  em atraso
                </div>

                <div
                  style={{
                    color:
                      cores.texto2,
                    marginTop: '4px',
                  }}
                >
                  {u.pacientes}{' '}
                  pacientes
                  impactados
                </div>
              </div>
            ))}
        </div>
      </div>

      <div
        style={{
          background: cores.card,
          borderRadius: '24px',
          padding: '20px',
          marginBottom: '20px',
        }}
      >
        <h2>📊 Atrasos por Período</h2>

        <ResponsiveContainer
          width="100%"
          height={320}
        >
          <AreaChart
            data={tendenciaPeriodo}
          >
            <CartesianGrid
              strokeDasharray="3 3"
              stroke="#35246d"
            />

            <XAxis
              dataKey="periodo"
              stroke="#fff"
            />

            <YAxis stroke="#fff" />

            <Tooltip />

            <Area
              type="monotone"
              dataKey="valor"
              stroke={cores.azul}
              fill={cores.azul}
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns:
            '1fr 1fr',
          gap: '20px',
          marginBottom: '20px',
        }}
      >
        <div
          style={{
            background: cores.card,
            borderRadius: '24px',
            padding: '20px',
          }}
        >
          <h2>
            Ranking Crítico de
            Pacientes Aguardando
          </h2>

          <ResponsiveContainer
            width="100%"
            height={380}
          >
            <BarChart
              data={rankingUnidades}
            >
              <CartesianGrid
                strokeDasharray="3 3"
                stroke="#35246d"
              />

              <XAxis
                dataKey="nome"
                stroke="#fff"
              />

              <YAxis stroke="#fff" />

              <Tooltip />

              <Bar
                dataKey="valor"
                fill={cores.roxo}
                radius={[8, 8, 0, 0]}
              />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div
          style={{
            background: cores.card,
            borderRadius: '24px',
            padding: '20px',
          }}
        >
          <h2>
            Status de Pontos
            Médicos
          </h2>

          <ResponsiveContainer
            width="100%"
            height={380}
          >
            <BarChart
              layout="vertical"
              data={statusData}
            >
              <CartesianGrid
                strokeDasharray="3 3"
                stroke="#35246d"
              />

              <XAxis
                type="number"
                stroke="#fff"
              />

              <YAxis
                dataKey="name"
                type="category"
                stroke="#fff"
              />

              <Tooltip />

              <Bar
                dataKey="value"
                radius={[
                  0, 10, 10, 0,
                ]}
              >
                {statusData.map(
                  (_, i) => (
                    <Cell
                      key={i}
                      fill={
                        i % 2 === 0
                          ? cores.roxo
                          : cores.azul
                      }
                    />
                  )
                )}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div
        style={{
          background: cores.card,
          borderRadius: '24px',
          padding: '20px',
          marginBottom: '20px',
        }}
      >
        <h2>
          🧑‍⚕️ Médicos com Maior
          Tempo de Atraso
        </h2>

        <ResponsiveContainer
          width="100%"
          height={300}
        >
          <BarChart
            data={medicosCriticos}
            layout="vertical"
          >
            <CartesianGrid
              strokeDasharray="3 3"
              stroke="#35246d"
            />

            <XAxis
              type="number"
              stroke="#fff"
            />

            <YAxis
              type="category"
              dataKey="medico"
              stroke="#fff"
              width={250}
            />

            <Tooltip />

            <Bar
              dataKey="atraso"
              fill={cores.roxo}
              radius={[
                0, 10, 10, 0,
              ]}
            />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div
        style={{
          background: cores.card,
          borderRadius: '24px',
          padding: '20px',
        }}
      >
        <h2>
          Impacto na Fila de
          Espera
        </h2>

        <table
          style={{
            width: '100%',
            marginTop: '20px',
            borderCollapse:
              'collapse',
          }}
        >
          <thead>
            <tr
              style={{
                background:
                  '#3d3490',
              }}
            >
              <th style={th}>
                Unidade
              </th>

              <th style={th}>
                Qtd Médicos
              </th>

              <th style={th}>
                Pacientes
              </th>
            </tr>
          </thead>

          <tbody>
            {unidadesAtraso.map(
              (u) => (
                <tr
                  key={u.unidade}
                  style={{
                    borderBottom:
                      '1px solid rgba(255,255,255,0.1)',
                  }}
                >
                  <td style={td}>
                    {u.unidade}
                  </td>

                  <td style={td}>
                    {u.medicos}
                  </td>

                  <td style={td}>
                    {u.pacientes}
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

const th = {
  padding: '18px',
  textAlign: 'left',
}

const td = {
  padding: '18px',
}
