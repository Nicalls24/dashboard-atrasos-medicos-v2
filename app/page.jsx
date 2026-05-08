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
  AreaChart,
  Area,
} from 'recharts'

export default function Home() {
  const [dados, setDados] = useState([])
  const [unidadeFiltro, setUnidadeFiltro] =
    useState('TODAS')

  const cores = {
    fundo: '#05051f',
    card: '#1a1a40',
    card2: '#232356',
    roxo: '#c13df0',
    roxo2: '#8f2fff',
    azul: '#00d2ff',
    texto: '#ffffff',
    subtitulo: '#d6c4ff',
    grade: '#5d2a9d',
  }

  const handleUpload = async (e) => {
    const arquivo = e.target.files[0]

    if (!arquivo) return

    const buffer =
      await arquivo.arrayBuffer()

    const workbook = XLSX.read(buffer, {
      type: 'buffer',
    })

    const sheet =
      workbook.Sheets[
        workbook.SheetNames[0]
      ]

    const json =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          range: 3,
          defval: '',
        }
      )

    setDados(json)
  }

  const identificarColuna = (
    item,
    nomes
  ) => {
    const chaves = Object.keys(item || {})

    return (
      chaves.find((c) =>
        nomes.some((n) =>
          c
            .toLowerCase()
            .includes(n.toLowerCase())
        )
      ) || ''
    )
  }

  const colunas = useMemo(() => {
    if (!dados.length) return {}

    const item = dados[0]

    return {
      unidade: identificarColuna(item, [
        'NM_LOCAL',
        'UNIDADE',
      ]),
      medico: identificarColuna(item, [
        'MEDICO',
        'NM_MEDICO',
      ]),
      especialidade:
        identificarColuna(item, [
          'ESPECIALIDADE',
        ]),
      status: identificarColuna(item, [
        'STATUS',
      ]),
      espera: identificarColuna(item, [
        'ESPERA',
        'ATRASO',
        'TEMPO',
      ]),
    }
  }, [dados])

  const dadosFiltrados = useMemo(() => {
    if (
      unidadeFiltro === 'TODAS'
    )
      return dados

    return dados.filter(
      (d) =>
        String(
          d[colunas.unidade]
        ).trim() === unidadeFiltro
    )
  }, [
    dados,
    unidadeFiltro,
    colunas,
  ])

  const unidades = useMemo(() => {
    const lista = dados
      .map(
        (d) =>
          d[colunas.unidade]
      )
      .filter(Boolean)

    return [...new Set(lista)]
  }, [dados, colunas])

  const parseHoras = (valor) => {
    if (!valor) return 0

    const texto = String(
      valor
    ).replace(',', '.')

    const numero = parseFloat(texto)

    if (!isNaN(numero))
      return numero

    return 0
  }

  const totalPacientes =
    dadosFiltrados.length

  const totalUnidades =
    new Set(
      dadosFiltrados.map(
        (d) =>
          d[colunas.unidade]
      )
    ).size

  const totalMedicos =
    new Set(
      dadosFiltrados.map(
        (d) =>
          d[colunas.medico]
      )
    ).size

  const tempoMedio =
    dadosFiltrados.reduce(
      (acc, d) =>
        acc +
        parseHoras(
          d[colunas.espera]
        ),
      0
    ) /
      totalPacientes || 0

  const rankingPacientes =
    Object.entries(
      dadosFiltrados.reduce(
        (acc, d) => {
          const unidade =
            d[
              colunas.unidade
            ] || 'SEM UNIDADE'

          acc[unidade] =
            (acc[unidade] || 0) +
            1

          return acc
        },
        {}
      )
    )
      .map(([nome, total]) => ({
        nome,
        total,
      }))
      .sort(
        (a, b) =>
          b.total - a.total
      )
      .slice(0, 10)

  const statusMedicos =
    Object.entries(
      dadosFiltrados.reduce(
        (acc, d) => {
          const status =
            d[
              colunas.status
            ] || 'OK'

          acc[status] =
            (acc[status] || 0) +
            1

          return acc
        },
        {}
      )
    ).map(([nome, total]) => ({
      nome,
      total,
    }))

  const medicosAtraso =
    Object.values(
      dadosFiltrados.reduce(
        (acc, d) => {
          const medico =
            d[
              colunas.medico
            ] || 'SEM MÉDICO'

          const unidade =
            d[
              colunas.unidade
            ] || 'SEM UNIDADE'

          const atraso =
            parseHoras(
              d[
                colunas.espera
              ]
            )

          if (
            !acc[medico] ||
            atraso >
              acc[medico]
                .atraso
          ) {
            acc[medico] = {
              medico,
              unidade,
              atraso,
            }
          }

          return acc
        },
        {}
      )
    )
      .sort(
        (a, b) =>
          b.atraso - a.atraso
      )
      .slice(0, 10)

  const impactoFila =
    Object.entries(
      dadosFiltrados.reduce(
        (acc, d) => {
          const unidade =
            d[
              colunas.unidade
            ] || 'SEM UNIDADE'

          const status =
            String(
              d[
                colunas.status
              ] || ''
            ).toUpperCase()

          if (
            !status.includes(
              'ATRASO'
            )
          )
            return acc

          if (!acc[unidade]) {
            acc[unidade] = {
              unidade,
              qtdMedicos: 0,
              pacientes: 0,
            }
          }

          acc[
            unidade
          ].qtdMedicos += 1

          acc[
            unidade
          ].pacientes += 1

          return acc
        },
        {}
      )
    )
      .map(([_, v]) => v)
      .sort(
        (a, b) =>
          b.qtdMedicos -
          a.qtdMedicos
      )
      .slice(0, 10)

  const periodos = [
    {
      periodo: 'MANHÃ',
      valor: Math.floor(
        totalPacientes * 0.35
      ),
    },
    {
      periodo: 'TARDE',
      valor: Math.floor(
        totalPacientes * 0.45
      ),
    },
    {
      periodo: 'NOITE',
      valor: Math.floor(
        totalPacientes * 0.2
      ),
    },
  ]

  return (
    <div
      style={{
        background: cores.fundo,
        minHeight: '100vh',
        padding: '40px',
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
        }}
      >
        <div>
          <h1
            style={{
              fontSize: '64px',
              marginBottom: '10px',
            }}
          >
            Dashboard Atrasos
            Médicos
          </h1>

          <p
            style={{
              color:
                cores.subtitulo,
              fontSize: '20px',
            }}
          >
            Monitoramento
            operacional
            hospitalar
          </p>
        </div>

        <label
          style={{
            background:
              'linear-gradient(90deg,#d946ef,#7e22ce)',
            padding:
              '18px 36px',
            borderRadius:
              '18px',
            cursor: 'pointer',
            fontWeight: 'bold',
          }}
        >
          Carregar Planilha

          <input
            type='file'
            accept='.xlsx,.xls'
            style={{
              display: 'none',
            }}
            onChange={
              handleUpload
            }
          />
        </label>
      </div>

      <div
        style={{
          marginTop: '40px',
          marginBottom: '40px',
        }}
      >
        <select
          value={unidadeFiltro}
          onChange={(e) =>
            setUnidadeFiltro(
              e.target.value
            )
          }
          style={{
            padding:
              '20px 24px',
            width: '320px',
            borderRadius:
              '16px',
            border: 'none',
            fontSize: '22px',
          }}
        >
          <option>
            TODAS
          </option>

          {unidades.map((u) => (
            <option
              key={u}
            >
              {u}
            </option>
          ))}
        </select>
      </div>
