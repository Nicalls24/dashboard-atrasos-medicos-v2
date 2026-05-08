'use client'

import { useState } from 'react'

export default function Home() {
  const [hospital, setHospital] = useState('Todos')

  const cards = [
    {
      titulo: 'TOTAL ATRASOS',
      valor: '148',
      cor: '#22c55e',
    },
    {
      titulo: 'MÉDIA TEMPO',
      valor: '42min',
      cor: '#38bdf8',
    },
    {
      titulo: 'UNIDADES',
      valor: '12',
      cor: '#f97316',
    },
    {
      titulo: 'CRÍTICOS',
      valor: '18',
      cor: '#ef4444',
    },
  ]

  const hospitais = [
    'Todos',
    'Hospital Central',
    'Hospital Norte',
    'Hospital Sul',
    'UPA Fortaleza',
  ]

  const ranking = [
    {
      unidade: 'Hospital Central',
      atrasos: 42,
    },
    {
      unidade: 'Hospital Norte',
      atrasos: 35,
    },
    {
      unidade: 'Hospital Sul',
      atrasos: 28,
    },
    {
      unidade: 'UPA Fortaleza',
      atrasos: 19,
    },
  ]

  return (
    <main style={styles.container}>
      <div style={styles.header}>
        <div>
          <h1 style={styles.title}>Dashboard Atrasos Médicos</h1>
          <p style={styles.subtitle}>
            Monitoramento inteligente de atrasos hospitalares
          </p>
        </div>

        <select
          style={styles.select}
          value={hospital}
          onChange={(e) => setHospital(e.target.value)}
        >
          {hospitais.map((item) => (
            <option key={item}>{item}</option>
          ))}
        </select>
      </div>

      <div style={styles.grid}>
        {cards.map((card) => (
          <div key={card.titulo} style={styles.card}>
            <p style={styles.cardTitle}>{card.titulo}</p>

            <h2 style={{ ...styles.cardValue, color: card.cor }}>
              {card.valor}
            </h2>
          </div>
        ))}
      </div>

      <div style={styles.content}>
        <div style={styles.chartCard}>
          <h2 style={styles.sectionTitle}>Ranking de Atrasos</h2>

          {ranking.map((item) => (
            <div key={item.unidade} style={styles.row}>
              <div style={styles.rowTop}>
                <span>{item.unidade}</span>
                <span>{item.atrasos}</span>
              </div>

              <div style={styles.barBackground}>
                <div
                  style={{
                    ...styles.bar,
                    width: `${item.atrasos * 2}%`,
                  }}
                />
              </div>
            </div>
          ))}
        </div>

        <div style={styles.chartCard}>
          <h2 style={styles.sectionTitle}>Status Geral</h2>

          <div style={styles.statusContainer}>
            <div style={styles.statusBox}>
              <h3 style={{ color: '#22c55e' }}>76%</h3>
              <p>Dentro do prazo</p>
            </div>

            <div style={styles.statusBox}>
              <h3 style={{ color: '#ef4444' }}>24%</h3>
              <p>Em atraso</p>
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}

const styles = {
  container: {
    minHeight: '100vh',
    background: '#061326',
    padding: '40px',
    color: 'white',
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '30px',
  },

  title: {
    fontSize: '38px',
    fontWeight: 'bold',
  },

  subtitle: {
    marginTop: '8px',
    color: '#94a3b8',
  },

  select: {
    background: '#0f172a',
    border: '1px solid #1e293b',
    color: 'white',
    padding: '12px',
    borderRadius: '10px',
    minWidth: '220px',
  },

  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: '20px',
    marginBottom: '30px',
  },

  card: {
    background: '#0f172a',
    padding: '25px',
    borderRadius: '18px',
    border: '1px solid #1e293b',
  },

  cardTitle: {
    color: '#94a3b8',
    marginBottom: '15px',
  },

  cardValue: {
    fontSize: '42px',
    fontWeight: 'bold',
  },

  content: {
    display: 'grid',
    gridTemplateColumns: '2fr 1fr',
    gap: '20px',
  },

  chartCard: {
    background: '#0f172a',
    borderRadius: '18px',
    padding: '25px',
    border: '1px solid #1e293b',
  },

  sectionTitle: {
    marginBottom: '25px',
    fontSize: '22px',
  },

  row: {
    marginBottom: '22px',
  },

  rowTop: {
    display: 'flex',
    justifyContent: 'space-between',
    marginBottom: '10px',
  },

  barBackground: {
    height: '14px',
    width: '100%',
    background: '#1e293b',
    borderRadius: '999px',
  },

  bar: {
    height: '14px',
    borderRadius: '999px',
    background: '#38bdf8',
  },

  statusContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '20px',
  },

  statusBox: {
    background: '#111c33',
    padding: '25px',
    borderRadius: '16px',
    textAlign: 'center',
  },
}
