export default function Home() {
  const hospitais = [
    { nome: 'Hospital Fortaleza', valor: 92 },
    { nome: 'Hospital Sul', valor: 75 },
    { nome: 'Hospital Norte', valor: 63 },
    { nome: 'Hospital Central', valor: 55 },
    { nome: 'Hospital Vida', valor: 40 },
  ]

  return (
    <div style={styles.container}>
      <h1 style={styles.title}>Dashboard Atrasos Médicos</h1>

      <div style={styles.cards}>
        <div style={styles.card}>
          <p style={styles.cardTitle}>Pacientes +30h</p>
          <h2 style={styles.cardNumber}>128</h2>
        </div>

        <div style={styles.card}>
          <p style={styles.cardTitle}>Visitas Pendentes</p>
          <h2 style={styles.cardNumber}>42</h2>
        </div>

        <div style={styles.card}>
          <p style={styles.cardTitle}>Média de Permanência</p>
          <h2 style={styles.cardNumber}>36h</h2>
        </div>

        <div style={styles.card}>
          <p style={styles.cardTitle}>Hospitais Críticos</p>
          <h2 style={styles.cardNumber}>7</h2>
        </div>
      </div>

      <div style={styles.content}>
        <div style={styles.chartBox}>
          <h2 style={styles.chartTitle}>
            Ranking de Hospitais
          </h2>

          <div style={styles.chart}>
            {hospitais.map((hospital, index) => (
              <div key={index} style={styles.row}>
                <span style={styles.label}>
                  {hospital.nome}
                </span>

                <div
                  style={{
                    ...styles.bar,
                    width: `${hospital.valor}%`,
                  }}
                />
              </div>
            ))}
          </div>
        </div>

        <div style={styles.statusContainer}>
          <div style={styles.statusBox}>
            <h3>Crítico</h3>
            <h1>18</h1>
          </div>

          <div style={styles.statusBox}>
            <h3>Atenção</h3>
            <h1>27</h1>
          </div>

          <div style={styles.statusBox}>
            <h3>Normal</h3>
            <h1>83</h1>
          </div>
        </div>
      </div>
    </div>
  )
}

const styles = {
  container: {
    minHeight: '100vh',
    background: '#020f2b',
    color: 'white',
    padding: '40px',
    fontFamily: 'Arial',
  },

  title: {
    fontSize: '48px',
    marginBottom: '40px',
    fontWeight: 'bold',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: '20px',
    marginBottom: '40px',
  },

  card: {
    background: '#111c44',
    padding: '25px',
    borderRadius: '20px',
  },

  cardTitle: {
    color: '#9ca3af',
    marginBottom: '10px',
  },

  cardNumber: {
    fontSize: '40px',
    fontWeight: 'bold',
  },

  content: {
    display: 'grid',
    gridTemplateColumns: '2fr 1fr',
    gap: '25px',
  },

  chartBox: {
    background: '#111c44',
    padding: '25px',
    borderRadius: '20px',
  },

  chartTitle: {
    fontSize: '24px',
    marginBottom: '25px',
    fontWeight: 'bold',
  },

  chart: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
  },

  row: {
    display: 'flex',
    alignItems: 'center',
    gap: '15px',
  },

  label: {
    width: '120px',
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
    background: '#111c44',
    padding: '25px',
    borderRadius: '16px',
    textAlign: 'center',
  },
}
