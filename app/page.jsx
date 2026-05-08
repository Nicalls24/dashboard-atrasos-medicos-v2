'use client'

export default function Home() {
  const pacientes = [
    {
      unidade: 'UTI Central',
      pacientes: 14,
      atraso: '38h',
      status: 'Crítico',
    },
    {
      unidade: 'Emergência',
      pacientes: 9,
      atraso: '31h',
      status: 'Alerta',
    },
    {
      unidade: 'Clínica Médica',
      pacientes: 6,
      atraso: '29h',
      status: 'Moderado',
    },
    {
      unidade: 'Internação',
      pacientes: 11,
      atraso: '41h',
      status: 'Crítico',
    },
  ]

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <h1 style={styles.title}>Dashboard Atrasos Médicos</h1>
        <p style={styles.subtitle}>
          Monitoramento de pacientes com permanência acima de 30 horas
        </p>
      </div>

      <div style={styles.cards}>
        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Pacientes +30h</h3>
          <p style={styles.cardValue}>40</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Tempo Médio</h3>
          <p style={styles.cardValue}>34h</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Unidades Críticas</h3>
          <p style={styles.cardValue}>2</p>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Visitas Pendentes</h3>
          <p style={styles.cardValue}>17</p>
        </div>
      </div>

      <div style={styles.tableContainer}>
        <h2 style={styles.tableTitle}>Unidades com Maior Atraso</h2>

        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>Unidade</th>
              <th style={styles.th}>Pacientes</th>
              <th style={styles.th}>Maior Atraso</th>
              <th style={styles.th}>Status</th>
            </tr>
          </thead>

          <tbody>
            {pacientes.map((item, index) => (
              <tr key={index}>
                <td style={styles.td}>{item.unidade}</td>
                <td style={styles.td}>{item.pacientes}</td>
                <td style={styles.td}>{item.atraso}</td>
                <td style={styles.td}>
                  <span
                    style={{
                      ...styles.status,
                      background:
                        item.status === 'Crítico'
                          ? '#ff4d4f'
                          : item.status === 'Alerta'
                          ? '#faad14'
                          : '#52c41a',
                    }}
                  >
                    {item.status}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={styles.footer}>
        Dashboard desenvolvido para monitoramento operacional hospitalar
      </div>
    </div>
  )
}

const styles = {
  page: {
    minHeight: '100vh',
    background: '#061529',
    padding: '40px',
    color: '#fff',
    fontFamily: 'Arial',
  },

  header: {
    marginBottom: '40px',
  },

  title: {
    fontSize: '42px',
    marginBottom: '10px',
  },

  subtitle: {
    fontSize: '18px',
    color: '#9fb3c8',
  },

  cards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: '20px',
    marginBottom: '40px',
  },

  card: {
    background: '#0d223d',
    padding: '25px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  cardTitle: {
    fontSize: '16px',
    color: '#9fb3c8',
    marginBottom: '15px',
  },

  cardValue: {
    fontSize: '38px',
    fontWeight: 'bold',
  },

  tableContainer: {
    background: '#0d223d',
    padding: '25px',
    borderRadius: '18px',
    border: '1px solid #16385f',
  },

  tableTitle: {
    marginBottom: '20px',
    fontSize: '24px',
  },

  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },

  th: {
    textAlign: 'left',
    padding: '16px',
    background: '#102944',
    color: '#8fb3d9',
  },

  td: {
    padding: '16px',
    borderBottom: '1px solid #16385f',
  },

  status: {
    padding: '8px 14px',
    borderRadius: '999px',
    color: '#fff',
    fontWeight: 'bold',
    fontSize: '14px',
  },

  footer: {
    marginTop: '30px',
    color: '#6f8aa6',
    fontSize: '14px',
    textAlign: 'center',
  },
}
