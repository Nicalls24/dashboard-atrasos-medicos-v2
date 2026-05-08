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
