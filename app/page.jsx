export default function Home() {
  return (
    <main
      style={{
        minHeight: '100vh',
        background: '#021028',
        color: 'white',
        padding: '40px',
        fontFamily: 'Arial',
      }}
    >
      <h1
        style={{
          fontSize: '48px',
          marginBottom: '40px',
        }}
      >
        Dashboard Atrasos Médicos
      </h1>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(4, 1fr)',
          gap: '20px',
          marginBottom: '40px',
        }}
      >
        <div style={card}>
          <p>Pacientes +30h</p>
          <h2>128</h2>
        </div>

        <div style={card}>
          <p>Visitas Pendentes</p>
          <h2>42</h2>
        </div>

        <div style={card}>
          <p>Média Permanência</p>
          <h2>36h</h2>
        </div>

        <div style={card}>
          <p>Hospitais Críticos</p>
          <h2>7</h2>
        </div>
      </div>

      <div
        style={{
          background: '#111c44',
          borderRadius: '20px',
          padding: '30px',
        }}
      >
        <h2
          style={{
            marginBottom: '20px',
          }}
        >
          Ranking Hospitais
        </h2>

        <div style={{ marginBottom: '15px' }}>
          <p>Hospital Fortaleza</p>

          <div style={barBackground}>
            <div
              style={{
                ...bar,
                width: '92%',
              }}
            />
          </div>
        </div>

        <div style={{ marginBottom: '15px' }}>
          <p>Hospital Sul</p>

          <div style={barBackground}>
            <div
              style={{
                ...bar,
                width: '75%',
              }}
            />
          </div>
        </div>

        <div style={{ marginBottom: '15px' }}>
          <p>Hospital Norte</p>

          <div style={barBackground}>
            <div
              style={{
                ...bar,
                width: '63%',
              }}
            />
          </div>
        </div>
      </div>
    </main>
  )
}

const card = {
  background: '#111c44',
  padding: '25px',
  borderRadius: '20px',
}

const barBackground = {
  width: '100%',
  height: '14px',
  background: '#1e293b',
  borderRadius: '999px',
}

const bar = {
  height: '14px',
  background: '#38bdf8',
  borderRadius: '999px',
}
