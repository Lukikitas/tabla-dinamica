import { useMemo, useState } from 'react';
import Button from '../components/UI/Button.jsx';
import StatusMessage from '../components/UI/StatusMessage.jsx';
import ProductoHoraTable from '../components/Tables/ProductoHoraTable.jsx';
import { crearMockProductoHoraCuatroSemanas } from '../data/mockProductoHora.js';
import { obtenerProductoHoraCuatroSemanas } from '../services/productoHoraService.js';
import { calcularCuatroSemanasAnteriores, formatFechaParaInput } from '../utils/dateUtils.js';

function sumarCantidad(rows) {
  return rows.reduce((total, row) => total + row.cantidad, 0);
}

function ImportarProductoHora() {
  const [fechaBase, setFechaBase] = useState(formatFechaParaInput());
  const [semanas, setSemanas] = useState([]);
  const [loading, setLoading] = useState(false);
  const [generalError, setGeneralError] = useState('');
  const [lastSource, setLastSource] = useState('');

  const fechasCalculadas = useMemo(() => {
    try {
      return calcularCuatroSemanasAnteriores(fechaBase);
    } catch {
      return [];
    }
  }, [fechaBase]);

  const totalGeneral = semanas.reduce((total, semana) => total + sumarCantidad(semana.datos), 0);
  const totalFilas = semanas.reduce((total, semana) => total + semana.datos.length, 0);
  const semanasConError = semanas.filter((semana) => semana.error).length;

  async function cargarDatosReales() {
    setLoading(true);
    setGeneralError('');
    setLastSource('api');

    try {
      const resultados = await obtenerProductoHoraCuatroSemanas(fechaBase);
      setSemanas(resultados);
    } catch (error) {
      setGeneralError(error.message);
    } finally {
      setLoading(false);
    }
  }

  function cargarDatosMock() {
    setGeneralError('');
    setLastSource('mock');
    setSemanas(crearMockProductoHoraCuatroSemanas(fechaBase));
  }

  return (
    <section className="page">
      <div className="page-header split-header">
        <div>
          <p className="eyebrow">Primera etapa</p>
          <h1>Importar Producto Hora</h1>
          <p>
            Replica la parte critica del VBA: fecha base, cuatro semanas
            anteriores, consulta ProductoHora y extraccion de Hora, Producto y
            Cantidad.
          </p>
        </div>
        <div className="source-pill">
          {lastSource ? `Origen: ${lastSource === 'mock' ? 'Mock' : 'API real'}` : 'Sin carga'}
        </div>
      </div>

      <div className="card import-panel">
        <label className="field">
          <span>Fecha base</span>
          <input
            type="date"
            value={fechaBase}
            onChange={(event) => setFechaBase(event.target.value)}
          />
        </label>

        <div className="actions">
          <Button disabled={loading || !fechaBase} onClick={cargarDatosReales}>
            {loading ? 'Cargando...' : 'Cargar datos'}
          </Button>
          <Button disabled={loading || !fechaBase} variant="secondary" onClick={cargarDatosMock}>
            Usar datos de prueba
          </Button>
        </div>
      </div>

      <div className="card">
        <h2>Fechas calculadas</h2>
        <div className="weeks-grid">
          {fechasCalculadas.map((semana) => (
            <div className="week-date" key={semana.semana}>
              <span>{semana.etiqueta}</span>
              <strong>{semana.fechaApi}</strong>
            </div>
          ))}
        </div>
      </div>

      {loading && <StatusMessage>Cargando datos desde la web interna...</StatusMessage>}
      {generalError && <StatusMessage tone="error">{generalError}</StatusMessage>}
      {semanasConError > 0 && (
        <StatusMessage tone="warning">
          {semanasConError} semana(s) no pudieron cargarse. Si el error menciona CORS, se
          necesita un backend/proxy en una etapa futura.
        </StatusMessage>
      )}

      {semanas.length > 0 && (
        <div className="summary-grid">
          <div className="summary-card">
            <span>Semanas cargadas</span>
            <strong>{semanas.length - semanasConError}/4</strong>
          </div>
          <div className="summary-card">
            <span>Filas importadas</span>
            <strong>{totalFilas.toLocaleString('es-AR')}</strong>
          </div>
          <div className="summary-card">
            <span>Total cantidad</span>
            <strong>{totalGeneral.toLocaleString('es-AR')}</strong>
          </div>
        </div>
      )}

      <div className="week-results">
        {semanas.map((semana) => {
          const totalSemana = sumarCantidad(semana.datos);

          return (
            <article className="card week-card" key={semana.semana}>
              <div className="week-header">
                <div>
                  <h2>{semana.etiqueta}</h2>
                  <p>{semana.fechaApi}</p>
                </div>
                <div className="week-total">
                  <span>Total cantidad</span>
                  <strong>{totalSemana.toLocaleString('es-AR')}</strong>
                </div>
              </div>

              {semana.error && <StatusMessage tone="error">{semana.error}</StatusMessage>}
              {!semana.error && <ProductoHoraTable rows={semana.datos} />}
            </article>
          );
        })}
      </div>
    </section>
  );
}

export default ImportarProductoHora;
