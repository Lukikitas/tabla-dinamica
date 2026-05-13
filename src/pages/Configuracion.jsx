import { productoHoraApiConfig } from '../services/apiConfig.js';

function Configuracion() {
  return (
    <section className="page">
      <div className="page-header">
        <p className="eyebrow">Parametros</p>
        <h1>Configuracion</h1>
        <p>
          Los valores de la API estan centralizados para evitar hardcodearlos en
          los componentes.
        </p>
      </div>

      <div className="card config-card">
        <div>
          <span>Base URL</span>
          <strong>{productoHoraApiConfig.baseUrl}</strong>
        </div>
        <div>
          <span>Restaurante</span>
          <strong>{productoHoraApiConfig.restaurante}</strong>
        </div>
        <div>
          <span>Cadena</span>
          <strong>{productoHoraApiConfig.cadena}</strong>
        </div>
        <div>
          <span>Usuario</span>
          <strong>{productoHoraApiConfig.usuario}</strong>
        </div>
      </div>
    </section>
  );
}

export default Configuracion;
