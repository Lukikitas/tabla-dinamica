import { productoHoraApiConfig, productoHoraPath } from './apiConfig.js';
import { calcularCuatroSemanasAnteriores, formatFechaParaApi } from '../utils/dateUtils.js';
import { parseProductoHoraResponse } from '../utils/productoHoraParser.js';

export function construirUrlProductoHora(fecha) {
  const fechaApi = formatFechaParaApi(fecha);
  const { baseUrl, restaurante, usuario, cadena } = productoHoraApiConfig;

  return `${baseUrl}${productoHoraPath}?restaurante=${restaurante}&fecha_inicio=${fechaApi}&fecha_fin=${fechaApi}&estado=&canal=undefined&usuario=${usuario}&cadena=${cadena}&cajero=undefined`;
}

export async function obtenerProductoHoraPorFecha(fecha) {
  const fechaApi = formatFechaParaApi(fecha);
  const url = construirUrlProductoHora(fecha);

  let response;

  try {
    response = await fetch(url);
  } catch (error) {
    throw new Error(
      `No se pudo consultar ProductoHora para ${fechaApi}. Puede ser un problema de red o CORS. Detalle: ${error.message}`,
    );
  }

  if (!response.ok) {
    throw new Error(`ProductoHora respondio HTTP ${response.status} para la fecha ${fechaApi}.`);
  }

  const responseText = await response.text();
  return parseProductoHoraResponse(responseText);
}

export async function obtenerProductoHoraCuatroSemanas(fechaBase) {
  const semanas = calcularCuatroSemanasAnteriores(fechaBase);

  return Promise.all(
    semanas.map(async (semana) => {
      try {
        const datos = await obtenerProductoHoraPorFecha(semana.fecha);

        return {
          ...semana,
          datos,
          error: null,
          source: 'api',
        };
      } catch (error) {
        return {
          ...semana,
          datos: [],
          error: error.message,
          source: 'api',
        };
      }
    }),
  );
}
