import { productoHoraApiConfig } from './apiConfig.js';
import { calcularCuatroSemanasAnteriores, formatFechaParaApi } from '../utils/dateUtils.js';
import { parseProductoHoraResponse } from '../utils/productoHoraParser.js';

export function construirUrlProductoHora(fecha) {
  const fechaApi = formatFechaParaApi(fecha);
  const params = new URLSearchParams({ fecha: fechaApi });

  return `${productoHoraApiConfig.proxyPath}?${params.toString()}`;
}

export async function obtenerProductoHoraPorFecha(fecha) {
  const fechaApi = formatFechaParaApi(fecha);
  const url = construirUrlProductoHora(fecha);

  let response;

  try {
    response = await fetch(url);
  } catch (error) {
    throw new Error(
      `No se pudo consultar el proxy ProductoHora para ${fechaApi}. Detalle: ${error.message}`,
    );
  }

  if (!response.ok) {
    let detalle = '';

    try {
      const errorBody = await response.json();
      detalle = errorBody.error ? ` ${errorBody.error}` : '';
    } catch {
      detalle = '';
    }

    throw new Error(
      `El proxy ProductoHora respondio HTTP ${response.status} para la fecha ${fechaApi}.${detalle}`,
    );
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
