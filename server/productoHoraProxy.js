import { createServer } from 'node:http';
import { productoHoraApiConfig, productoHoraPath } from '../src/services/apiConfig.js';

const PORT = Number(process.env.PORT ?? 3001);

function sendJson(res, status, body) {
  res.writeHead(status, {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json; charset=utf-8',
  });
  res.end(JSON.stringify(body));
}

function buildProductoHoraUrl(fechaApi) {
  const { baseUrl, restaurante, usuario, cadena } = productoHoraApiConfig;
  const url = new URL(productoHoraPath, baseUrl);

  url.searchParams.set('restaurante', String(restaurante));
  url.searchParams.set('fecha_inicio', fechaApi);
  url.searchParams.set('fecha_fin', fechaApi);
  url.searchParams.set('estado', '');
  url.searchParams.set('canal', 'undefined');
  url.searchParams.set('usuario', usuario);
  url.searchParams.set('cadena', String(cadena));
  url.searchParams.set('cajero', 'undefined');

  return url;
}

function isValidFechaApi(fecha) {
  return /^\d{4}\/\d{2}\/\d{2}$/.test(fecha);
}

async function handleProductoHora(req, res, requestUrl) {
  const fecha = requestUrl.searchParams.get('fecha');

  if (!fecha || !isValidFechaApi(fecha)) {
    sendJson(res, 400, {
      error: 'Parametro fecha invalido. Use el formato yyyy/mm/dd.',
    });
    return;
  }

  const upstreamUrl = buildProductoHoraUrl(fecha);

  try {
    const upstreamResponse = await fetch(upstreamUrl);
    const responseText = await upstreamResponse.text();

    if (!upstreamResponse.ok) {
      sendJson(res, 502, {
        error: `ProductoHora respondio HTTP ${upstreamResponse.status} para ${fecha}.`,
        status: upstreamResponse.status,
        body: responseText,
      });
      return;
    }

    res.writeHead(200, {
      'Access-Control-Allow-Origin': '*',
      'Content-Type':
        upstreamResponse.headers.get('content-type') ?? 'text/plain; charset=utf-8',
      'X-Upstream-Status': String(upstreamResponse.status),
    });
    res.end(responseText);
  } catch (error) {
    sendJson(res, 502, {
      error:
        'No se pudo conectar con la API interna ProductoHora desde el proxy. Verifique que el entorno tenga acceso a la red interna.',
      detail: error.message,
      upstream: upstreamUrl.toString(),
    });
  }
}

const server = createServer(async (req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);

  if (req.method === 'OPTIONS') {
    sendJson(res, 204, {});
    return;
  }

  if (req.method === 'GET' && requestUrl.pathname === '/api/health') {
    sendJson(res, 200, {
      ok: true,
      service: 'producto-hora-proxy',
    });
    return;
  }

  if (req.method === 'GET' && requestUrl.pathname === '/api/producto-hora') {
    await handleProductoHora(req, res, requestUrl);
    return;
  }

  sendJson(res, 404, {
    error: 'Endpoint no encontrado.',
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`ProductoHora proxy escuchando en http://localhost:${PORT}`);
});
