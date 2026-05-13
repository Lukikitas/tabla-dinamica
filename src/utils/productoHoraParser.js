function limpiarProducto(producto) {
  return String(producto ?? '').replaceAll('*', '').trim();
}

function normalizarCantidad(cantidad) {
  const numero = Number(cantidad);
  return Number.isFinite(numero) ? numero : null;
}

function crearRegistro(hora, producto, cantidad) {
  const productoLimpio = limpiarProducto(producto);
  const cantidadNumero = normalizarCantidad(cantidad);

  if (!hora || !productoLimpio || cantidadNumero === null) {
    return null;
  }

  return {
    hora: String(hora).trim(),
    producto: productoLimpio,
    cantidad: cantidadNumero,
  };
}

function extraerDesdeObjeto(data) {
  const registros = [];

  function recorrer(node, horaActual = null) {
    if (Array.isArray(node)) {
      node.forEach((item) => recorrer(item, horaActual));
      return;
    }

    if (!node || typeof node !== 'object') {
      return;
    }

    const hora = node.hora ?? node.cabecera ?? node.Hora ?? node.Cabecera ?? horaActual;

    if ('producto' in node || 'Producto' in node) {
      const registro = crearRegistro(
        hora,
        node.producto ?? node.Producto,
        node.cantidad ?? node.Cantidad,
      );

      if (registro) {
        registros.push(registro);
      }
    }

    if (Array.isArray(node.cuerpo)) {
      recorrer(node.cuerpo, hora);
    }

    Object.entries(node).forEach(([key, value]) => {
      if (key !== 'cuerpo') {
        recorrer(value, hora);
      }
    });
  }

  recorrer(data);
  return registros;
}

function leerPrimerString(texto) {
  const match = String(texto).match(/"([^"]+)"/);
  return match ? match[1] : '';
}

function leerCantidad(texto) {
  const match = String(texto).match(/"cantidad"\s*:\s*(-?\d+(?:[.,]\d+)?)/i);
  if (!match) return null;

  return normalizarCantidad(match[1].replace(',', '.'));
}

function parsearBloquesTexto(responseText, campoHora) {
  const registros = [];
  const partesHora = responseText.split(new RegExp(`"${campoHora}"\\s*:`, 'i'));

  for (let i = 1; i < partesHora.length; i += 1) {
    const bloqueHora = partesHora[i];
    const hora = leerPrimerString(bloqueHora);

    if (!hora) continue;

    const cuerpoMatch = bloqueHora.match(/"cuerpo"\s*:\s*\[([\s\S]*?)\](?=\s*[,}])/i);
    if (!cuerpoMatch) continue;

    const partesProducto = cuerpoMatch[1].split(/"producto"\s*:/i);

    for (let j = 1; j < partesProducto.length; j += 1) {
      const bloqueProducto = partesProducto[j];
      const producto = leerPrimerString(bloqueProducto);
      const cantidad = leerCantidad(bloqueProducto);
      const registro = crearRegistro(hora, producto, cantidad);

      if (registro) {
        registros.push(registro);
      }
    }
  }

  return registros;
}

function parsearTextoTolerante(responseText) {
  const porHora = parsearBloquesTexto(responseText, 'hora');
  if (porHora.length > 0) return porHora;

  return parsearBloquesTexto(responseText, 'cabecera');
}

export function parseProductoHoraResponse(responseText) {
  if (!responseText || !String(responseText).trim()) {
    return [];
  }

  try {
    const json = JSON.parse(responseText);
    const registros = extraerDesdeObjeto(json);

    if (registros.length > 0) {
      return registros;
    }
  } catch {
    // Si el responseText no es JSON estricto, se replica el enfoque tolerante del VBA.
  }

  return parsearTextoTolerante(String(responseText));
}
