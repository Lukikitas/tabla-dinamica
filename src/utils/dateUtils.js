function toDateOnly(fecha) {
  if (!fecha) {
    throw new Error('La fecha base es obligatoria.');
  }

  const date = fecha instanceof Date ? new Date(fecha) : new Date(`${fecha}T00:00:00`);

  if (Number.isNaN(date.getTime())) {
    throw new Error('La fecha ingresada no es valida.');
  }

  return date;
}

export function formatFechaParaApi(fecha) {
  const date = toDateOnly(fecha);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}/${month}/${day}`;
}

export function formatFechaParaInput(fecha = new Date()) {
  const date = toDateOnly(fecha);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

export function calcularCuatroSemanasAnteriores(fechaBase) {
  const base = toDateOnly(fechaBase);

  return [1, 2, 3, 4].map((semana) => {
    const fecha = new Date(base);
    fecha.setDate(base.getDate() - 7 * semana);

    return {
      semana,
      fecha,
      fechaApi: formatFechaParaApi(fecha),
      etiqueta: `Semana ${semana}`,
    };
  });
}
