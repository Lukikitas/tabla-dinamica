function ProductoHoraTable({ rows }) {
  if (!rows.length) {
    return <p className="empty-state">No hay datos para mostrar.</p>;
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Hora</th>
            <th>Producto</th>
            <th className="number-cell">Cantidad</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={`${row.hora}-${row.producto}-${index}`}>
              <td>{row.hora}</td>
              <td>{row.producto}</td>
              <td className="number-cell">{row.cantidad.toLocaleString('es-AR')}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default ProductoHoraTable;
