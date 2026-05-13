import { calcularCuatroSemanasAnteriores } from '../utils/dateUtils.js';

export const mockProductoHora = [
  { hora: '11:00 - 12:00', producto: 'COMBO CRISPY', cantidad: 12 },
  { hora: '11:00 - 12:00', producto: 'MEGA BOX', cantidad: 8 },
  { hora: '12:00 - 13:00', producto: 'BUCKET MIX', cantidad: 18 },
  { hora: '13:00 - 14:00', producto: 'COMBO RUSTER', cantidad: 9 },
  { hora: '20:00 - 21:00', producto: 'COMBO POP', cantidad: 15 },
];

export function crearMockProductoHoraCuatroSemanas(fechaBase) {
  return calcularCuatroSemanasAnteriores(fechaBase).map((semana, index) => ({
    ...semana,
    datos: mockProductoHora.map((row, rowIndex) => ({
      ...row,
      cantidad: row.cantidad + index * 3 + rowIndex,
    })),
    error: null,
    source: 'mock',
  }));
}
