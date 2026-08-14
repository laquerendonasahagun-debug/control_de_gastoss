# Control de gastos · La Querendona

Aplicación web estática para convertir el archivo `Control de Gastos.xlsx` en un control operativo de gastos para Tepeapulco.

## Incluye

- Resumen de gastos reales por semana, mes, rango de fechas y concepto.
- Captura individual y masiva de gastos con fecha, semana, concepto, responsable, monto, forma de pago y nota.
- Registro de gastos operativos y fijos con acumulados diarios, semanales y mensuales.
- Histórico semanal basado en las hojas `Hoja1`, `sep-dic25`, `2026-1sem` y `2026-2sem`.
- Exportación de movimientos e histórico a CSV.
- Persistencia local en el navegador mediante `localStorage`.

## Uso

Abre `index.html` en un navegador o publícalo como sitio estático (por ejemplo, GitHub Pages). No requiere build ni dependencias de servidor.

Los montos demo y las semanas iniciales fueron transcritos del libro de Excel proporcionado. Los nuevos registros se guardan en el navegador del usuario y se recalculan al instante.
