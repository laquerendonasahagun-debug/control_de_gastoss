# Control de gastos · La Querendona

Aplicación web estática para convertir el archivo `Control de Gastos.xlsx` en un control operativo de gastos para Tepeapulco.

## Incluye

- Resumen de gastos reales por semana, mes, rango de fechas y concepto.
- Captura individual y masiva de gastos con fecha, semana, concepto, responsable, monto, forma de pago y nota.
- Registro de gastos operativos y fijos con acumulados diarios, semanales y mensuales.
- Histórico semanal basado en las hojas `Hoja1`, `sep-dic25`, `2026-1sem` y `2026-2sem`.
- Exportación de movimientos e histórico a CSV.
- Persistencia compartida en Neon Postgres para consultar los mismos gastos desde varios dispositivos.

## Uso

La interfaz se publica en Vercel y usa la función `/api/expenses` para acceder a Neon sin exponer credenciales en el navegador. El proyecto requiere la variable de entorno `DATABASE_URL` con la cadena de conexión de Neon.

Los montos demo y las semanas iniciales fueron transcritos del libro de Excel proporcionado. Al abrir la nueva versión por primera vez, los registros existentes en `localStorage` se migran automáticamente a Neon y luego se eliminan del almacenamiento local.
