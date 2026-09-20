# Almacén Pro v2.0.0 — Google Apps Script + Google Sheets

Sistema web para control de inventario, Kardex, entregas a personal, reportes XLSX y backups en Google Drive.

## Qué cambia en v2.0

- Login con usuarios y roles (`ADMIN`, `SUPERVISOR`, `ALMACENERO`, `CONSULTA`).
- Auditoría: cada movimiento guarda el usuario que lo registró.
- Trazabilidad de salida: beneficiario, DNI, área/dependencia, cargo y documento/vale/PECOSA.
- Nuevo movimiento `DEVOLUCION`.
- Autocompletado de SKU con muestra de stock disponible.
- Dashboard con barras por familia, KPIs, alertas y últimas entregas.
- Reportes por fechas, SKU, familia, tipo, condición, persona y área.
- Exportación de reportes e inventario a `.xlsx`.
- Acta/constancia de entrega imprimible desde el reporte.
- Backup manual XLSX a la carpeta de Drive `Almacen Pro Backups`.
- Backup automático semanal o cada 4 semanas.
- Descarga JSON de los datos principales.

## Hojas creadas / ampliadas

- `PRODUCTOS`
- `MOVIMIENTOS`
- `USUARIOS`
- `CONFIG`
- `BACKUPS`

### MOVIMIENTOS

`ID | FECHA | FECHA_HORA | CODIGO | TIPO_MOVIMIENTO | CANTIDAD | CONDICION | MOTIVO | ORIGEN_DESTINO | TIPO_DESTINO | BENEFICIARIO | DNI_BENEFICIARIO | AREA_DESTINO | CARGO_BENEFICIARIO | DOCUMENTO_ENTREGA | USUARIO | OBSERVACION | CREADO_EN`

Con esta estructura, una salida puede quedar así:

`SALIDA | CHA-TAC-NEG-L | 1 | NUEVO | ENTREGA A PERSONAL | PERSONA | JUAN PEREZ | 41872874 | SERENAZGO | AGENTE | VALE-025 | admin | ...`

## Instalación

1. Abre Google Apps Script y crea un proyecto nuevo.
2. Copia `Code.gs`, `Index.html`, `Stylesheet.html`, `JavaScript.html` y `appsscript.json`.
3. Ejecuta `setupSistema()` una sola vez y concede permisos.
4. Verifica las hojas creadas y el usuario `admin`.
5. Implementa como `Aplicación web`, ejecutando como tu cuenta.
6. Abre la URL y entra con:

- Usuario: `admin`
- Contraseña: `admin`

7. Cambia posteriormente el acceso usando el módulo `Usuarios`.

## Notas

- La autenticación de la aplicación no reemplaza las políticas de acceso de la propia implementación de Google Apps Script.
- Los backups XLSX y la exportación Excel requieren permisos de Drive y Spreadsheet porque se usan para crear, exportar y guardar archivos.
- El indicador “Rotación referencial 30 días” es una métrica orientativa: salidas de 30 días / stock actual. No sustituye el cálculo contable clásico de inventario promedio.
