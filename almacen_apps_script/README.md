# Almacén Pro — Prendas & Equipamiento

Aplicación web para controlar un almacén de prendas y equipamiento con Google Apps Script + Google Sheets.

## Funcionalidades

- Dashboard con SKUs, stock total, stock bajo y agotados.
- Catálogo de variantes: tipo, modelo, color, talla, unidad y stock mínimo.
- Entradas, salidas y ajustes (+/-).
- Validación para impedir salidas superiores al stock disponible.
- Kardex por SKU con saldo corrido.
- Historial de movimientos.
- Exportación CSV del inventario.
- Interfaz responsive para PC y móvil.
- Registro de usuario mediante `Session.getActiveUser().getEmail()` cuando Google lo expone.
- Bloqueo concurrente con `LockService` para evitar movimientos simultáneos que rompan el stock.
- Google Sheets como base de datos auditable.

## Archivos

- `Code.gs` — lógica servidor, base de datos y API interna.
- `Index.html` — interfaz principal.
- `Stylesheet.html` — estilos.
- `JavaScript.html` — lógica del frontend.
- `appsscript.json` — configuración del proyecto.

## Instalación

1. Ve a https://script.google.com/ y crea un proyecto nuevo.
2. Crea los archivos anteriores con exactamente esos nombres y pega su contenido.
3. En el editor, ejecuta `setupSistema()` una sola vez.
4. Autoriza los permisos solicitados.
5. `setupSistema()` creará una hoja de cálculo llamada **Almacén de Prendas y Equipamiento** y guardará su ID en Script Properties.
6. Revisa la hoja y reemplaza los datos de demostración por tu inventario real.
7. En Apps Script entra a **Implementar → Nueva implementación → Aplicación web**.
8. Configura quién ejecuta la aplicación y quién tiene acceso según la política de tu organización.
9. Abre la URL generada.

## Datos de demostración

Se incluyen como ejemplo las variantes de polacas y chalecos mencionadas en el requerimiento. La carga inicial aparece con el motivo `Carga de demostración` y observación `Revisar y reemplazar por inventario real`.

Las cantidades demo de polacas son: S 3, M 11, L 23, XL 21, XXL 23, 3XL 8, 4XL 3. Su suma es 92.

Las cantidades demo de chaleco táctico negro son: S 2, M 3, L 17. Su suma es 22.

También se crea una variante demo de manga de protección negra talla L con stock 0 para mostrar el caso sin existencias.

## Modelo de datos

### PRODUCTOS
Código, tipo, modelo, color, talla, descripción, unidad, stock mínimo, activo.

### MOVIMIENTOS
Fecha, código, tipo de movimiento, cantidad, motivo, origen/destino, usuario, observación.

El stock no se almacena en una celda editable: se calcula a partir de los movimientos. Esto permite conservar un historial tipo Kardex.

## Recomendaciones de despliegue

Para uso institucional, limita el acceso de la Web App a las cuentas de Google de la organización o al dominio autorizado. No publiques una aplicación de inventario interna como pública sin una política de acceso apropiada.

## Evolución sugerida

La siguiente versión puede agregar: roles (Administrador/Almacenero/Consulta), lectores de código de barras/QR, impresión de vales de salida, firma de entrega, actas PDF, usuarios/áreas, proveedores, importación masiva desde Excel/CSV y respaldo automático.
