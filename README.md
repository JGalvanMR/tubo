Rama `master` encontrada.Ahora reviso form_tubo2.cs y form_tubo3.cs.Ahora reviso form_tubo3.cs (el selector de pallets/tarimas).📦 Nombre del Módulo: **tubo**

## 🧭 Propósito

Aplicación de escritorio en Windows Forms (.NET Framework 4.0) para **Comercializadora GAB / mrlucky** que registra y controla el proceso de **verificación de enfriamiento de producto** (materia prima y producto terminado) al pasar por distintas estaciones de frío: TUBO1, TUBO2, TUBO3, enfriador de aire forzado (coliflor), hidrocooler, hidrocooler grande y enhieladora. El sistema captura tiempos de llegada/entrada/salida, temperaturas del producto y del agua, pH y cloro del agua, y genera reportes exportables a Excel para control de calidad y trazabilidad.

## ⚙️ Responsabilidades

- Validar sesión de usuario contra un sistema externo (`SIPGAB`) antes de permitir el uso de la aplicación.
- Listar y filtrar por fecha los registros de enfriamiento capturados (`tb_mstr_tubo`).
- Dar de alta nuevos registros de enfriamiento vinculando un folio de recepción de materia prima (MP) o producto terminado (PTC) con tarimas/productos específicos.
- Completar (editar) registros existentes con datos de salida: temperatura de salida, temperatura final de agua, duración, pH y cloro.
- Seleccionar tarimas/cajas asociadas a un folio de recepción o de orden de producción, evitando duplicar tarimas ya capturadas.
- Calcular tiempos transcurridos entre horas de llegada/entrada/salida, incluyendo cruces de medianoche.
- Exportar a Excel (vía interoperabilidad COM) un reporte de verificación de producto de enfriamiento por rango de fechas, con detalle opcional por tarima.
- Registrar y notificar errores por correo electrónico a un destinatario fijo.

## 🔄 Flujo de Funcionamiento

1. **Arranque (`Program.cs`)**: se valida la IP de la máquina (`Utilerias.Class1.validar_ip()`) y luego se ejecuta `validar_login()`, que consulta `tb_cat_historial_dia` para verificar que exista una sesión activa de `SIPGAB` en la máquina actual. Si no hay sesión, se muestra un aviso y se lanza `SIPGAB.exe` desde una ruta fija, cerrando la aplicación.
2. **Pantalla principal (`Form1`)**: al cargar, construye varias `DataTable` en memoria (tiempos, exportación, detalle) y ejecuta `cargargrid()`, que trae de `tb_mstr_tubo` los registros del día actual y los pinta en una grilla (`dtgTiempos`). Un `timer1` refresca automáticamente la grilla cada 300 "ticks".
3. **Filtro por fecha** (`pbxFiltro_Click`): repite la consulta anterior pero con la fecha elegida por el usuario.
4. **Alta de registro** (`pbxNuevo_Click` → `form_tubo2` con `opcion = "A"`): abre el formulario de captura.
5. **Edición/consulta de registro** (doble clic en la columna de folio → `form_tubo2` con `opcion = "C"`): carga los datos existentes del folio seleccionado, incluyendo información de trazabilidad (proveedor, rancho, tabla) según si el folio es MP o PTC.
6. **Dentro de `form_tubo2` (alta)**: el usuario captura folio, hora de llegada/entrada, operador, lugar de enfriamiento, temperaturas de entrada, y selecciona tarimas mediante `form_tubo3`. Al guardar, se inserta en `tb_det_tubo` (detalle por tarima) y en `tb_mstr_tubo` (encabezado), o se actualiza la cantidad si el folio/hora ya existe.
7. **Dentro de `form_tubo2` (edición)**: solo se capturan los datos de salida (temperatura de salida, temperatura final de agua, duración, pH, cloro) mediante un `UPDATE`.
8. **`form_tubo3`**: consulta las tarimas disponibles para un folio (según sea recepción de planta, MP o PTC), excluyendo las que ya fueron registradas en `tb_det_tubo`, permite seleccionarlas individualmente o "Todos", y regresa los datos seleccionados al formulario padre mediante una clase estática compartida (`SharedData.Polino`).
9. **Exportación a Excel** (`pbxExcel_Click`): valida el rango de fechas, arma un reporte con encabezado de empresa, título, rango de fechas y tabla de datos (con detalle por tarima si se marca el checkbox), usando Excel vía COM Interop y el portapapeles para pegar los datos de la grilla.

## 📐 Reglas de Negocio

### 🔒 Restricciones
- La aplicación no permite operar sin una sesión activa validada contra `SIPGAB`; si no la hay, se cierra y relanza `SIPGAB.exe`.
- En modo edición/consulta ("C"), los campos de captura inicial quedan bloqueados; solo son editables los campos de salida si el recibo aún no está completo.
- Si el recibo ya tiene capturados temperatura de salida, temperatura final de agua y duración (> 0), se marca como "solo consulta" y se bloquea toda edición.
- En la exportación a Excel, la fecha de inicio no puede ser posterior a la fecha fin.
- Una tarima ya registrada en `tb_det_tubo` para un folio/tipo no vuelve a aparecer como disponible en `form_tubo3` (evita doble captura).

### ✅ Validaciones
- Alta de registro: son obligatorios folio, hora de entrada, operador, lugar de refrigeración, hora de llegada, cantidad de cajas y temperatura de entrada; la temperatura inicial del agua es obligatoria salvo que el lugar seleccionado sea "FRESCO".
- Las horas de llegada y entrada se validan como formato de hora convertible (`validahora`); en `form_tubo2`, además se valida con expresión regular `^(?:0?[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$`.
- Edición/salida: son obligatorios temperatura de salida, temperatura final de agua y duración; cada uno debe ser numérico (`validavalor`) y mayor a cero.
- Antes de aplicar tarimas seleccionadas en `form_tubo3`, se valida que existan tarimas listadas y que al menos una esté marcada.

### 🔁 Agrupaciones
- El tipo de recepción se determina por rango numérico de folio: folios ≤ 150000 se tratan como Materia Prima (MP) en algunas consultas, y folios > 170000 como Producto Terminado (PTC) en la lógica de alta; existe además un umbral de 120000 usado para decidir si el folio corresponde a una línea de producción o a un recibo de planta.
- Si un folio/hora de entrada ya existe en `tb_mstr_tubo`, la cantidad se acumula (`UPDATE ... SET cantidad = cantidad + ...`) en vez de crear un nuevo encabezado.
- Los lugares de enfriamiento se normalizan a un conjunto fijo de claves: TUBO1, TUBO2, TUBO3, ENFRIA, HIDROC, HIDROG, ENHIE, cada una con su nombre descriptivo (`traelugar2`).

### ⚙️ Reglas Operativas
- El cálculo de tiempo transcurrido (`calculatiempo`) contempla el caso de que la hora final cruce la medianoche (hora 00), ajustando el cálculo de horas/minutos en vez de usar una simple resta de `DateTime`.
- La hora de salida (`calculosalida`) se calcula sumando la duración (en minutos) a la hora de entrada.
- Para folios de tipo PTC/planta, la hora de encabezado (`hora_ent`/`hora`) se obtiene de `tb_mstr_recepcion_pt` o `tb_det_eti_final` en vez de la capturada manualmente por el usuario.
- La información de trazabilidad (proveedor, rancho, tabla) se obtiene de tablas distintas según el tipo de recibo: `tb_mstr_recepcion_mp` para MP, `tb_det_trazabilidad` para el resto.
- Ante cualquier `SqlException` o excepción general en las operaciones principales, se envía un correo de notificación fijo a `jbravo@mrlucky.com.mx` con el detalle del error antes de mostrar el mensaje al usuario.

## 🔗 Dependencias

- **Framework**: .NET Framework 4.0 (Client Profile), Windows Forms.
- **Base de datos**: SQL Server (`System.Data.SqlClient`), cadena de conexión centralizada en `Utilerias.Class1.ConnectionString`.
- **Librería interna**: `Utilerias.dll` (referencia local `HintPath`), que provee `ConnectionString`, `Login`, `Usu_login`, `Inicio_sesion`, `Nombre_equipo`, `validar_ip()` y `SendMail()`.
- **Interop COM**: `Microsoft.Office.Interop.Excel` (requiere Excel instalado en la máquina) y `VBIDE`.
- **Tablas SQL Server involucradas**: `tb_mstr_tubo`, `tb_det_tubo`, `tb_cat_producto`, `tb_cat_historial_dia`, `tb_mstr_ordenes_prod`, `tb_hist_recepcion`, `tb_mstr_recepcion_mp`, `tb_mstr_recepcion_pt`, `tb_det_trazabilidad`, `tb_det_eti_final`, `tb_cat_proveedor`, `tb_cat_ranchos`, `tb_cat_tablas`, `tb_cat_linea`.
- **Recurso de archivo fijo**: `C:\SisGabWeb\fondo_formularios.jpg` (imagen de fondo) y `C:\SisGabWeb\SIPGAB.exe` (ejecutable externo).
- **No determinable con la información disponible**: versión exacta de SQL Server, mecanismo de autenticación de la cadena de conexión (usuario/contraseña vs. integrada).

## ⚠️ Riesgos Técnicos

- **Inyección SQL generalizada**: todas las consultas e instrucciones INSERT/UPDATE se construyen por concatenación directa de texto de controles de UI (folios, fechas, horas, temperaturas), sin parámetros ni sanitización, en los tres formularios.
- **Credenciales y rutas de correo hardcodeadas**: el destinatario y remitente de notificación de errores (`jbravo@mrlucky.com.mx`, `jbravo`, `juanjose`) están escritos directamente en el código en múltiples puntos.
- **Rutas de archivo absolutas y fijas**: `C:\SisGabWeb\...` asume una estructura de carpetas específica en cada máquina cliente; si no existe, la app falla al iniciar (`Bitmap.FromFile`).
- **Dependencia frágil de Excel Interop**: requiere Microsoft Excel instalado localmente; el uso de Interop COM es propenso a fugas de memoria (objetos COM no liberados con `Marshal.ReleaseComObject`) y a bloqueos si Excel queda en segundo plano.
- **Uso del portapapeles del sistema** (`Clipboard.SetDataObject`) para transferir datos a Excel: puede fallar si otro proceso está usando el portapapeles simultáneamente, y no es thread-safe.
- **Lógica de negocio basada en rangos numéricos de folio hardcodeados** (120000, 150000, 170000): frágil ante cambios en la numeración/consecutivos del sistema origen; no hay una tabla de configuración.
- **Alto acoplamiento con `Utilerias.dll` y con el proceso externo `SIPGAB.exe`**, cuya ruta y comportamiento no están documentados en este contexto.
- **Cálculo manual de intervalos de tiempo** (`calculatiempo`) con múltiples ramas condicionales propensas a errores de borde (cruces de medianoche, horas iguales), sin usar directamente `TimeSpan` en todos los casos.
- **Estado compartido estático** (`form_tubo3.SharedData.Polino`) entre formularios: no es seguro en escenarios concurrentes y dificulta las pruebas unitarias.
- **Manejo de errores que expone detalles técnicos al usuario final** (`ex.ToString()` completo en `MessageBox.Show`), lo cual puede filtrar información sensible de la infraestructura.
- **Falta de transacciones** en operaciones que combinan múltiples INSERT/UPDATE (detalle y encabezado): un fallo a mitad de proceso puede dejar datos inconsistentes.

## 🧪 Casos Edge

- Cruce de medianoche entre hora de llegada y hora de entrada/salida: manejado con lógica condicional específica, pero con ramas ligeramente distintas entre `Form1.calculatiempo` y `form_tubo2.calculatiempo`.
- Un folio de recepción "ya enfriado completamente" (todas las cajas procesadas) genera un aviso pero no bloquea completamente el flujo en todos los casos.
- Selección de tarimas cuando no hay ninguna disponible o ninguna marcada: se maneja con mensajes, pero el conteo (`counter`) se recalcula manualmente en vez de usar una única fuente de verdad.
- Recibos de planta (`chkplanta` marcado) con folio no encontrado en `tb_mstr_ordenes_prod`: deshabilita el flujo de selección de tarimas sin cerrar el formulario.
- Cantidad acumulada en `tb_mstr_tubo` cuando el mismo folio/hora se captura más de una vez: se suma en vez de generar un nuevo registro, lo cual puede ocultar dobles capturas si el usuario repite la operación por error.
- **No determinable con la información disponible**: comportamiento si `Utilerias.Class1.SendMail` falla (por ejemplo, sin conexión SMTP) durante el manejo de una excepción.

## 🧱 Suposiciones Detectadas

- Se asume que la máquina cliente siempre tiene la ruta `C:\SisGabWeb\` disponible con los recursos necesarios.
- Se asume que el usuario ya inició sesión previamente en el sistema `SIPGAB` desde la misma máquina antes de abrir `tubo`.
- Se asume que Microsoft Excel está instalado y con licencia activa en cada equipo donde se use la exportación.
- Se asume una relación fija entre rangos numéricos de folio y tipo de recepción (MP/PTC), sin validación explícita contra un catálogo.
- Se asume que las horas capturadas por el usuario siempre corresponden al mismo día operativo, salvo el caso de cruce de medianoche contemplado manualmente.
- Se asume conectividad de red constante con el servidor SQL Server desde la máquina cliente.

## 📈 Recomendaciones Técnicas

- Migrar todas las consultas a **comandos parametrizados** (`SqlParameter`) para eliminar el riesgo de inyección SQL.
- Externalizar credenciales, direcciones de correo y rutas de archivo a un archivo de configuración (`app.config`) o variables de entorno.
- Reemplazar la dependencia de Excel Interop por generación de archivos `.xlsx` mediante una librería sin dependencia de Office instalado (p. ej. ClosedXML o EPPlus), evitando fugas de objetos COM.
- Sustituir los rangos numéricos de folio hardcodeados por un catálogo o campo explícito de "tipo de recepción" en la base de datos.
- Envolver las operaciones de INSERT/UPDATE relacionadas (detalle + encabezado) en una transacción SQL para garantizar consistencia.
- Centralizar y simplificar el cálculo de intervalos de tiempo en una sola función reutilizable y probada con pruebas unitarias, evitando duplicación entre formularios.
- Registrar errores en un log estructurado (archivo o tabla de auditoría) además de notificar por correo, para poder diagnosticar fallas sin exponer detalles técnicos al usuario final.
- Sustituir el estado estático compartido entre formularios por paso explícito de datos (parámetros de constructor o eventos), facilitando pruebas y mantenimiento.

## 🧾 Resumen Ejecutivo

Este sistema es la herramienta que el personal de calidad usa para verificar que la fruta y otros productos pasen correctamente por las distintas etapas de enfriamiento (tubos, hidrocoolers, enfriadores) antes de continuar su proceso, registrando tiempos y temperaturas para asegurar que se cumplan los estándares de frío requeridos. También sirve para generar reportes en Excel que documentan este control de calidad por periodo. El sistema funciona, pero fue construido de forma poco robusta: guarda contraseñas y direcciones de correo directamente en el código, es vulnerable a manipulación de la base de datos por la forma en que arma sus consultas, depende de que cada computadora tenga Excel instalado y una carpeta específica con archivos, y basa reglas importantes del negocio en rangos de números de folio en lugar de catálogos configurables. Esto significa que, aunque cumple su función diaria, representa un riesgo de seguridad y de estabilidad que conviene atender antes de que crezca su uso o se integre con otros sistemas.
