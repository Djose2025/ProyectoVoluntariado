# Manual Técnico — Automatización de Registro de Asistencia (Excel VBA)

Última actualización: 2025-10-28  
Autor: Daniel José Colindres Fuentes
Registro academico: 201404445

## Resumen
Este manual describe la macro VBA "Asistencia_Gmail_CDO" que:
- importa marcajes desde un archivo Excel,
- normaliza y busca correos en una hoja `Empleados`,
- detecta omisiones (entrada/salida/ausencia) por jornada,
- genera un informe TXT con el detalle,
- envía correos automáticos por Gmail (CDO) usando App Password (opcional).

Se diseñó para Excel 2007 en Windows 8/10 y pretende ser fácil de auditar y mantener.

---

## Contenido del manual
1. Requisitos previos  
2. Archivos / Hojas esperadas  
3. Instalación del módulo VBA  
4. Configuración principal (constantes)  
5. Descripción técnica del código (módulos y funciones)  
6. Flujo de ejecución paso a paso  
7. Plantilla de asunto y cuerpo del correo  
8. Pruebas y modo seguro (DryRun)  
9. Gestión de credenciales (App Password)  
10. Registro e informe (auditoría)  
11. Resolución de problemas comunes  
12. Buenas prácticas y seguridad  
13. Extensiones y mejoras futuras

---

## 1. Requisitos previos
- Microsoft Excel 2007 (habilitar macros)  
- Windows 8 o Windows 10 (mejor compatibilidad TLS para CDO)  
- Cuenta Gmail para envío (opcional): activar 2-Step Verification y crear App Password (si se usará envío automático).  
- Permisos para ejecutar macros y acceso a Internet desde la máquina que ejecuta el envío SMTP.  

---

## 2. Archivos / Hojas esperadas
Libro (Workbook) donde se instala la macro, con dos hojas:

- Hoja: `Hoja1`  
  - Encabezados (fila 13): D13..I13 -> `Nombre`, `Dia`, `HoraEnt`, `HoraSal`, `Marc-Ent`, `Marc-Sal`  
  - Datos desde D14 hacia abajo (cada fila = 1 jornada).

- Hoja: `Empleados`  
  - Encabezados (fila 4): D4..F4 -> `Nombre completo`, `Área`, `Correo`  
  - Datos desde D5 hacia abajo (columna D = nombre, F = correo).

Nota: Si tus encabezados están en otra posición, cambiar las constantes en la sección de configuración del módulo.

---

## 3. Instalación del módulo VBA
1. Abrir el libro en Excel y habilitar macros (guardar como `.xlsm`).  
2. ALT + F11 → Insertar → Módulo.  
3. Copiar y pegar el módulo VBA proporcionado (archivo `Asistencia_Gmail_CDO_Final.bas`).  
4. Guardar el libro.

Recomendación: crear una copia de seguridad del libro antes de modificar macros.

---

## 4. Configuración principal (constantes)
En la parte superior del módulo hay una sección `CONFIGURACIÓN` con constantes para ajustar:

- `SHEET_MARCAJES` — nombre de la hoja de marcajes (por defecto `"Hoja1"`).  
- `ROW_HEADERS_MAR`, `START_ROW_MAR` — fila de encabezado y primera fila de datos (por defecto 13 y 14).  
- `SHEET_EMPLEADOS`, `ROW_HEADERS_EMP`, `START_ROW_EMP` — hoja y filas para `Empleados` (por defecto `"Empleados"`, 4 y 5).  
- Columnas (números): `COL_MAR_NOMBRE`, `COL_MAR_DIA`, `COL_MAR_HORAENT`, `COL_MAR_HORASAL`, `COL_MAR_MARC_ENT`, `COL_MAR_MARC_SAL`.  
- Columnas de `Empleados`: `COL_EMP_NOMBRE`, `COL_EMP_AREA`, `COL_EMP_CORREO`.  
- Modo de envío: `SendMode = "SMTP_CDO"` (usar Gmail) o `"REPORT_ONLY"` (solo informe).  
- SMTP / Gmail:
  - `SMTP_SERVER = "smtp.gmail.com"`
  - `SMTP_PORT = 465`
  - `SMTP_USE_SSL = True`
  - `SMTP_USER = "tuCuenta@gmail.com"` (cambiar por la cuenta remitente)
  - `EMAIL_FROM = "Secretaría Académica <tuCuenta@gmail.com>"`

- `DryRun` (True/False): True = simula (no envía).  
- `MostrarMensajes`: controla MsgBox final.

Cambiar únicamente estas constantes según su entorno.

---

## 5. Descripción técnica del código (funciones y módulos)
El módulo está organizado y documentado. Resumen por secciones:

### Helpers / Normalización
- `NormalizeName(s As String) As String`  
  Normaliza nombres: trim, lowercase, elimina diacríticos, quita puntuación y colapsa espacios. Permite emparejar nombres sin tildes (archivo marcajes) con nombres con acentos (Empleados).

- `RemoveDiacritics(s As String) As String`  
  Reemplaza caracteres acentuados por su equivalente sin tilde (á→a, ñ→n, etc.)

- `DateToSpanishLong(d As Variant) As String`  
  Devuelve la fecha en formato "08 de junio de 2025" para incluir en cuerpo del correo.

### Carga de empleados
- `CargarEmpleadosEnDicts(wsEmp As Worksheet, dictEmail As Object, dictName As Object)`  
  Lee la hoja `Empleados` y crea dos diccionarios:
  - `dictEmail(claveNorm) = correo`
  - `dictName(claveNorm) = nombreOriginal`
  donde `claveNorm = NormalizeName(nombre)`.

### Matching parcial
- `FindBestMatchEmail(nameNorm As String, dictEmail As Object, dictName As Object, ByRef bestNameOut As String, ByRef bestScoreOut As Long) As String`  
  Si no hay match exacto, intenta encontrar el mejor match por tokens (palabras) y devuelve el email si la coincidencia supera un umbral (al menos 2 tokens o la mitad de tokens). También devuelve el `bestNameOut` y `bestScoreOut`.

### Construcción de asunto y cuerpo
- `BuildSubject(diaVal As Variant) As String`  
  Genera el asunto solicitado:  
  `Recordatorio y Solicitud de Regularización de Marcaje de Asistencia - Jornada del dd/mm/yyyy`

- `BuildBody(nombre, jornadaLabel, diaVal, horaEnt, horaSal, marcEnt, marcSal, estadoDetectado) As String`  
  Crea el cuerpo del correo en español con la plantilla solicitada, incluyendo:
  - Saludo personalizado
  - Jornada y fecha (fecha larga)
  - Detalle de entradas/salidas (con "[Ausente]" cuando falta)
  - Estado detectado con texto legible
  - Solicitud de regularización y despedida

### Envío SMTP (Gmail / CDO)
- `EnviarPorSMTP_CDO_Gmail(dest, subj, body, smtpPass) As String`  
  Usa `CDO.Message` y `CDO.Configuration` para enviar por SMTP. Requiere `smtpPass` (App Password). En `DryRun=True` no envía y retorna "DryRun - simulación (SMTP)".

### Análisis y generación de informe
- `AnalizarMarcajes_GenerarInforme(wsMarc, wsEmp, mode, smtpPass)`  
  Flujo principal:
  - Carga `Empleados` en diccionarios normalizados.
  - Recorre `Hoja1` desde `START_ROW_MAR`.
  - Cuenta jornadas por (nombre + día) para etiquetar "Jornada 1", "Jornada 2".
  - Determina estado: `OK`, `Falta: Entrada`, `Falta: Salida`, `Ausencia`.
  - Busca correo: exacto o sugerido (token matching).
  - Construye `asunto` y `cuerpo`.
  - Envía (según `mode`) o solo registra (REPORT_ONLY).
  - Genera un archivo `Informe_Marcajes_YYYYMMDD_HHMMSS.txt` con columnas:
    ```
    Nombre | Dia | Jornada | HorarioEsperado | Marc-Ent | Marc-Sal | Estado | EmailUsado | TipoMatch | EnvioResultado
    ```

### Punto de entrada
- `AnalizarMarcajes_Principal()`:
  - Pide App Password (InputBox) si `SendMode = "SMTP_CDO"`.
  - Llama a `AnalizarMarcajes_GenerarInforme`.

---

## 6. Flujo de ejecución (paso a paso)
1. Preparar archivo de marcajes y hoja `Empleados`.  
2. Abrir libro `.xlsm` con macros.  
3. Opcional: usar macro `CargarArchivoMarcajes` para importar marcajes (A1:F1 → D14..).  
4. Ajustar configuración (`SMTP_USER`, `EMAIL_FROM`, `DryRun` si se desea).  
5. Ejecutar `AnalizarMarcajes_Principal` (ALT+F8).  
   - Si `SMTP_CDO`, la macro solicita App Password (no se guarda en código).  
6. Revisar `Informe_Marcajes_*.txt` generado en la misma carpeta del workbook.  
7. Si DryRun = True, la macro no envía correos; revisar informe y ajustar Empleados si es necesario.  
8. Poner DryRun = False y volver a ejecutar para envío real (recomendar enviar por lotes en pruebas).

---

## 7. Plantilla de asunto y cuerpo
- Asunto:
  ```
  Recordatorio y Solicitud de Regularización de Marcaje de Asistencia - Jornada del dd/mm/yyyy
  ```
- Cuerpo (estructura generada por `BuildBody`):
  ```
  Estimada <Nombre>,

  El presente correo tiene como finalidad informarle sobre una incidencia detectada en el registro de su marcaje de asistencia correspondiente a la <Jornada X> del día <08 de junio de 2025>.

  Detalle del Registro Detectado:
  ----------------------------------------
  Hora Registrada
  Entrada: 05:37
  Salida: [Ausente]

  Estado Detectado: Falta de registro de la hora de Salida.

  Agradeceremos su colaboración para regularizar su marcaje a la brevedad posible, en caso de que corresponda una justificación o corrección de la hora de salida.

  Por favor, siga el procedimiento interno establecido para realizar dicha regularización.

  Agradecemos de antemano su atención y pronta gestión.

  Atentamente,
  Secretaría Académica
  Facultad de Ingeniería
  ```

El texto se ajusta automáticamente a la fila actual (nombre, jornada, fecha, horas y estado).

---

## 8. Pruebas y modo seguro (DryRun)
- Antes de enviar correos masivos, establecer `DryRun = True`.  
- Ejecutar `AnalizarMarcajes_Principal`.  
- Revisar el archivo `Informe_Marcajes_...txt`: validación de `EmailUsado`, `TipoMatch` y `EnvioResultado`.  
- Si los matches sugeridos son correctos, pasar a `DryRun = False` y probar con 3-5 registros.

---

## 9. Gestión de credenciales (App Password)
Recomendado: crear una cuenta Gmail dedicada para envíos automáticos.

Pasos para App Password:
1. Habilitar Verificación en dos pasos en la cuenta Gmail.  
2. Acceder a https://myaccount.google.com/security → "Contraseñas de aplicaciones".  
3. Crear una nueva contraseña para “Correo” o “Otro” y copiar la clave de 16 caracteres.  
4. Al ejecutar la macro la primera vez (o cada ejecución si así lo desea), pegar la App Password cuando se solicite.  
5. No almacenar la App Password en el código en texto plano. Opciones disponibles:
   - Ingresar por `InputBox` cada ejecución (actual implementacion).
   - Guardar en hoja oculta protegida (menos seguro) — se puede añadir si se requiere.

Límites y advertencias:
- Gmail impone límites de envío diario y por minuto. Para alto volumen, usar cuenta institucional o servicio de correo transaccional (SendGrid, Mailgun).

---

## 10. Registro e informe (auditoría)
- Cada ejecución genera `Informe_Marcajes_YYYYMMDD_HHMMSS.txt` en la carpeta del workbook.  
- El informe contiene:
  - Nombre, fecha, jornada, horario esperado, marc-Ent, marc-Sal, estado, email usado, tipo match (Exacto/Sugerido/No encontrado), resultado del envío.  
- Mantener estos informes archivados para auditoría.

---

## 11. Resolución de problemas comunes

- Error 9 (Subscript out of range) en `Worksheets("Hoja1")`:  
  - Verifique el nombre exacto de la pestaña. Use la macro `ListarNombresDeHojas` o renombre la pestaña.

- Error al abrir archivo importado / Error 424:  
  - El archivo seleccionado puede no ser un Excel válido o estar bloqueado. Cerrar el archivo en otras aplicaciones y reintentar.

- No encuentra correo (EmailUsado vacío):  
  - Revise coincidencia de nombres. Ejecute en `DryRun=True`.  
  - Use la hoja `Empleados` con nombres completos. Si existe diferencia en tildes o acentos, la normalización las quita; sin embargo, si en `Empleados` los nombres tienen apellidos adicionales, el algoritmo usa token-matching y sugiere coincidencias. Revise `TipoMatch` en el informe.

- Error SMTP/TLS al enviar:  
  - Verifique App Password, puerto (465), acceso a internet, firewall que bloquee salida por ese puerto.  
  - Si el equipo no soporta TLS moderno, considere usar un servidor SMTP institucional o API (SendGrid).

- Correos marcados como SPAM:  
  - Revisar `From`, configurar dominio remitente verificado o usar servicios transaccionales con autenticación (DKIM/SPF).

---

## 12. Buenas prácticas y seguridad
- No guardar App Password en texto plano dentro del libro.  
- Usar cuenta de correo dedicada para envíos automatizados.  
- Ejecutar en modo DryRun antes de producción.  
- Mantener informe por ejecución para trazabilidad.  
- Si varios operadores usarán la macro: documentar proceso y permisos de ejecución (quién tiene acceso a App Password).  
- Para volúmenes altos o mayor confiabilidad, migrar a una solución centralizada (script Python en servidor, Power Automate, o servicio de colas y envío).

---

## 13. Extensiones y mejoras futuras (sugeridas)
- Almacenar histórico de envíos en hoja `EnviosHistorico` o base de datos (SQLite).  
- Reemplazar matching por función Levenshtein para fuzzy matching más robusto.  
- Implementar interfaz (UserForm) para ingresar App Password de forma segura (campo oculto).  
- Integrar envío mediante API (SendGrid) para evitar problemas TLS/SMTP.  
- Añadir reintentos y backoff en envíos fallidos.  
- Crear versión con logging en tabla Excel además del TXT.

---
