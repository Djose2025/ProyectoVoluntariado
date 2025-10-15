# 🧭 Manual Técnico del Proyecto de Análisis de Marcajes  
**Proyecto de Prácticas Iniciales - USAC 2025**  
Autor: **[Daniel José Colindres Fuentes / Carné 2014004445]**  
Versión: **1.0 (Etapa de Prueba y Validación en Equipo Destinado)**  
Fecha: **Octubre 2025**

---

## 🧱 1. Descripción General

Este proyecto tiene como objetivo **automatizar la verificación diaria de marcajes de asistencia** del personal de una institución.  
El sistema analiza los datos de marcaje del **día anterior**, identifica a los empleados con **ausencias o marcajes incompletos**, y genera **informes individuales** que posteriormente podrán ser enviados por correo electrónico.

Actualmente, esta versión:
- Genera los informes de manera **automática** en formato `.txt`.
- Crea la carpeta de salida **Marcajes** en el **escritorio del usuario**.
- Está diseñada para funcionar en **equipos antiguos con Windows 8** y **Microsoft Excel 2010 o posterior**.
- No requiere conexión a internet ni instalación adicional.

> 🔧 En versiones futuras se agregará el módulo de **envío automático de correos**, una vez verificada la compatibilidad con el equipo de destino (configuración de Outlook).

---

## 💻 2. Requisitos del Sistema

| Requisito | Especificación recomendada |
|------------|----------------------------|
| **Sistema operativo** | Windows 8, 8.1 o superior (64 bits preferiblemente) |
| **Microsoft Excel** | Versión 2010, 2013, 2016 o posterior |
| **Extensión del archivo** | `.xlsm` (Libro habilitado para macros) |
| **Procesador mínimo** | Intel Core i3 o i5 |
| **Permisos de usuario** | Permitir macros y acceso al sistema de archivos |
| **Software opcional** | Outlook configurado (para versión futura de envío automático) |

---

## 📄 3. Estructura del Archivo de Marcajes

La hoja principal del proyecto tiene la siguiente estructura:

| Columna | Encabezado | Descripción |
|----------|-------------|-------------|
| D13 | Nombre | Nombre completo del empleado |
| E13 | Día | Fecha del registro del marcaje |
| F13 | HoraEnt | Hora programada de entrada |
| G13 | HoraSal | Hora programada de salida |
| H13 | Marc-Ent | Hora real de entrada |
| I13 | Marc-Sal | Hora real de salida |

Los datos inician en la fila **14**.  
El sistema analiza todos los registros a partir de esa fila.

> ⚠️ Es fundamental que los encabezados estén **exactamente escritos** como se muestran arriba.  
> Si se modifican, el análisis fallará.

---

## ⚙️ 4. Macros Principales

El archivo contiene **dos macros principales**:

### 🔹 A. Subir archivo de marcajes

**Nombre interno:** `Sub CargarArchivoMarcajes()`

#### Función:
Permite cargar automáticamente un archivo externo de Excel que contenga los registros de marcaje del día anterior o de cualquier fecha.

#### Flujo:
1. Muestra una ventana para seleccionar el archivo (`.xls`, `.xlsx`, o `.xlsm`).
2. Abre el archivo seleccionado y copia los datos.
3. Pega los registros en la hoja principal, respetando la estructura.
4. Verifica los encabezados antes de copiar.
5. Muestra un mensaje confirmando la carga exitosa.

#### Posibles mensajes:
- “No se encontró el encabezado ‘Nombre’” → Los encabezados no coinciden.  
- “Estás intentando abrir un tipo de archivo bloqueado” → El archivo es antiguo; debe guardarse como `.xlsx`.

---

### 🔹 B. Analizar marcajes y generar informes

**Nombre interno:** `Sub AnalizarMarcajes_Auto_Mapeado()`

#### Función:
Recorre los datos de marcajes cargados, identifica a los empleados con faltas o marcajes incompletos, y genera **un informe individual en texto (.txt)** con los detalles del caso.

#### Flujo de trabajo:
1. Verifica la presencia de encabezados en la fila 13.  
2. Crea una carpeta en el escritorio llamada **Marcajes** (si no existe).  
3. Recorre todos los registros desde la fila 14 hasta la última.  
4. Detecta empleados que:
   - No marcaron entrada ni salida (ausencia completa).  
   - Solo marcaron entrada o salida (marcaje incompleto).  
5. Genera un archivo `.txt` por cada caso con nombre:
6. Guarda los informes dentro de la carpeta `Marcajes`.
7. Muestra un mensaje final indicando cuántos informes se generaron.

---

## 🧾 5. Ejemplo de informe generado

Ejemplo de archivo `Maria_Meneses_18-10-2025.txt`:


---

## 🧰 6. Instalación y ejecución paso a paso

1. Copia el archivo **Proyecto_Practicas_Iniciales_2014004445.xlsm** en tu equipo.  
2. Abre Excel y habilita las macros:
   - Ve a: **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Configuración de macros**.  
   - Selecciona:  
     ✅ *Habilitar todas las macros*  
     ✅ *Confiar en el acceso al modelo de objetos de VBA*
3. Abre el archivo `.xlsm`.  
4. Si Excel muestra una barra amarilla con el botón **Habilitar contenido**, haz clic en él.  
5. En la hoja de trabajo:
   - Pulsa **Cargar archivo** → selecciona tu archivo de marcajes.  
   - Luego pulsa **Crear informe** → el sistema generará los reportes en el escritorio.  
6. Verifica que se haya creado la carpeta **Marcajes** y que los archivos `.txt` estén dentro.

---

## 🧩 7. Funcionamiento interno

- **Control de fecha:** actualmente analiza todos los registros disponibles en la hoja, sin filtrar por fecha.  
  En futuras versiones se incluirá el filtrado automático por “día anterior”.
- **Generación automática:** los informes se crean con `Open For Output`, sin requerir intervención.  
- **Control de errores:** incluye validaciones para:
  - Falta de encabezados.  
  - Celdas vacías.  
  - Rutas inexistentes (crea carpetas si no están).  
- **Independencia:** el macro no necesita conexión a Outlook o internet en esta etapa.

---

## ⚠️ 8. Posibles errores y soluciones

| Error o mensaje | Causa probable | Solución |
|-----------------|----------------|-----------|
| Error 76: No se encontró la ruta de acceso | La carpeta del escritorio tiene otro nombre (Escritorio vs Desktop). | Cambiar la ruta en el código: `Environ("USERPROFILE") & "\Escritorio\Marcajes"`. |
| No se encontró el encabezado "Nombre" | Encabezados movidos o cambiados. | Verificar que “Nombre” esté en `D13`. |
| No se generaron informes | No hay empleados con marcajes incompletos. | Verificar las columnas “Marc-Ent” y “Marc-Sal”. |
| Error de macros deshabilitadas | Excel bloquea VBA. | Activar macros en el Centro de confianza. |
| Archivo bloqueado al abrir | El archivo de marcajes es antiguo. | Guardar una copia como `.xlsx`. |

---

## 🚀 9. Próximas ampliaciones

El proyecto está planificado para evolucionar en varias etapas:

1. **Etapa actual (v1.0)**  
   - Lectura de archivo externo  
   - Análisis de marcajes  
   - Generación de informes `.txt`

2. **Etapa siguiente (v2.0)**  
   - **Envío automático de correos** a los empleados detectados con ausencias.  
     - Integración con Outlook mediante `CreateObject("Outlook.Application")`.  
     - Envío del archivo generado como adjunto.  
     - Registro de envíos exitosos.

3. **Etapa avanzada (v3.0)**  
   - Exportación en formato **PDF o Word (.docx)**.  
   - Panel de control o formulario gráfico con calendario.  
   - Registro histórico consolidado de ausencias.  
   - Configuración personalizada de destinatarios y copia oculta (CC/BCC).

---

## 🧠 10. Notas del autor

- Este sistema fue desarrollado como parte de las **Prácticas Iniciales** de la carrera de **Ingeniería en Ciencias y Sistemas (USAC)**.  
- El diseño busca ser **ligero, compatible y autónomo**, evitando dependencias externas.  
- La documentación se ha elaborado cuidadosamente para que cualquier técnico o encargado pueda **instalar, ejecutar y diagnosticar errores** sin conocimientos avanzados de programación.  
- Se recomienda mantener una **copia de respaldo** del archivo antes de cada prueba y registrar los resultados en un documento de control.

---

## 📅 Historial de versiones

| Versión | Fecha | Descripción |
|----------|--------|-------------|
| 1.0 | 15/10/2025 | Versión inicial funcional. Análisis y generación automática de informes. |
| 1.1 (planificada) | 10/2025 | Integración con Outlook para envío de correos. |

---

✍️ **Documento redactado manualmente y revisado para uso institucional.**

