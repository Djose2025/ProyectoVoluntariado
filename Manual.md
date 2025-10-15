# 📘 Proyecto de Control de Marcajes con Macros de Excel (VBA)

## 📖 Descripción general

Este proyecto fue desarrollado para **analizar marcajes de asistencia** de trabajadores y **generar automáticamente informes** de ausencias o marcajes incompletos.  
Está diseñado para funcionar en computadoras con **Windows 8 o posterior** y **Microsoft Excel 2010 o superior**, incluso en equipos con recursos limitados (procesador i3/i5).

El sistema trabaja con **dos macros principales**:

1. **Subir archivo de marcajes**  
   Permite cargar un archivo externo de Excel (de cualquier formato) con los registros de marcajes.

2. **Crear informe / Analizar marcajes**  
   Examina los datos cargados, identifica empleados con ausencias o marcajes incompletos, y genera un informe `.txt` por cada caso en una carpeta del escritorio.

---

## ⚙️ Estructura del archivo Excel

El archivo principal contiene una hoja (por ejemplo, llamada “Menú” o “Marcajes”) con los datos organizados así:

| Columna | Encabezado  | Descripción |
|----------|--------------|--------------|
| D13 | Nombre | Nombre completo del empleado |
| E13 | Día | Fecha del registro de asistencia |
| F13 | HoraEnt | Hora programada de entrada |
| G13 | HoraSal | Hora programada de salida |
| H13 | Marc-Ent | Hora real de marcaje de entrada |
| I13 | Marc-Sal | Hora real de marcaje de salida |

Los datos comienzan en la fila **14**.

---

## 🧩 Macro 1: Subir archivo de marcajes

### 📌 Nombre sugerido:
`Sub CargarArchivoMarcajes()`

### 🧠 Función:
Permite seleccionar y abrir un archivo externo de marcajes, leer su contenido y copiarlo automáticamente a la hoja principal, respetando el formato de columnas y encabezados.

### 🔍 Flujo de trabajo:
1. Aparece un cuadro de diálogo para elegir el archivo (`Application.GetOpenFilename`).
2. Abre el archivo seleccionado sin importar si está en formato `.xls`, `.xlsx`, o `.xlsm`.
3. Copia los datos desde la hoja de origen y los pega en la hoja principal (a partir de la celda D14).
4. Verifica que los encabezados coincidan con los esperados (Nombre, Día, HoraEnt, etc.).
5. Muestra un mensaje de confirmación al finalizar.

### ⚠️ Posibles errores:

| Error | Causa probable | Solución |
|--------|----------------|----------|
| “Estás intentando abrir un tipo de archivo bloqueado…” | Excel 2010 bloquea formatos antiguos | Desbloquear en **Centro de confianza → Configuración de bloqueo de archivos** o guardar el archivo como `.xlsx`. |
| “No se encontró el encabezado ‘Nombre’” | Los encabezados del archivo externo no coinciden exactamente | Verificar que los títulos sean idénticos: “Nombre”, “Día”, “HoraEnt”, “HoraSal”, “Marc-Ent”, “Marc-Sal”. |
| No copia los datos | La hoja fuente no tiene el formato esperado | Revisar que los datos inicien en la primera hoja del archivo de marcajes. |

---

## 🧮 Macro 2: Analizar marcajes y crear informes

### 📌 Nombre real:
`Sub AnalizarMarcajes_Auto_Mapeado()`

### 🧠 Función:
Analiza los marcajes cargados en la hoja activa, identifica registros incompletos (sin hora de entrada o salida), y genera **un informe de texto (.txt)** por cada empleado afectado.

Cada informe se guarda en una carpeta llamada **“Marcajes”** que se crea automáticamente en el escritorio.

### 🧩 Flujo de ejecución:

1. Verifica que existan encabezados en la fila 13.  
2. Crea la carpeta `Marcajes` en el escritorio si no existe.  
3. Recorre todas las filas con datos (desde la fila 14).  
4. Por cada empleado:
   - Si falta el marcaje de entrada o salida, crea un documento de texto.  
   - El nombre del archivo combina el nombre del empleado y la fecha del registro (ej. `Maria_Meneses_18-10-2025.txt`).  
5. Escribe dentro del archivo:
   - El nombre del empleado.  
   - El día del registro.  
   - La fecha de notificación (día actual).  
   - El motivo (ausencia completa, falta de entrada o de salida).  
6. Al finalizar, muestra un mensaje con el número total de informes generados.

### 📄 Ejemplo de contenido generado:

