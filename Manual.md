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
