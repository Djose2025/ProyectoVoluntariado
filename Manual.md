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
| **Permisos de usuario** | Permitir macros y
