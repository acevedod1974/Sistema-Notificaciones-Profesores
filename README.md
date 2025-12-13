# Sistema de Gestión Académica y Notificaciones (V6.1)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Un sistema avanzado basado en Google Apps Script para automatizar la gestión académica, el envío de reportes de progreso y el análisis de datos en tiempo real, integrado con Google Sheets, AppSheet y Telegram.

## Descripción

Este proyecto optimiza la comunicación del rendimiento académico. Permite a los profesores enviar correos electrónicos personalizados (HTML) y notificaciones de Telegram con un solo clic. A diferencia de sistemas tradicionales, este script incluye un **Dashboard Inteligente** que calcula el rendimiento real del estudiante basándose solo en las evaluaciones ya realizadas, evitando promedios engañosos al inicio del semestre.

---

### 📖 Guía Visual

Para una referencia visual, puedes consultar el manual base (nota: las funciones V6.1 como el Dashboard son nuevas y se explican abajo):

➡️ **[Ver la Guía de Usuario en PDF](./Gu%C3%ADa%20de%20Usuario_%20Asistente%20Autom%C3%A1tico%20de%20Calificaciones%20v3.1.pdf)**

---

## ✨ Características Principales (Versión 6.1)

### 🧠 Inteligencia y Análisis
-   **🖥️ Dashboard del Profesor:** Un panel de control visual (Google Charts) que muestra:
    -   **Velocímetro de Salud:** Estado general del grupo en tiempo real.
    -   **Gráfico de Barras:** Promedios por evaluación (coloreados dinámicamente según rendimiento).
    -   **Lista de Riesgo:** Tabla automática con estudiantes que tienen un rendimiento real < 50%.
-   **📊 Cálculo de Rendimiento Real:** El sistema detecta automáticamente qué evaluaciones ya han ocurrido y calcula el porcentaje del alumno sobre esa base (ej. *45/50 puntos evaluados = 90%*), en lugar de diluirlo sobre el total del semestre.

### 📧 Comunicación
-   **📝 Notas Personalizadas:** Escribe un mensaje específico en la **Columna P** (`MENSAJE_PERSONALIZADO`) y el sistema lo insertará automáticamente en el correo de ese estudiante como una "Nota del Profesor" destacada.
-   **📨 Reportes Duales:** Envía correos HTML detallados (con tablas y gráficos de progreso) y alertas instantáneas a **Telegram**.
-   **🔄 Modo Prueba Dinámico:** Activa o desactiva el envío de correos reales desde el menú `📧 Notificaciones` sin tocar el código.

### 🖨️ Utilidades
-   **Impresión de Reportes:** El Dashboard incluye una vista optimizada para imprimir o guardar como PDF limpio.

## 🛠️ Tecnologías Utilizadas

-   **Backend:** Google Apps Script
-   **Frontend:** HTML5 / CSS3 (para correos y dashboard)
-   **Datos:** Google Sheets
-   **Visualización:** Google Charts API
-   **Mensajería:** Gmail API & Telegram Bot API

## 🚀 Guía de Instalación Rápida

1.  **Obtener la Plantilla**
    * <a href="https://docs.google.com/spreadsheets/d/1C_5Hez9VQD8Uv5zTGe6BLQTV1LOMx80oVF23oQ4A8YA/copy" target="_blank">HAZ CLIC AQUÍ PARA CREAR TU PROPIA COPIA DE LA HOJA DE PLANTILLA</a>

2.  **Configuración Inicial**
    * Abre tu copia y ve al menú `Extensiones > Apps Script`.
    * En `CONFIGURACIÓN PRINCIPAL`, ajusta tu nombre, asignatura y verifica los mapeos de columnas si cambias el diseño de la hoja.

3.  **Configurar Telegram (Opcional pero recomendado)**
    * Refresca la hoja de cálculo (F5).
    * Ve al menú `📧 Notificaciones > 🔐 Configurar Token Telegram`.
    * Pega el Token de tu bot.

## 📋 Cómo Usar las Nuevas Funciones

### 1. Ver Estadísticas del Grupo
Ve al menú `📧 Notificaciones > 📊 Ver Estadísticas del Grupo`. Se abrirá una ventana emergente con los gráficos de rendimiento y la lista de alumnos en riesgo. Puedes usar el botón "Imprimir" para generar un PDF del estado actual.

### 2. Enviar Mensajes Personalizados
Si quieres decirle algo específico a un alumno (ej. *"Excelente mejora en el ensayo"*):
1.  Ve a la columna **P** (`MENSAJE_PERSONALIZADO`).
2.  Escribe tu mensaje en la fila del estudiante.
3.  Al enviar el reporte (Individual o Semestral), este texto aparecerá en un recuadro amarillo destacado dentro del correo.

### 3. Modo Prueba
Usa el menú `📧 Notificaciones > 🔄 Alternar Modo Prueba` para cambiar entre:
* 🟢 **ACTIVO:** Los correos llegan a TI (para verificar que todo se ve bien).
* 🔴 **INACTIVO:** Los correos se envían a los ESTUDIANTES reales.

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## ✨ Agradecimientos

-   **Creado por:** Daniel Acevedo
-   **Desarrollo Asistido por:** Google Gemini
