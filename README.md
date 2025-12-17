# Asistente Automático de Calificaciones para Profesores

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Un sistema basado en Google Apps Script diseñado para automatizar el envío de reportes de progreso a estudiantes, integrándose con Google Sheets, AppSheet y Looker Studio para un flujo de trabajo completo.

## Descripción

Este proyecto nace de la necesidad de optimizar la comunicación del rendimiento académico a los estudiantes. El sistema permite a los profesores, con una configuración inicial mínima, enviar correos electrónicos personalizados y detallados que incluyen tablas de notas, gráficos de progreso y mensajes contextuales basados en el desempeño del estudiante.

## ✨ Características Principales

- **📧 Notificaciones Personalizadas:** Envía correos únicos a cada estudiante con su progreso detallado.
- **📊 Reportes Visuales:** Incluye en cada correo una tabla con el historial de notas y un gráfico de barras comparativo.
- **🕹️ Control Total del Envío:** Utiliza un menú personalizado en Google Sheets y un botón en AppSheet para que el profesor decida exactamente cuándo enviar las notificaciones.
- **📱 Integración con AppSheet:** Permite crear una aplicación móvil/web para una entrada de datos más amigable y para activar envíos individuales sobre la marcha.
- **📈 Reporte Semestral:** Una función para enviar un correo masivo a toda la clase con un análisis de su proyección final y mensajes de motivación o alerta.
- **🤖 Integración con Telegram:** Envío opcional de notificaciones instantáneas a través de un Bot de Telegram.

## 🛠️ Estructura del Código (Refactorizado)

El proyecto ha sido modularizado para facilitar su mantenimiento:

- `Config.gs`: Contiene todas las constantes de configuración (columnas, nombres, umbrales) y gestión de propiedades.
- `UI.gs`: Maneja el menú personalizado, los disparadores (`onEdit`, `doPost`) y el Dashboard de estadísticas.
- `Email.gs`: Contiene la lógica central para el envío de correos individuales y masivos.
- `Helpers.gs`: Funciones de utilidad para cálculos matemáticos, formateo de números y generación de componentes HTML.

## 🚀 Guía de Instalación (Getting Started)

Para utilizar este sistema, solo necesitas hacer una copia de la plantilla de Google Sheets, que ya contiene todo el código del asistente configurado.

1.  **Crear tu Propia Copia de la Hoja**
    * <a href="https://docs.google.com/spreadsheets/d/1C_5Hez9VQD8Uv5zTGe6BLQTV1LOMx80oVF23oQ4A8YA/copy" target="_blank">HAZ CLIC AQUÍ PARA CREAR TU PROPIA COPIA DE LA HOJA DE PLANTILLA</a>

2.  **Configurar el Script**
    * Una vez que tengas tu copia, ábrela y ve al menú `Extensiones > Apps Script`.
    * Abre el archivo `Config.gs`.
    * Modifica las variables `NOMBRE_PROFESOR`, `NOMBRE_ASIGNATURA`, `EMAIL_PROFESOR`, y verifica que los números de las columnas (`COL_...`) coincidan con tu hoja.
    * **Importante:** Configura `MODO_PRUEBA: true` y tu `EMAIL_PRUEBA` para realizar pruebas de forma segura.

3.  **Autorizar el Script**
    * Refresca tu hoja de cálculo. Aparecerá un nuevo menú `📧 Notificaciones`.
    * La primera vez que hagas clic en una de sus opciones (ej. `Enviar Correos Pendientes`), Google te pedirá que autorices el script. Sigue los pasos y concede los permisos necesarios.

## 📋 Uso

El sistema tiene dos flujos de trabajo principales:

- **Envío por Lote (desde Google Sheets):** Introduce las calificaciones en la hoja. Las filas modificadas se marcarán como "Pendiente". Usa el menú `📧 Notificaciones > Enviar Correos Pendientes` para enviar todos los reportes de una vez.
- **Envío Individual (desde AppSheet):** Crea una app desde `Extensiones > AppSheet`. Configura un `Bot` en la sección `Automation` para que llame al script como un `webhook`. Esto te permitirá enviar notificaciones individuales desde un botón en la app de tu teléfono.

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## ✨ Agradecimientos

-   **Creado por:** Daniel Acevedo
-   **Asistente de Desarrollo IA:** Google Gemini
-   **Desarrollo Asistido por:** Google Gemini
