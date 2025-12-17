/**
 * ARCHIVO: Config.gs
 * RESPONSABILIDAD: Configuración global y gestión de propiedades del script.
 */

function obtenerModoPrueba() {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty('MODO_PRUEBA');
    return prop === 'false' ? false : true; 
  } catch (e) { return true; }
}

const CONFIG = {
  NOMBRE_PROFESOR: "Daniel Acevedo",
  NOMBRE_ASIGNATURA: "Procesos de Fabricacion 1",
  EMAIL_PROFESOR: "dacevedo@unexpo.edu.ve",
  
  NOMBRE_HOJA: "Hoja 1",
  FILA_ENCABEZADOS: 1,
  FILA_NOTA_MAXIMA: 2,
  FILA_INICIO_ESTUDIANTES: 3,
  
  // MAPEO DE COLUMNAS
  COL_NOMBRES: 3,
  COL_APELLIDOS: 2,
  COL_EMAIL: 4, 
  
  COLUMNAS_CALIFICACIONES: [5, 6, 7, 8, 9, 10],
  COLUMNAS_RESTANTES: [8, 9, 10], 
  
  COL_SUMA_100: 11, 
  COL_ESTADO_CORREO: 13, 
  COL_TELEGRAM_ID: 14,
  COL_MENSAJE_PERSONALIZADO: 15, 
  
  UMBRAL_APROBACION: 50,
  UMBRAL_BUEN_ESTADO: 65,
  UMBRAL_EXAMEN_ALTO: 0.80, 
  UMBRAL_NOTA_BAJA: 0.50, 
  UMBRAL_PROMEDIO_ALTO: 80, 
  UMBRAL_RIESGO: 60,
  
  MODO_PRUEBA: obtenerModoPrueba(), 
  EMAIL_PRUEBA: "acevedod1974@gmail.com"
};

// --- Funciones de Configuración de Usuario ---

function alternarModoPrueba() {
  const estadoActual = obtenerModoPrueba();
  const nuevoEstado = !estadoActual;
  PropertiesService.getScriptProperties().setProperty('MODO_PRUEBA', nuevoEstado.toString());
  const estadoTexto = nuevoEstado ? "🟢 ACTIVADO (Correos a ti)" : "🔴 DESACTIVADO (Correos a alumnos)";
  SpreadsheetApp.getActive().toast(`Modo Prueba: ${estadoTexto}`, 'Configuración Actualizada');
  // Nota: onOpen no se puede llamar directamente desde aquí para refrescar menú sin recargar hoja, 
  // pero la propiedad queda guardada.
}

function guardarTokenTelegram() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt('Configuración de Telegram', 'Pega aquí el Token de tu Bot de Telegram:', ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    const token = respuesta.getResponseText().trim();
    if (token) {
      PropertiesService.getScriptProperties().setProperty('TELEGRAM_BOT_TOKEN', token);
      ui.alert('Éxito', 'El token se ha guardado de forma segura.', ui.ButtonSet.OK);
    } else { ui.alert('Error', 'El token no puede estar vacío.', ui.ButtonSet.OK); }
  }
}
