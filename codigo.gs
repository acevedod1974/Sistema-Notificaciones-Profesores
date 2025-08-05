/**
 * @license
 * Copyright 2025 Daniel Acevedo
 *
 * Sistema Automatizado de Notificaciones para Calificaciones
 * Asistente de Desarrollo IA: Google Gemini
 * * Versión con integración de Telegram.
 */

// =================================================================
//                CONFIGURACIÓN PRINCIPAL
// =================================================================
const CONFIG = {
  NOMBRE_PROFESOR: "Nombre Profesor",
  NOMBRE_ASIGNATURA: "Nombre Asignatura",
  EMAIL_PROFESOR: "emailprofesor@unexpo.edu.ve",
  TELEGRAM_BOT_TOKEN: "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxs", // <-- ¡NUEVO! Token de tu bot.
  
  NOMBRE_HOJA: "Hoja 1",
  FILA_ENCABEZADOS: 1,
  FILA_NOTA_MAXIMA: 2,
  FILA_INICIO_ESTUDIANTES: 3,
  
  COL_NOMBRES: 3,
  COL_APELLIDOS: 2,
  COL_EMAIL: 4, 
  COL_SUMA_100: 12, 
  COL_ESTADO_CORREO: 14,
  COL_TELEGRAM_ID: 15, // <-- ¡NUEVO! Asegúrate de que este sea el número de tu nueva columna.
  
  COLUMNAS_CALIFICACIONES: [5, 6, 7, 8, 9, 10, 11],
  COLUMNAS_RESTANTES: [9, 10, 11], 
  
  UMBRAL_APROBACION: 50,
  UMBRAL_BUEN_ESTADO: 65,
  
  MODO_PRUEBA: true, 
  EMAIL_PRUEBA: "emailprueba@gmail.com"
};

// ... [Las funciones onOpen, onEdit, doPost, enviarCorreosPendientes no cambian] ...
function onOpen() {SpreadsheetApp.getUi().createMenu('📧 Notificaciones').addItem('Enviar Correos Pendientes', 'enviarCorreosPendientes').addSeparator().addItem('Enviar Reporte Semestral a Todos', 'enviarReporteSemestral').addToUi();}
function onEdit(e) {const rango = e.range;const hoja = e.source.getActiveSheet();const filaEditada = rango.getRow();const colEditada = rango.getColumn();if (hoja.getName() !== CONFIG.NOMBRE_HOJA || filaEditada < CONFIG.FILA_INICIO_ESTUDIANTES || !CONFIG.COLUMNAS_CALIFICACIONES.includes(colEditada)) { return; }const valorEditado = rango.getDisplayValue();const celdaEstado = hoja.getRange(filaEditada, CONFIG.COL_ESTADO_CORREO);if (valorEditado === "") { rango.setBackground(null); celdaEstado.setValue('Actualizado'); return; }celdaEstado.setValue(`Pendiente:${rango.getA1Notation()}`);const valorEstandarizado = valorEditado.replace(',', '.');if (isNaN(parseFloat(valorEstandarizado))) { return; }const notaObtenida = parseFloat(valorEstandarizado);const notaMaximaEvaluacion = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue());if (isNaN(notaMaximaEvaluacion) || notaMaximaEvaluacion <= 0) { return; }if (notaObtenida > notaMaximaEvaluacion) { rango.setBackground('#FFC0CB'); } else { rango.setBackground(null); }}
function doPost(e) {Logger.log("Llamada POST recibida de AppSheet.");try {const postData = JSON.parse(e.postData.contents);const numeroDeFila = postData.fila;Logger.log("Fila recibida: " + numeroDeFila);if (!numeroDeFila) {return ContentService.createTextOutput("Error: No se proporcionó el número de fila.");}const resultado = enviarCorreoAUnaFila(numeroDeFila, null);return ContentService.createTextOutput(resultado);} catch (error) {Logger.log("Error en doPost: " + error.toString());return ContentService.createTextOutput("Error en el script: " + error.toString());}}
function enviarCorreosPendientes() {const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) {SpreadsheetApp.getUi().alert('No hay estudiantes para procesar.');return;}const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());const todosLosDatos = rangoDatos.getValues();let correosEnviados = 0;todosLosDatos.forEach((datosFila, index) => {const estadoCorreo = datosFila[CONFIG.COL_ESTADO_CORREO - 1].toString();const filaActual = CONFIG.FILA_INICIO_ESTUDIANTES + index;if (estadoCorreo.startsWith('Pendiente:')) {enviarCorreoAUnaFila(filaActual, estadoCorreo);correosEnviados++;}});SpreadsheetApp.getUi().alert('Proceso finalizado', `Se han procesado y/o enviado ${correosEnviados} correos pendientes.`, SpreadsheetApp.getUi().ButtonSet.OK);}


// =================================================================
//    FUNCIÓN CENTRAL DE ENVÍO (MODIFICADA PARA LLAMAR A TELEGRAM)
// =================================================================
function enviarCorreoAUnaFila(numeroDeFila, estadoCorreo) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const filaActual = parseInt(numeroDeFila);
  const datosFila = hoja.getRange(filaActual, 1, 1, hoja.getLastColumn()).getValues()[0];
  const promediosPorColumna = calcularPromedios(hoja);
  
  const nombreEstudiante = `${datosFila[CONFIG.COL_NOMBRES - 1]} ${datosFila[CONFIG.COL_APELLIDOS - 1]}`;
  const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1];
  const telegramChatId = datosFila[CONFIG.COL_TELEGRAM_ID - 1]; // <-- Obtener el Chat ID
  
  if (!emailEstudiante) {
    hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue('Error: Falta email');
    return "Error: Falta email";
  }

  // ... [Lógica para obtener notaObtenida, nombreEvaluacion, etc. - Sin cambios] ...
  let notaObtenida, notaMaximaEvaluacion, nombreEvaluacion;
  if (estadoCorreo && estadoCorreo.startsWith('Pendiente:')) {
    const a1Notation = estadoCorreo.split(':')[1];
    const celdaEditada = hoja.getRange(a1Notation);
    notaObtenida = celdaEditada.getValue();
    const colEditada = celdaEditada.getColumn();
    notaMaximaEvaluacion = hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue();
    nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, colEditada).getValue();
  } else {
     CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
      const valorCelda = datosFila[col-1];
       if (valorCelda !== "" && !isNaN(parseFloat(valorCelda.toString().replace(',','.')))) {
         notaObtenida = parseFloat(valorCelda.toString().replace(',','.'));
         notaMaximaEvaluacion = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
         nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
       }
     });
  }
  const suma100 = parseFloat(datosFila[CONFIG.COL_SUMA_100 - 1]);
  
  let asunto = "";
  let textoIntroductorio = "";
  let botonHTML = "";
  let mensajeTelegram = "";

  // ... [Lógica de selección de mensaje - Sin cambios en el texto del email] ...
  if (notaObtenida / notaMaximaEvaluacion < CONFIG.UMBRAL_NOTA_BAJA) {
    asunto = `Importante: Revisión sobre tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
    textoIntroductorio = `Hola ${nombreEstudiante},<br><br>Te escribo para revisar tu calificación de <b>${notaObtenida.toFixed(2)}</b> en "<b>${nombreEvaluacion}</b>"...`; // (Texto abreviado)
    const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=Solicitud%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
    botonHTML = crearBotonCtaHTML("Agendar Asesoría por Correo", enlaceMailto);
    mensajeTelegram = `Hola ${nombreEstudiante}. Se ha registrado tu nota en *${nombreEvaluacion}*: ${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}. Tu acumulado es ${suma100.toFixed(2)}/100. He notado que esta nota es baja, si necesitas ayuda no dudes en consultarme.`;

  } else { // (Otros casos de mensaje)
    asunto = `Actualización de tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
    textoIntroductorio = `Estimado ${nombreEstudiante},<br><br>Este correo es para confirmar que tu calificación en "<b>${nombreEvaluacion}</b>" fue de <b>${notaObtenida.toFixed(2)}</b>...`; // (Texto abreviado)
    mensajeTelegram = `¡Hola ${nombreEstudiante}! Se ha registrado tu nota en *${nombreEvaluacion}*: *${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}*. Tu acumulado actual es ${suma100.toFixed(2)}/100.`;
  }
  
  // -- INICIO DE LA INTEGRACIÓN --
  // Enviar notificación por Telegram si existe un Chat ID
  if (telegramChatId && telegramChatId.toString().trim() !== '') {
    enviarNotificacionTelegram(telegramChatId, mensajeTelegram);
  }
  // -- FIN DE LA INTEGRACIÓN --

  const encabezadoHTML = crearEncabezadoEmailHTML();
  const tablaHTML = crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna);
  const graficoBlob = crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna);
  let cuerpoFinalHTML = `<div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; color: #333;">${encabezadoHTML}<div style="padding: 20px;"><p>${textoIntroductorio}</p>${botonHTML}${tablaHTML}<p>Para una mejor visualización de tu avance, puedes ver el siguiente gráfico:</p><p style="text-align:center;"><img src="cid:graficoProgreso"></p><p>Saludos,<br><b>${CONFIG.NOMBRE_PROFESOR}</b><br>Profesor de ${CONFIG.NOMBRE_ASIGNATURA}</p></div></div>`;
  const opcionesEmail = { htmlBody: cuerpoFinalHTML };
  if (graficoBlob) { opcionesEmail.inlineImages = { graficoProgreso: graficoBlob }; }

  try {
    GmailApp.sendEmail(emailEstudiante, asunto, "", opcionesEmail);
    hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue(`Enviado ${new Date().toLocaleString()}`);
    return "Correo enviado exitosamente a la fila " + numeroDeFila;
  } catch (error) {
    Logger.log("ERROR al enviar correo: " + error.toString());
    hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue(`Error al enviar: ${error.message}`);
    return "Error al enviar correo: " + error.message;
  }
}

// ... [La función de Reporte Semestral no se modifica por ahora] ...
function enviarReporteSemestral(){ /* ... */ }
function analizarYEnviarCorreoSemestral(){ /* ... */ }

// =================================================================
//          NUEVA FUNCIÓN AYUDANTE PARA TELEGRAM
// =================================================================
function enviarNotificacionTelegram(chatId, mensaje) {
  const token = CONFIG.TELEGRAM_BOT_TOKEN;
  if (!token || token === "PEGA_AQUÍ_EL_TOKEN_DE_TU_BOT") {
    Logger.log("Token de Telegram no configurado en el script. No se envió la notificación.");
    return;
  }
  
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({
      'chat_id': String(chatId),
      'text': mensaje,
      'parse_mode': 'Markdown' // Permite usar *negritas* y _cursivas_
    })
  };

  try {
    UrlFetchApp.fetch(url, payload);
    Logger.log(`Notificación de Telegram enviada a Chat ID: ${chatId}`);
  } catch (e) {
    Logger.log(`Error al enviar notificación de Telegram a ${chatId}: ${e.toString()}`);
  }
}


// ... [El resto de las funciones ayudantes (crearEncabezadoEmailHTML, crearBotonCtaHTML, etc.) van aquí sin cambios] ...
function crearEncabezadoEmailHTML(){return `<div style="background-color: #f8f9fa; padding: 20px; border-bottom: 5px solid #005A9C; text-align: center;"><h2 style="color: #005A9C; margin-top: 10px; margin-bottom: 0;">Reporte de Progreso Académico</h2><p style="color: #6c757d; margin-top: 5px; margin-bottom: 0;">${CONFIG.NOMBRE_ASIGNATURA}</p></div>`;}
function crearBotonCtaHTML(texto, enlace){return `<table border="0" cellpadding="0" cellspacing="0" style="margin: 20px 0;"><tr><td align="center" bgcolor="#005A9C" style="border-radius: 5px;"><a href="${enlace}" target="_blank" style="font-size: 16px; font-family: Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 5px; padding: 12px 25px; border: 1px solid #005A9C; display: inline-block;">${texto}</a></td></tr></table>`;}
function calcularPromedios(hoja){const promedios = {};const ultimaFila = hoja.getLastRow();if (ultimaFila < CONFIG.FILA_INICIO_ESTUDIANTES) { CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => promedios[col] = 0); return promedios; }const rangoEstudiantes = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, ultimaFila - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());const datosEstudiantes = rangoEstudiantes.getValues();CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {let suma = 0;let contador = 0;datosEstudiantes.forEach(fila => {const nota = fila[col - 1];if (nota !== "" && !isNaN(nota)) { suma += parseFloat(nota); contador++; }});promedios[col] = (contador > 0) ? (suma / contador) : 0;});return promedios;}
function crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna){let tablaHTML = `<p>Para darte un contexto completo, aquí tienes un resumen detallado de tu progreso:</p><table style="width:100%; border-collapse: collapse; border: 1px solid #dee2e6; font-family: Arial, sans-serif; font-size: 14px;"><thead style="background-color: #005A9C; color: #ffffff;"><tr><th style="padding: 12px; border: 1px solid #005A9C; text-align: left;">Evaluación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Tu Calificación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Nota Máxima</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Promedio del Grupo</th></tr></thead><tbody>`;let rowIndex = 0;let puntajeAcumulado = 0;let puntajeMaximoPosible = 0;CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {const notaEstudiante = datosFila[col - 1];if (notaEstudiante !== "" && !isNaN(notaEstudiante)) {const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();const notaMaxima = hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue();const promedioGrupo = promediosPorColumna[col];puntajeAcumulado += parseFloat(notaEstudiante);puntajeMaximoPosible += parseFloat(notaMaxima);const rowColor = (rowIndex % 2 === 0) ? '#ffffff' : '#f8f9fa';tablaHTML += `<tr style="background-color: ${rowColor};"><td style="padding: 12px; border: 1px solid #dee2e6;">${nombreEvaluacion}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;"><b>${parseFloat(notaEstudiante).toFixed(2)}</b></td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${parseFloat(notaMaxima).toFixed(2)}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${promedioGrupo.toFixed(2)}</td></tr>`;rowIndex++;}});tablaHTML += `</tbody><tfoot style="background-color: #005A9C; color: #ffffff; font-weight: bold;"><tr><td style="padding: 12px; border: 1px solid #005A9C;">Total Acumulado Parcial</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeAcumulado.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeMaximoPosible.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">-</td></tr></tfoot></table>`;return tablaHTML;}
function crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna){const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Evaluación").addColumn(Charts.ColumnType.NUMBER, "Tu Calificación").addColumn(Charts.ColumnType.NUMBER, "Promedio del Grupo");let hayDatos = false;CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {const notaEstudiante = datosFila[col - 1];if (notaEstudiante !== "" && !isNaN(notaEstudiante)) {const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();const promedioGrupo = promediosPorColumna[col];dataTable.addRow([nombreEvaluacion, parseFloat(notaEstudiante), promedioGrupo]);hayDatos = true;}});if (!hayDatos) return null;const chart = Charts.newColumnChart().setTitle("Tu Progreso vs. el Promedio del Grupo").setDataTable(dataTable).setOption('legend', { position: 'top' }).setOption('vAxis', { title: 'Calificación', minValue: 0 }).setDimensions(600, 400).build();return chart.getAs('image/png');}
