/**
 * ARCHIVO: Email.gs
 * RESPONSABILIDAD: Lógica de negocio para envíos de correo (Individual y Masivo).
 */

function enviarCorreosPendientes() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  
  if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) {
    SpreadsheetApp.getActive().toast('No hay estudiantes para procesar.', '⚠️ Aviso');
    return;
  }

  try {
    const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
    const todosLosDatos = rangoDatos.getValues();
    
    // OPTIMIZACIÓN: Calculamos promedios UNA SOLA VEZ antes del bucle
    const promediosPrecalculados = calcularPromedios(hoja);
    
    let correosEnviados = 0;
    
    todosLosDatos.forEach((datosFila, index) => {
      const estadoCorreo = datosFila[CONFIG.COL_ESTADO_CORREO - 1].toString();
      const filaActual = CONFIG.FILA_INICIO_ESTUDIANTES + index;
      
      if (estadoCorreo.startsWith('Pendiente:')) {
        enviarCorreoAUnaFila(filaActual, estadoCorreo, promediosPrecalculados, hoja, datosFila);
        correosEnviados++;
      }
    });
    
    SpreadsheetApp.flush(); 
    SpreadsheetApp.getActive().toast(`Se enviaron ${correosEnviados} correos pendientes.`, '✅ Proceso Finalizado', 5);
    
  } catch (error) {
    Logger.log("Error crítico: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function enviarCorreoAUnaFila(numeroDeFila, estadoCorreo, promediosInyectados = null, hojaInyectada = null, datosFilaInyectados = null) {
  const hoja = hojaInyectada || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const filaActual = parseInt(numeroDeFila);
  const datosFila = datosFilaInyectados || hoja.getRange(filaActual, 1, 1, hoja.getLastColumn()).getValues()[0];
  const promediosPorColumna = promediosInyectados || calcularPromedios(hoja); // Usa helper
  
  const nombreEstudiante = `${datosFila[CONFIG.COL_NOMBRES - 1]} ${datosFila[CONFIG.COL_APELLIDOS - 1]}`;
  const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1]; 
  const telegramChatId = datosFila[CONFIG.COL_TELEGRAM_ID - 1]; 
  
  // Refactorizado: Usa helper para HTML repetitivo
  const mensajePersonalizado = datosFila[CONFIG.COL_MENSAJE_PERSONALIZADO - 1];
  const htmlMensajePersonalizado = construirNotaPersonalizadaHTML(mensajePersonalizado);

  if (!emailEstudiante) {
    hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue('Error: Falta email');
    return "Error: Falta email";
  }

  let notaObtenida, notaMaximaEvaluacion, nombreEvaluacion;
  
  // Lógica de detección de evaluación
  if (estadoCorreo && estadoCorreo.startsWith('Pendiente:')) {
    const a1Notation = estadoCorreo.split(':')[1];
    const celdaEditada = hoja.getRange(a1Notation);
    notaObtenida = toNumero(celdaEditada.getValue()); // Refactorizado
    const colEditada = celdaEditada.getColumn();
    notaMaximaEvaluacion = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue()); // Refactorizado
    nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, colEditada).getValue();
  } else {
     CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
       const valorCelda = datosFila[col-1];
       const valorNumerico = toNumero(valorCelda); // Refactorizado
       if (!isNaN(valorNumerico)) {
         notaObtenida = valorNumerico;
         notaMaximaEvaluacion = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
         nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
       }
     });
  }
  
  const suma100 = toNumero(datosFila[CONFIG.COL_SUMA_100 - 1]); // Refactorizado
  let puntajeAcumulado = 0;
  let puntajeMaximoPosible = 0;
  
  CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
    const valorNumerico = toNumero(datosFila[col - 1]); // Refactorizado
    if (!isNaN(valorNumerico)) {
      puntajeAcumulado += valorNumerico;
      puntajeMaximoPosible += toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
    }
  });

  if (puntajeMaximoPosible === 0) {
    hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue('Error: Sin notas');
    return "Error: Sin notas";
  }
  
  const rendimientoNormalizado = puntajeAcumulado / puntajeMaximoPosible * 100;
  const rendimientoExamen = notaObtenida / notaMaximaEvaluacion;
  
  let asunto = "", textoIntroductorio = "", botonHTML = "", mensajeTelegram = "";

  if (rendimientoExamen < CONFIG.UMBRAL_NOTA_BAJA || (rendimientoNormalizado < CONFIG.UMBRAL_RIESGO && rendimientoExamen < CONFIG.UMBRAL_EXAMEN_ALTO) ) {
    asunto = `Importante: Revisión sobre tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
    textoIntroductorio = `Hola ${nombreEstudiante},<br><br>Te escribo para revisar tu calificación de <b>${notaObtenida.toFixed(2)}</b> en "<b>${nombreEvaluacion}</b>". He notado que tu rendimiento requiere atención en esta evaluación.<br><br>Es un momento importante para corregir el rumbo. Estoy aquí para ayudarte. Por favor, agenda una asesoría.`;
    const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=Solicitud%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
    botonHTML = crearBotonCtaHTML("Agendar Asesoría por Correo", enlaceMailto);
    mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. Se ha registrado tu nota en *${nombreEvaluacion}*: ${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)} (Baja). Tu acumulado es ${suma100.toFixed(2)}/100. Contáctame si necesitas ayuda.`;
  } else if (rendimientoExamen >= CONFIG.UMBRAL_EXAMEN_ALTO) {
    asunto = `¡Excelente trabajo en ${nombreEvaluacion}, ${nombreEstudiante}!`;
    textoIntroductorio = `¡Felicidades, ${nombreEstudiante}!<br><br>Quería destacar especialmente tu calificación de <b>${notaObtenida.toFixed(2)}</b> en "<b>${nombreEvaluacion}</b>". ¡Un resultado sobresaliente!`;
    mensajeTelegram = `🌟 ¡Hola ${nombreEstudiante}! Excelente nota en *${nombreEvaluacion}*: *${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}*. Tu acumulado es ${suma100.toFixed(2)}/100. ¡Sigue así!`;
  } else {
    asunto = `Actualización de tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
    textoIntroductorio = `Estimado ${nombreEstudiante},<br><br>Este correo es para confirmar que tus calificaciones han sido actualizadas. Tu rendimiento en "<b>${nombreEvaluacion}</b>" fue de <b>${notaObtenida.toFixed(2)}</b>.`;
    mensajeTelegram = `📝 Hola ${nombreEstudiante}. Se ha actualizado tu nota en *${nombreEvaluacion}*: ${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}. Tu acumulado es ${suma100.toFixed(2)}/100.`;
  }
  
  if (telegramChatId && telegramChatId.toString().trim() !== '') {
    let msgTgFinal = mensajeTelegram;
    if (mensajePersonalizado && String(mensajePersonalizado).trim() !== "") {
      msgTgFinal += `\n\n📝 *Nota del Profesor:* ${mensajePersonalizado}`;
    }
    enviarNotificacionTelegram(telegramChatId, msgTgFinal); // Usa Helper
  }
  
  const encabezadoHTML = crearEncabezadoEmailHTML(); // Usa Helper
  const tablaHTML = crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna); // Usa Helper
  const graficoBlob = crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna); // Usa Helper
  
  let cuerpoFinalHTML = `<div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; color: #333;">${encabezadoHTML}<div style="padding: 20px;"><p>${textoIntroductorio}</p>${htmlMensajePersonalizado}${botonHTML}${tablaHTML}<p>Para una mejor visualización de tu avance, puedes ver el siguiente gráfico:</p><p style="text-align:center;"><img src="cid:graficoProgreso"></p><p>Saludos,<br><b>${CONFIG.NOMBRE_PROFESOR}</b><br>Profesor de ${CONFIG.NOMBRE_ASIGNATURA}</p></div></div>`;
  
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

function enviarReporteSemestral() {
  const ui = SpreadsheetApp.getUi();
  const confirmacion = ui.alert('Confirmación de Envío Masivo', 'Estás a punto de enviar un reporte de progreso a TODOS los estudiantes. ¿Deseas continuar?', ui.ButtonSet.YES_NO);
  if (confirmacion !== ui.Button.YES) { return; }

  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
    if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) { 
      SpreadsheetApp.getActive().toast('No hay estudiantes para procesar.', 'Aviso'); return; 
    }
    const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
    const todosLosDatos = rangoDatos.getValues();
    let correosEnviados = 0;

    const promediosPorColumna = calcularPromedios(hoja); // Usa Helper
    
    let puntajeMaximoRestante = 0;
    CONFIG.COLUMNAS_RESTANTES.forEach(col => {
      const val = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
      if (!isNaN(val)) puntajeMaximoRestante += val;
    });

    todosLosDatos.forEach((datosFila) => {
      const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1]; 
      if (emailEstudiante) {
        analizarYEnviarCorreoSemestral(hoja, datosFila, puntajeMaximoRestante, promediosPorColumna);
        correosEnviados++;
      }
    });
    
    SpreadsheetApp.flush(); 
    SpreadsheetApp.getActive().toast(`Se han enviado ${correosEnviados} reportes semestrales.`, '✅ Proceso Finalizado', 5);
    
  } catch (error) {
    Logger.log("Error en reporte semestral: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

function analizarYEnviarCorreoSemestral(hoja, datosFila, puntajeMaximoRestante, promediosPorColumna) {
  const nombreEstudiante = `${datosFila[CONFIG.COL_NOMBRES - 1]} ${datosFila[CONFIG.COL_APELLIDOS - 1]}`;
  const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1];
  const telegramChatId = datosFila[CONFIG.COL_TELEGRAM_ID - 1];
  const mensajePersonalizado = datosFila[CONFIG.COL_MENSAJE_PERSONALIZADO - 1];

  const htmlMensajePersonalizado = construirNotaPersonalizadaHTML(mensajePersonalizado); // Usa Helper
  
  const puntajeAcumuladoActual = toNumero(datosFila[CONFIG.COL_SUMA_100 - 1]); // Refactorizado
  const proyeccionFinalMaxima = puntajeAcumuladoActual + puntajeMaximoRestante;

  let puntajeTotalDelCurso = 0;
  CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
    const val = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
    if(!isNaN(val)) puntajeTotalDelCurso += val;
  });
  
  const puntajeMaximoEvaluado = puntajeTotalDelCurso - puntajeMaximoRestante;
  const rendimientoActual = (puntajeMaximoEvaluado > 0) ? (puntajeAcumuladoActual / puntajeMaximoEvaluado) * 100 : 0;

  let asunto = "", cuerpoMensaje = "", botonHTML = "", mensajeTelegram = "";

  if (proyeccionFinalMaxima < CONFIG.UMBRAL_APROBACION) {
    asunto = `URGENTE: Reunión sobre tu futuro en ${CONFIG.NOMBRE_ASIGNATURA}`;
    cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo con urgencia sobre tu situación. El análisis indica que, <b>incluso obteniendo la máxima calificación en lo que resta, la proyección no alcanza el mínimo aprobatorio.</b></p><p>Es muy importante que nos reunamos esta semana para discutir tu caso y explorar opciones.</p>`;
    const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=URGENTE:%20Reunión%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
    botonHTML = crearBotonCtaHTML("Contactar al Profesor para Agendar Reunión", enlaceMailto);
    mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. URGENTE: Tu proyección actual no alcanza para aprobar la materia. Por favor contáctame urgente.`;

  } else if (rendimientoActual < CONFIG.UMBRAL_BUEN_ESTADO) {
    asunto = `Información Importante sobre tu Situación en ${CONFIG.NOMBRE_ASIGNATURA}`;
    cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo para conversar sobre tu situación académica. He realizado un análisis y la buena noticia es que <b>todavía es posible que alcances la calificación final aprobatoria de ${CONFIG.UMBRAL_APROBACION} puntos.</b></p><p>Para lograrlo, se requiere un rendimiento excepcional en las evaluaciones restantes. Si quieres que conversemos para trazar un plan, no dudes en contactarme.</p>`;
    const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=Solicitud%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
    botonHTML = crearBotonCtaHTML("Solicitar Asesoría", enlaceMailto);
    mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. Reporte Semestral: Tu situación es delicada pero salvable. Aún puedes aprobar con esfuerzo máximo. Contáctame si necesitas asesoría.`;
  } else {
    asunto = `Reconocimiento de tu Progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
    cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo para darte un reconocimiento por tu desempeño y esfuerzo. Tus resultados hasta ahora son sólidos y te posicionan favorablemente para el cierre del curso. ¡Sigue así!</p>`;
    mensajeTelegram = `🌟 Hola ${nombreEstudiante}. Reporte Semestral: ¡Felicidades! Vas por excelente camino. Mantén el esfuerzo.`;
  }
  
  if (telegramChatId && telegramChatId.toString().trim() !== '') {
    let msgTgFinal = mensajeTelegram;
    if (mensajePersonalizado && String(mensajePersonalizado).trim() !== "") {
      msgTgFinal += `\n\n📝 *Nota del Profesor:* ${mensajePersonalizado}`;
    }
    enviarNotificacionTelegram(telegramChatId, msgTgFinal); // Usa Helper
  }

  const tablaHTML = crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna); // Usa Helper
  const graficoBlob = crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna); // Usa Helper
  
  let cuerpoFinalHTML = `<div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; color: #333;">${crearEncabezadoEmailHTML()}<div style="padding:20px;"><p>${cuerpoMensaje}</p>${htmlMensajePersonalizado}${tablaHTML}${botonHTML}<p>Para una mejor visualización de tu avance, puedes ver el siguiente gráfico:</p><p style="text-align:center;"><img src="cid:graficoProgreso"></p><p>Saludos,<br><b>${CONFIG.NOMBRE_PROFESOR}</b></p></div></div>`;
  const opcionesEmail = { htmlBody: cuerpoFinalHTML };
  if (graficoBlob) { opcionesEmail.inlineImages = { graficoProgreso: graficoBlob }; }

  try {
    GmailApp.sendEmail(emailEstudiante, asunto, "", opcionesEmail);
  } catch(e) {
    Logger.log(`Error enviando correo semestral a ${emailEstudiante}: ${e.toString()}`);
  }
}
