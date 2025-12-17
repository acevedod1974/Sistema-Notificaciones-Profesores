// /**
//  * @license
//  * Copyright 2025 Daniel Acevedo
//  *
//  * Sistema Automatizado de Notificaciones para Calificaciones
//  * Asistente de Desarrollo IA: Google Gemini
//  * * Versión 6.2: OPTIMIZACIÓN DE RENDIMIENTO (Cálculo único de promedios).
//  */

// // =================================================================
// //                CONFIGURACIÓN PRINCIPAL
// // =================================================================
// function obtenerModoPrueba() {
//   try {
//     const prop = PropertiesService.getScriptProperties().getProperty('MODO_PRUEBA');
//     return prop === 'false' ? false : true; 
//   } catch (e) { return true; }
// }

// const CONFIG = {
//   NOMBRE_PROFESOR: "Daniel Acevedo",
//   NOMBRE_ASIGNATURA: "Procesos de Fabricacion 1",
//   EMAIL_PROFESOR: "dacevedo@unexpo.edu.ve",
  
//   NOMBRE_HOJA: "Hoja 1",
//   FILA_ENCABEZADOS: 1,
//   FILA_NOTA_MAXIMA: 2,
//   FILA_INICIO_ESTUDIANTES: 3,
  
//   // MAPEO DE COLUMNAS
//   COL_NOMBRES: 3,
//   COL_APELLIDOS: 2,
//   COL_EMAIL: 4, 
  
//   COLUMNAS_CALIFICACIONES: [5, 6, 7, 8, 9, 10],
//   COLUMNAS_RESTANTES: [8, 9, 10], 
  
//   COL_SUMA_100: 11, 
//   COL_ESTADO_CORREO: 13, 
//   COL_TELEGRAM_ID: 14,
//   COL_MENSAJE_PERSONALIZADO: 15, 
  
//   UMBRAL_APROBACION: 50,
//   UMBRAL_BUEN_ESTADO: 65,
//   UMBRAL_EXAMEN_ALTO: 0.80, 
//   UMBRAL_NOTA_BAJA: 0.50, 
//   UMBRAL_PROMEDIO_ALTO: 80, 
//   UMBRAL_RIESGO: 60,
  
//   MODO_PRUEBA: obtenerModoPrueba(), 
//   EMAIL_PRUEBA: "acevedod1974@gmail.com"
// };

// // =================================================================
// //          INTERFAZ DE USUARIO Y DISPARADORES
// // =================================================================
// function onOpen() {
//   const esModoPrueba = obtenerModoPrueba();
//   const icono = esModoPrueba ? "🟢" : "🔴";
//   const estadoTexto = esModoPrueba ? "ACTIVO" : "INACTIVO";
  
//   SpreadsheetApp.getUi()
//       .createMenu('📧 Notificaciones')
//       .addItem('Enviar Correos Pendientes', 'enviarCorreosPendientes')
//       .addSeparator()
//       .addItem('Enviar Reporte Semestral a Todos', 'enviarReporteSemestral')
//       .addSeparator()
//       .addItem('📊 Ver Estadísticas del Grupo', 'mostrarEstadisticasProfesor') 
//       .addSeparator()
//       .addItem(`🔄 Alternar Modo Prueba (${icono} ${estadoTexto})`, 'alternarModoPrueba')
//       .addItem('🔐 Configurar Token Telegram', 'guardarTokenTelegram')
//       .addToUi();
// }

// function onEdit(e) {
//   const rango = e.range;
//   const hoja = e.source.getActiveSheet();
//   const filaEditada = rango.getRow();
//   const colEditada = rango.getColumn();
  
//   if (hoja.getName() !== CONFIG.NOMBRE_HOJA || filaEditada < CONFIG.FILA_INICIO_ESTUDIANTES || !CONFIG.COLUMNAS_CALIFICACIONES.includes(colEditada)) { return; }
  
//   const valorEditado = rango.getDisplayValue();
//   const celdaEstado = hoja.getRange(filaEditada, CONFIG.COL_ESTADO_CORREO);
  
//   if (valorEditado === "") { 
//     rango.setBackground(null); 
//     celdaEstado.setValue('Actualizado'); 
//     return; 
//   }
  
//   celdaEstado.setValue(`Pendiente:${rango.getA1Notation()}`);
  
//   const valorEstandarizado = valorEditado.replace(',', '.');
//   if (isNaN(parseFloat(valorEstandarizado))) { return; }
  
//   const notaObtenida = parseFloat(valorEstandarizado);
//   const notaMaximaEvaluacion = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue());
  
//   if (isNaN(notaMaximaEvaluacion) || notaMaximaEvaluacion <= 0) { return; }
//   if (notaObtenida > notaMaximaEvaluacion) { rango.setBackground('#FFC0CB'); } else { rango.setBackground(null); }
// }

// function doPost(e) {
//   Logger.log("Llamada POST recibida de AppSheet.");
//   try {
//     const postData = JSON.parse(e.postData.contents);
//     const numeroDeFila = postData.fila;
//     if (!numeroDeFila) return ContentService.createTextOutput("Error: No se proporcionó el número de fila.");
    
//     // Al venir de AppSheet es una sola fila, calculamos promedios al momento
//     const resultado = enviarCorreoAUnaFila(numeroDeFila, null); 
//     return ContentService.createTextOutput(resultado);
//   } catch (error) {
//     Logger.log("Error en doPost: " + error.toString());
//     return ContentService.createTextOutput("Error en el script: " + error.toString());
//   }
// }

// // =================================================================
// //          LÓGICA DE ENVÍO INDIVIDUAL (OPTIMIZADA)
// // =================================================================
// function enviarCorreosPendientes() {
//   const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  
//   if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) {
//     SpreadsheetApp.getActive().toast('No hay estudiantes para procesar.', '⚠️ Aviso');
//     return;
//   }

//   try {
//     const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
//     const todosLosDatos = rangoDatos.getValues();
    
//     // OPTIMIZACIÓN: Calculamos promedios UNA SOLA VEZ antes del bucle
//     const promediosPrecalculados = calcularPromedios(hoja);
    
//     let correosEnviados = 0;
    
//     todosLosDatos.forEach((datosFila, index) => {
//       const estadoCorreo = datosFila[CONFIG.COL_ESTADO_CORREO - 1].toString();
//       const filaActual = CONFIG.FILA_INICIO_ESTUDIANTES + index;
      
//       if (estadoCorreo.startsWith('Pendiente:')) {
//         // Pasamos los promedios ya calculados
//         enviarCorreoAUnaFila(filaActual, estadoCorreo, promediosPrecalculados, hoja, datosFila);
//         correosEnviados++;
//       }
//     });
    
//     SpreadsheetApp.flush(); 
//     SpreadsheetApp.getActive().toast(`Se enviaron ${correosEnviados} correos pendientes.`, '✅ Proceso Finalizado', 5);
    
//   } catch (error) {
//     Logger.log("Error crítico: " + error.toString());
//     SpreadsheetApp.getUi().alert('Error: ' + error.toString());
//   }
// }

// // Se añade soporte para recibir promedios, hoja y datos para evitar lecturas redundantes
// function enviarCorreoAUnaFila(numeroDeFila, estadoCorreo, promediosInyectados = null, hojaInyectada = null, datosFilaInyectados = null) {
//   const hoja = hojaInyectada || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
//   const filaActual = parseInt(numeroDeFila);
//   const datosFila = datosFilaInyectados || hoja.getRange(filaActual, 1, 1, hoja.getLastColumn()).getValues()[0];
  
//   // Si no nos dieron promedios (ej. llamada desde AppSheet), los calculamos aquí
//   const promediosPorColumna = promediosInyectados || calcularPromedios(hoja);
  
//   const nombreEstudiante = `${datosFila[CONFIG.COL_NOMBRES - 1]} ${datosFila[CONFIG.COL_APELLIDOS - 1]}`;
//   const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1]; 
//   const telegramChatId = datosFila[CONFIG.COL_TELEGRAM_ID - 1]; 
//   const mensajePersonalizado = datosFila[CONFIG.COL_MENSAJE_PERSONALIZADO - 1];

//   let htmlMensajePersonalizado = "";
//   if (mensajePersonalizado && mensajePersonalizado.toString().trim() !== "") {
//     htmlMensajePersonalizado = `
//       <div style="background-color: #fff3cd; color: #856404; padding: 15px; margin: 15px 0; border: 1px solid #ffeeba; border-radius: 4px;">
//         <strong>Nota del Profesor:</strong><br>${mensajePersonalizado}
//       </div>`;
//   }

//   if (!emailEstudiante) {
//     hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue('Error: Falta email');
//     return "Error: Falta email";
//   }

//   let notaObtenida, notaMaximaEvaluacion, nombreEvaluacion;
  
//   // Lógica para determinar qué evaluación gatilló el correo
//   if (estadoCorreo && estadoCorreo.startsWith('Pendiente:')) {
//     const a1Notation = estadoCorreo.split(':')[1];
//     // Optimizacion: Intentar obtener columna del A1 Notation sin llamar a getRange si es posible, 
//     // pero para seguridad mantenemos getRange solo para esta celda especifica.
//     const celdaEditada = hoja.getRange(a1Notation);
//     notaObtenida = celdaEditada.getValue();
//     const colEditada = celdaEditada.getColumn();
//     notaMaximaEvaluacion = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue());
//     nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, colEditada).getValue();
//   } else {
//      // Si no hay celda específica (ej. AppSheet o reenvío genérico), buscamos la última válida o relevante
//      CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//        const valorCelda = datosFila[col-1];
//        if (valorCelda !== "" && !isNaN(parseFloat(valorCelda.toString().replace(',','.')))) {
//          notaObtenida = parseFloat(valorCelda.toString().replace(',','.'));
//          notaMaximaEvaluacion = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
//          nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
//        }
//      });
//   }
  
//   const suma100 = parseFloat(datosFila[CONFIG.COL_SUMA_100 - 1]);
//   let puntajeAcumulado = 0;
//   let puntajeMaximoPosible = 0;
  
//   CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//     const valorCelda = datosFila[col - 1];
//     if (valorCelda !== "" && !isNaN(parseFloat(valorCelda.toString().replace(',','.')))) {
//       puntajeAcumulado += parseFloat(valorCelda.toString().replace(',','.'));
//       puntajeMaximoPosible += parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
//     }
//   });

//   if (puntajeMaximoPosible === 0) {
//     hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue('Error: Sin notas');
//     return "Error: Sin notas";
//   }
  
//   const rendimientoNormalizado = puntajeAcumulado / puntajeMaximoPosible * 100;
//   const rendimientoExamen = notaObtenida / notaMaximaEvaluacion;
  
//   let asunto = "", textoIntroductorio = "", botonHTML = "", mensajeTelegram = "";

//   if (rendimientoExamen < CONFIG.UMBRAL_NOTA_BAJA || (rendimientoNormalizado < CONFIG.UMBRAL_RIESGO && rendimientoExamen < CONFIG.UMBRAL_EXAMEN_ALTO) ) {
//     asunto = `Importante: Revisión sobre tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
//     textoIntroductorio = `Hola ${nombreEstudiante},<br><br>Te escribo para revisar tu calificación de <b>${notaObtenida.toFixed(2)}</b> en "<b>${nombreEvaluacion}</b>". He notado que tu rendimiento requiere atención en esta evaluación.<br><br>Es un momento importante para corregir el rumbo. Estoy aquí para ayudarte. Por favor, agenda una asesoría.`;
//     const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=Solicitud%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
//     botonHTML = crearBotonCtaHTML("Agendar Asesoría por Correo", enlaceMailto);
//     mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. Se ha registrado tu nota en *${nombreEvaluacion}*: ${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)} (Baja). Tu acumulado es ${suma100.toFixed(2)}/100. Contáctame si necesitas ayuda.`;
//   } else if (rendimientoExamen >= CONFIG.UMBRAL_EXAMEN_ALTO) {
//     asunto = `¡Excelente trabajo en ${nombreEvaluacion}, ${nombreEstudiante}!`;
//     textoIntroductorio = `¡Felicidades, ${nombreEstudiante}!<br><br>Quería destacar especialmente tu calificación de <b>${notaObtenida.toFixed(2)}</b> en "<b>${nombreEvaluacion}</b>". ¡Un resultado sobresaliente!`;
//     mensajeTelegram = `🌟 ¡Hola ${nombreEstudiante}! Excelente nota en *${nombreEvaluacion}*: *${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}*. Tu acumulado es ${suma100.toFixed(2)}/100. ¡Sigue así!`;
//   } else {
//     asunto = `Actualización de tu progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
//     textoIntroductorio = `Estimado ${nombreEstudiante},<br><br>Este correo es para confirmar que tus calificaciones han sido actualizadas. Tu rendimiento en "<b>${nombreEvaluacion}</b>" fue de <b>${notaObtenida.toFixed(2)}</b>.`;
//     mensajeTelegram = `📝 Hola ${nombreEstudiante}. Se ha actualizado tu nota en *${nombreEvaluacion}*: ${notaObtenida.toFixed(2)}/${notaMaximaEvaluacion.toFixed(2)}. Tu acumulado es ${suma100.toFixed(2)}/100.`;
//   }
  
//   if (telegramChatId && telegramChatId.toString().trim() !== '') {
//     let msgTgFinal = mensajeTelegram;
//     if (mensajePersonalizado && mensajePersonalizado.toString().trim() !== "") {
//       msgTgFinal += `\n\n📝 *Nota del Profesor:* ${mensajePersonalizado}`;
//     }
//     enviarNotificacionTelegram(telegramChatId, msgTgFinal);
//   }
  
//   const encabezadoHTML = crearEncabezadoEmailHTML();
//   const tablaHTML = crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna);
//   const graficoBlob = crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna);
  
//   let cuerpoFinalHTML = `<div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; color: #333;">${encabezadoHTML}<div style="padding: 20px;"><p>${textoIntroductorio}</p>${htmlMensajePersonalizado}${botonHTML}${tablaHTML}<p>Para una mejor visualización de tu avance, puedes ver el siguiente gráfico:</p><p style="text-align:center;"><img src="cid:graficoProgreso"></p><p>Saludos,<br><b>${CONFIG.NOMBRE_PROFESOR}</b><br>Profesor de ${CONFIG.NOMBRE_ASIGNATURA}</p></div></div>`;
  
//   const opcionesEmail = { htmlBody: cuerpoFinalHTML };
//   if (graficoBlob) { opcionesEmail.inlineImages = { graficoProgreso: graficoBlob }; }

//   try {
//     GmailApp.sendEmail(emailEstudiante, asunto, "", opcionesEmail);
//     hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue(`Enviado ${new Date().toLocaleString()}`);
//     return "Correo enviado exitosamente a la fila " + numeroDeFila;
//   } catch (error) {
//     Logger.log("ERROR al enviar correo: " + error.toString());
//     hoja.getRange(filaActual, CONFIG.COL_ESTADO_CORREO).setValue(`Error al enviar: ${error.message}`);
//     return "Error al enviar correo: " + error.message;
//   }
// }

// // =================================================================
// //          FUNCIONES PARA REPORTE SEMESTRAL
// // =================================================================
// function enviarReporteSemestral() {
//   const ui = SpreadsheetApp.getUi();
//   const confirmacion = ui.alert('Confirmación de Envío Masivo', 'Estás a punto de enviar un reporte de progreso a TODOS los estudiantes. ¿Deseas continuar?', ui.ButtonSet.YES_NO);
//   if (confirmacion !== ui.Button.YES) { return; }

//   try {
//     const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
//     if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) { 
//       SpreadsheetApp.getActive().toast('No hay estudiantes para procesar.', 'Aviso'); return; 
//     }
//     const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
//     const todosLosDatos = rangoDatos.getValues();
//     let correosEnviados = 0;

//     // OPTIMIZACIÓN: Cálculo único
//     const promediosPorColumna = calcularPromedios(hoja);
    
//     let puntajeMaximoRestante = 0;
//     CONFIG.COLUMNAS_RESTANTES.forEach(col => {
//       const val = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
//       if (!isNaN(val)) puntajeMaximoRestante += val;
//     });

//     todosLosDatos.forEach((datosFila) => {
//       const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1]; 
//       if (emailEstudiante) {
//         analizarYEnviarCorreoSemestral(hoja, datosFila, puntajeMaximoRestante, promediosPorColumna);
//         correosEnviados++;
//       }
//     });
    
//     SpreadsheetApp.flush(); 
//     SpreadsheetApp.getActive().toast(`Se han enviado ${correosEnviados} reportes semestrales.`, '✅ Proceso Finalizado', 5);
    
//   } catch (error) {
//     Logger.log("Error en reporte semestral: " + error.toString());
//     ui.alert('Error: ' + error.toString());
//   }
// }

// function analizarYEnviarCorreoSemestral(hoja, datosFila, puntajeMaximoRestante, promediosPorColumna) {
//   const nombreEstudiante = `${datosFila[CONFIG.COL_NOMBRES - 1]} ${datosFila[CONFIG.COL_APELLIDOS - 1]}`;
//   const emailEstudiante = CONFIG.MODO_PRUEBA ? CONFIG.EMAIL_PRUEBA : datosFila[CONFIG.COL_EMAIL - 1];
//   const telegramChatId = datosFila[CONFIG.COL_TELEGRAM_ID - 1];
//   const mensajePersonalizado = datosFila[CONFIG.COL_MENSAJE_PERSONALIZADO - 1];

//   let htmlMensajePersonalizado = "";
//   if (mensajePersonalizado && mensajePersonalizado.toString().trim() !== "") {
//     htmlMensajePersonalizado = `<div style="background-color: #fff3cd; color: #856404; padding: 15px; margin: 15px 0; border: 1px solid #ffeeba; border-radius: 4px;"><strong>Nota del Profesor:</strong><br>${mensajePersonalizado}</div>`;
//   }
  
//   const puntajeAcumuladoActual = parseFloat(datosFila[CONFIG.COL_SUMA_100 - 1] || 0);
//   const proyeccionFinalMaxima = puntajeAcumuladoActual + puntajeMaximoRestante;

//   let puntajeTotalDelCurso = 0;
//   CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//     const val = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue());
//     if(!isNaN(val)) puntajeTotalDelCurso += val;
//   });
  
//   const puntajeMaximoEvaluado = puntajeTotalDelCurso - puntajeMaximoRestante;
//   const rendimientoActual = (puntajeMaximoEvaluado > 0) ? (puntajeAcumuladoActual / puntajeMaximoEvaluado) * 100 : 0;

//   let asunto = "", cuerpoMensaje = "", botonHTML = "", mensajeTelegram = "";

//   if (proyeccionFinalMaxima < CONFIG.UMBRAL_APROBACION) {
//     asunto = `URGENTE: Reunión sobre tu futuro en ${CONFIG.NOMBRE_ASIGNATURA}`;
//     cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo con urgencia sobre tu situación. El análisis indica que, <b>incluso obteniendo la máxima calificación en lo que resta, la proyección no alcanza el mínimo aprobatorio.</b></p><p>Es muy importante que nos reunamos esta semana para discutir tu caso y explorar opciones.</p>`;
//     const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=URGENTE:%20Reunión%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
//     botonHTML = crearBotonCtaHTML("Contactar al Profesor para Agendar Reunión", enlaceMailto);
//     mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. URGENTE: Tu proyección actual no alcanza para aprobar la materia. Por favor contáctame urgente.`;

//   } else if (rendimientoActual < CONFIG.UMBRAL_BUEN_ESTADO) {
//     asunto = `Información Importante sobre tu Situación en ${CONFIG.NOMBRE_ASIGNATURA}`;
//     cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo para conversar sobre tu situación académica. He realizado un análisis y la buena noticia es que <b>todavía es posible que alcances la calificación final aprobatoria de ${CONFIG.UMBRAL_APROBACION} puntos.</b></p><p>Para lograrlo, se requiere un rendimiento excepcional en las evaluaciones restantes. Si quieres que conversemos para trazar un plan, no dudes en contactarme.</p>`;
//     const enlaceMailto = `mailto:${CONFIG.EMAIL_PROFESOR}?subject=Solicitud%20de%20Asesoría%20-%20${encodeURIComponent(CONFIG.NOMBRE_ASIGNATURA)}`;
//     botonHTML = crearBotonCtaHTML("Solicitar Asesoría", enlaceMailto);
//     mensajeTelegram = `⚠️ Hola ${nombreEstudiante}. Reporte Semestral: Tu situación es delicada pero salvable. Aún puedes aprobar con esfuerzo máximo. Contáctame si necesitas asesoría.`;
//   } else {
//     asunto = `Reconocimiento de tu Progreso en ${CONFIG.NOMBRE_ASIGNATURA}`;
//     cuerpoMensaje = `<p>Estimado ${nombreEstudiante},</p><p>Te escribo para darte un reconocimiento por tu desempeño y esfuerzo. Tus resultados hasta ahora son sólidos y te posicionan favorablemente para el cierre del curso. ¡Sigue así!</p>`;
//     mensajeTelegram = `🌟 Hola ${nombreEstudiante}. Reporte Semestral: ¡Felicidades! Vas por excelente camino. Mantén el esfuerzo.`;
//   }
  
//   if (telegramChatId && telegramChatId.toString().trim() !== '') {
//     let msgTgFinal = mensajeTelegram;
//     if (mensajePersonalizado && mensajePersonalizado.toString().trim() !== "") {
//       msgTgFinal += `\n\n📝 *Nota del Profesor:* ${mensajePersonalizado}`;
//     }
//     enviarNotificacionTelegram(telegramChatId, msgTgFinal);
//   }

//   const tablaHTML = crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna);
//   const graficoBlob = crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna);
  
//   let cuerpoFinalHTML = `<div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; color: #333;">${crearEncabezadoEmailHTML()}<div style="padding:20px;"><p>${cuerpoMensaje}</p>${htmlMensajePersonalizado}${tablaHTML}${botonHTML}<p>Para una mejor visualización de tu avance, puedes ver el siguiente gráfico:</p><p style="text-align:center;"><img src="cid:graficoProgreso"></p><p>Saludos,<br><b>${CONFIG.NOMBRE_PROFESOR}</b></p></div></div>`;
//   const opcionesEmail = { htmlBody: cuerpoFinalHTML };
//   if (graficoBlob) { opcionesEmail.inlineImages = { graficoProgreso: graficoBlob }; }

//   try {
//     GmailApp.sendEmail(emailEstudiante, asunto, "", opcionesEmail);
//   } catch(e) {
//     Logger.log(`Error enviando correo semestral a ${emailEstudiante}: ${e.toString()}`);
//   }
// }

// // =================================================================
// //          DASHBOARD DE ESTADÍSTICAS
// // =================================================================
// function mostrarEstadisticasProfesor() {
//   const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
//   if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) { SpreadsheetApp.getActive().toast("No hay datos suficientes.", "Aviso"); return; }
//   const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
//   const datos = rangoDatos.getValues();
//   const promediosRaw = calcularPromedios(hoja); 
//   let totalEstudiantes = 0;
//   let estudiantesEnRiesgo = [];
//   let maxPuntosEvaluadosHastaAhora = 0;
//   let evaluacionesActivas = [];

//   CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//     const promedioCol = promediosRaw[col];
//     const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
//     if (promedioCol > 0) {
//        const maxNota = parseFloat(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue() || 0);
//        maxPuntosEvaluadosHastaAhora += maxNota;
//        evaluacionesActivas.push({nombre: nombreEvaluacion, promedio: promedioCol, max: maxNota});
//     }
//   });
//   if (maxPuntosEvaluadosHastaAhora === 0) maxPuntosEvaluadosHastaAhora = 1; 

//   let sumaPuntajesAcumulados = 0;
//   datos.forEach(fila => {
//     const nombreCompleto = `${fila[CONFIG.COL_NOMBRES - 1]} ${fila[CONFIG.COL_APELLIDOS - 1]}`;
//     if (fila[CONFIG.COL_NOMBRES - 1]) { 
//       totalEstudiantes++;
//       const acumuladoTotal = parseFloat(fila[CONFIG.COL_SUMA_100 - 1] || 0);
//       const rendimientoReal = (acumuladoTotal / maxPuntosEvaluadosHastaAhora) * 100;
//       sumaPuntajesAcumulados += acumuladoTotal;
//       if (rendimientoReal < CONFIG.UMBRAL_APROBACION) {
//         estudiantesEnRiesgo.push({ nombre: nombreCompleto, rendimiento: rendimientoReal.toFixed(1), puntos: acumuladoTotal.toFixed(1) });
//       }
//     }
//   });

//   const promedioClasePuntos = (totalEstudiantes > 0) ? (sumaPuntajesAcumulados / totalEstudiantes) : 0;
//   const porcentajeRendimientoGlobal = (promedioClasePuntos / maxPuntosEvaluadosHastaAhora) * 100;
  
//   let dataEvaluaciones = [['Evaluación', 'Promedio', { role: 'style' }]];
//   evaluacionesActivas.forEach(ev => {
//     const porcentaje = (ev.promedio / ev.max);
//     const color = porcentaje < 0.5 ? '#dc3545' : '#005A9C';
//     dataEvaluaciones.push([ev.nombre, ev.promedio, color]);
//   });

//   const htmlTemplate = HtmlService.createTemplate(`<html><head><script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script><script type="text/javascript">google.charts.load('current', {'packages':['corechart', 'bar', 'gauge']});google.charts.setOnLoadCallback(drawCharts);function drawCharts() {var dataGauge = google.visualization.arrayToDataTable([['Label', 'Value'],['Rendimiento', ${porcentajeRendimientoGlobal.toFixed(1)}]]);var optionsGauge = {width: 400, height: 200,redFrom: 0, redTo: 50,yellowFrom: 50, yellowTo: 75,greenFrom: 75, greenTo: 100,minorTicks: 5};var chartGauge = new google.visualization.Gauge(document.getElementById('chart_div_gauge'));chartGauge.draw(dataGauge, optionsGauge);var dataBar = google.visualization.arrayToDataTable(${JSON.stringify(dataEvaluaciones)});var optionsBar = {title: 'Promedios por Evaluación (Sobre nota máxima real)',legend: { position: 'none' },vAxis: {title: 'Puntos'},hAxis: {title: 'Evaluación'},animation: {startup: true, duration: 1000, easing: 'out'}};var chartBar = new google.visualization.ColumnChart(document.getElementById('chart_div_bar'));chartBar.draw(dataBar, optionsBar);}</script><style>body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f4f6f9; color: #333; } .header { text-align: center; margin-bottom: 20px; } .kpi-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; } .kpi-card { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; flex: 1; } .kpi-value { font-size: 28px; font-weight: bold; color: #005A9C; } .kpi-sub { font-size: 12px; color: #777; margin-top: 5px; } .chart-row { display: flex; gap: 20px; margin-bottom: 20px; } .chart-box { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex: 1; min-height: 250px; display: flex; justify-content: center; align-items: center; flex-direction: column; } .risk-table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); } .risk-table th { background: #dc3545; color: white; padding: 10px; text-align: left; } .risk-table td { padding: 10px; border-bottom: 1px solid #eee; } .btn-print { background: #005A9C; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; margin-right: 10px; } .btn-close { background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; } @media print { .no-print { display: none !important; } body { background: white; } .chart-box, .kpi-card, .risk-table { box-shadow: none; border: 1px solid #ddd; } }</style></head><body><div class="header"><h2>Reporte de Rendimiento: ${CONFIG.NOMBRE_ASIGNATURA}</h2><p>Profesor: ${CONFIG.NOMBRE_PROFESOR} | Fecha: ${new Date().toLocaleDateString()}</p></div><div class="kpi-container"><div class="kpi-card"><div class="kpi-value">${totalEstudiantes}</div><div>Estudiantes</div></div><div class="kpi-card"><div class="kpi-value">${maxPuntosEvaluadosHastaAhora} pts</div><div>Evaluado hasta hoy</div><div class="kpi-sub">Total acumulable a la fecha</div></div><div class="kpi-card"><div class="kpi-value">${promedioClasePuntos.toFixed(2)}</div><div>Promedio Puntos</div><div class="kpi-sub">Promedio real (Base 100)</div></div><div class="kpi-card"><div class="kpi-value">${porcentajeRendimientoGlobal.toFixed(1)}%</div><div>Efectividad Global</div><div class="kpi-sub">% de puntos obtenidos vs posibles</div></div></div><div class="chart-row"><div class="chart-box"><h4 style="margin:0 0 10px 0;">Salud General del Grupo</h4><div id="chart_div_gauge"></div></div><div class="chart-box"><div id="chart_div_bar" style="width:100%; height:100%;"></div></div></div><h3 style="color: #dc3545;">⚠️ Alerta: Estudiantes con Rendimiento Bajo (< 50%)</h3><table class="risk-table"><thead><tr><th>Estudiante</th><th>Puntos Acumulados</th><th>Rendimiento Real (%)</th></tr></thead><tbody>${estudiantesEnRiesgo.length > 0 ? estudiantesEnRiesgo.map(e => `<tr><td>${e.nombre}</td><td>${e.puntos} / ${maxPuntosEvaluadosHastaAhora}</td><td><strong>${e.rendimiento}%</strong></td></tr>`).join('') : '<tr><td colspan="3" style="text-align:center; padding:20px;">🎉 ¡Excelente! No hay estudiantes en zona crítica.</td></tr>'}</tbody></table><div class="no-print" style="text-align:center; margin-top:30px;"><button class="btn-print" onclick="window.print()">🖨️ Imprimir Reporte</button><button class="btn-close" onclick="google.script.host.close()">Cerrar</button></div></body></html>`);
//   const htmlOutput = htmlTemplate.evaluate().setWidth(900).setHeight(1000);
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Estadísticas del Grupo');
// }

// // =================================================================
// //          FUNCIONES AYUDANTES
// // =================================================================

// function alternarModoPrueba() {
//   const estadoActual = obtenerModoPrueba();
//   const nuevoEstado = !estadoActual;
//   PropertiesService.getScriptProperties().setProperty('MODO_PRUEBA', nuevoEstado.toString());
//   const estadoTexto = nuevoEstado ? "🟢 ACTIVADO (Correos a ti)" : "🔴 DESACTIVADO (Correos a alumnos)";
//   SpreadsheetApp.getActive().toast(`Modo Prueba: ${estadoTexto}`, 'Configuración Actualizada');
//   onOpen();
// }

// function guardarTokenTelegram() {
//   const ui = SpreadsheetApp.getUi();
//   const respuesta = ui.prompt('Configuración de Telegram', 'Pega aquí el Token de tu Bot de Telegram:', ui.ButtonSet.OK_CANCEL);
//   if (respuesta.getSelectedButton() == ui.Button.OK) {
//     const token = respuesta.getResponseText().trim();
//     if (token) {
//       PropertiesService.getScriptProperties().setProperty('TELEGRAM_BOT_TOKEN', token);
//       ui.alert('Éxito', 'El token se ha guardado de forma segura.', ui.ButtonSet.OK);
//     } else { ui.alert('Error', 'El token no puede estar vacío.', ui.ButtonSet.OK); }
//   }
// }

// function enviarNotificacionTelegram(chatId, mensaje) {
//   const scriptProperties = PropertiesService.getScriptProperties();
//   const token = scriptProperties.getProperty('TELEGRAM_BOT_TOKEN');
//   if (!token) { Logger.log("ADVERTENCIA: Token Telegram no configurado."); return; }
//   const url = `https://api.telegram.org/bot${token}/sendMessage`;
//   const payload = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify({'chat_id': String(chatId),'text': mensaje,'parse_mode': 'Markdown'}) };
//   try { UrlFetchApp.fetch(url, payload); Logger.log(`Notificación Telegram enviada a: ${chatId}`); } catch (e) { Logger.log(`Error Telegram a ${chatId}: ${e.toString()}`); }
// }

// function crearEncabezadoEmailHTML() { return `<div style="background-color: #f8f9fa; padding: 20px; border-bottom: 5px solid #005A9C; text-align: center;"><h2 style="color: #005A9C; margin-top: 10px; margin-bottom: 0;">Reporte de Progreso Académico</h2><p style="color: #6c757d; margin-top: 5px; margin-bottom: 0;">${CONFIG.NOMBRE_ASIGNATURA}</p></div>`; }
// function crearBotonCtaHTML(texto, enlace) { return `<table border="0" cellpadding="0" cellspacing="0" style="margin: 20px 0;"><tr><td align="center" bgcolor="#005A9C" style="border-radius: 5px;"><a href="${enlace}" target="_blank" style="font-size: 16px; font-family: Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 5px; padding: 12px 25px; border: 1px solid #005A9C; display: inline-block;">${texto}</a></td></tr></table>`; }

// function calcularPromedios(hoja) {
//   const promedios = {};
//   const ultimaFila = hoja.getLastRow();
//   if (ultimaFila < CONFIG.FILA_INICIO_ESTUDIANTES) { CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => promedios[col] = 0); return promedios; }
//   const rangoEstudiantes = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, ultimaFila - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
//   const datosEstudiantes = rangoEstudiantes.getValues();
//   CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//     let suma = 0; let contador = 0;
//     datosEstudiantes.forEach(fila => { const nota = fila[col - 1]; if (nota !== "" && !isNaN(nota)) { suma += parseFloat(nota); contador++; } });
//     promedios[col] = (contador > 0) ? (suma / contador) : 0;
//   });
//   return promedios;
// }

// function crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna) {
//   let tablaHTML = `<p>Para darte un contexto completo, aquí tienes un resumen detallado de tu progreso:</p><table style="width:100%; border-collapse: collapse; border: 1px solid #dee2e6; font-family: Arial, sans-serif; font-size: 14px;"><thead style="background-color: #005A9C; color: #ffffff;"><tr><th style="padding: 12px; border: 1px solid #005A9C; text-align: left;">Evaluación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Tu Calificación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Nota Máxima</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Promedio del Grupo</th></tr></thead><tbody>`;
//   let rowIndex = 0; let puntajeAcumulado = 0; let puntajeMaximoPosible = 0;
//   CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//     const notaEstudiante = datosFila[col - 1];
//     if (notaEstudiante !== "" && !isNaN(notaEstudiante)) {
//       const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
//       const notaMaxima = hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue();
//       const promedioGrupo = promediosPorColumna[col];
//       puntajeAcumulado += parseFloat(notaEstudiante); puntajeMaximoPosible += parseFloat(notaMaxima);
//       const rowColor = (rowIndex % 2 === 0) ? '#ffffff' : '#f8f9fa';
//       tablaHTML += `<tr style="background-color: ${rowColor};"><td style="padding: 12px; border: 1px solid #dee2e6;">${nombreEvaluacion}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;"><b>${parseFloat(notaEstudiante).toFixed(2)}</b></td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${parseFloat(notaMaxima).toFixed(2)}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${promedioGrupo.toFixed(2)}</td></tr>`;
//       rowIndex++;
//     }
//   });
//   tablaHTML += `</tbody><tfoot style="background-color: #005A9C; color: #ffffff; font-weight: bold;"><tr><td style="padding: 12px; border: 1px solid #005A9C;">Total Acumulado Parcial</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeAcumulado.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeMaximoPosible.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">-</td></tr></tfoot></table>`;
//   return tablaHTML;
// }

// function crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna) {
//     const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Evaluación").addColumn(Charts.ColumnType.NUMBER, "Tu Calificación").addColumn(Charts.ColumnType.NUMBER, "Promedio del Grupo");
//     let hayDatos = false;
//     CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
//         const notaEstudiante = datosFila[col - 1];
//         if (notaEstudiante !== "" && !isNaN(notaEstudiante)) {
//             const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
//             const promedioGrupo = promediosPorColumna[col];
//             dataTable.addRow([nombreEvaluacion, parseFloat(notaEstudiante), promedioGrupo]);
//             hayDatos = true;
//         }
//     });
//     if (!hayDatos) return null;
//     const chart = Charts.newColumnChart().setTitle("Tu Progreso vs. el Promedio del Grupo").setDataTable(dataTable).setOption('legend', { position: 'top' }).setOption('vAxis', { title: 'Calificación', minValue: 0 }).setDimensions(600, 400).build();
//     return chart.getAs('image/png');
// }
