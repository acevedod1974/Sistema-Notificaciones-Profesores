/**
 * ARCHIVO: Helpers.gs
 * RESPONSABILIDAD: Funciones utilitarias, formateo y cálculos matemáticos.
 */

// --- NUEVO: Helper Universal de Números (Evita errores de parseo) ---
function toNumero(val) {
  if (val === "" || val == null) return NaN;
  // Convierte a string, reemplaza coma por punto y parsea
  return parseFloat(String(val).replace(',', '.'));
}

// --- NUEVO: Helper para HTML repetitivo ---
function construirNotaPersonalizadaHTML(texto) {
  if (!texto || !String(texto).trim()) return "";
  return `
    <div style="background-color: #fff3cd; color: #856404; padding: 15px; margin: 15px 0; border: 1px solid #ffeeba; border-radius: 4px;">
      <strong>Nota del Profesor:</strong><br>
      ${texto}
    </div>`;
}

function crearEncabezadoEmailHTML() { 
  return `<div style="background-color: #f8f9fa; padding: 20px; border-bottom: 5px solid #005A9C; text-align: center;"><h2 style="color: #005A9C; margin-top: 10px; margin-bottom: 0;">Reporte de Progreso Académico</h2><p style="color: #6c757d; margin-top: 5px; margin-bottom: 0;">${CONFIG.NOMBRE_ASIGNATURA}</p></div>`; 
}

function crearBotonCtaHTML(texto, enlace) { 
  return `<table border="0" cellpadding="0" cellspacing="0" style="margin: 20px 0;"><tr><td align="center" bgcolor="#005A9C" style="border-radius: 5px;"><a href="${enlace}" target="_blank" style="font-size: 16px; font-family: Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 5px; padding: 12px 25px; border: 1px solid #005A9C; display: inline-block;">${texto}</a></td></tr></table>`; 
}

// --- Lógica de Telegram ---
function enviarNotificacionTelegram(chatId, mensaje) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const token = scriptProperties.getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) { Logger.log("ADVERTENCIA: Token Telegram no configurado."); return; }
  
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = { 
    'method': 'post', 
    'contentType': 'application/json', 
    'payload': JSON.stringify({'chat_id': String(chatId),'text': mensaje,'parse_mode': 'Markdown'}) 
  };
  
  try { 
    UrlFetchApp.fetch(url, payload); 
    Logger.log(`Notificación Telegram enviada a: ${chatId}`); 
  } catch (e) { 
    Logger.log(`Error Telegram a ${chatId}: ${e.toString()}`); 
  }
}

// --- Cálculos y Tablas ---
function calcularPromedios(hoja) {
  const promedios = {};
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < CONFIG.FILA_INICIO_ESTUDIANTES) { CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => promedios[col] = 0); return promedios; }
  
  const rangoEstudiantes = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, ultimaFila - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
  const datosEstudiantes = rangoEstudiantes.getValues();
  
  CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
    let suma = 0; let contador = 0;
    datosEstudiantes.forEach(fila => { 
      const nota = toNumero(fila[col - 1]); // Refactorizado: Uso de toNumero
      if (!isNaN(nota)) { suma += nota; contador++; } 
    });
    promedios[col] = (contador > 0) ? (suma / contador) : 0;
  });
  return promedios;
}

function crearTablaDeCalificacionesHTML(hoja, datosFila, promediosPorColumna) {
  let tablaHTML = `<p>Para darte un contexto completo, aquí tienes un resumen detallado de tu progreso:</p><table style="width:100%; border-collapse: collapse; border: 1px solid #dee2e6; font-family: Arial, sans-serif; font-size: 14px;"><thead style="background-color: #005A9C; color: #ffffff;"><tr><th style="padding: 12px; border: 1px solid #005A9C; text-align: left;">Evaluación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Tu Calificación</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Nota Máxima</th><th style="padding: 12px; border: 1px solid #005A9C; text-align: center;">Promedio del Grupo</th></tr></thead><tbody>`;
  let rowIndex = 0; let puntajeAcumulado = 0; let puntajeMaximoPosible = 0;
  
  CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
    const notaEstudiante = toNumero(datosFila[col - 1]); // Refactorizado
    if (!isNaN(notaEstudiante)) {
      const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
      const notaMaxima = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue()); // Refactorizado
      const promedioGrupo = promediosPorColumna[col];
      
      puntajeAcumulado += notaEstudiante; 
      puntajeMaximoPosible += notaMaxima;
      
      const rowColor = (rowIndex % 2 === 0) ? '#ffffff' : '#f8f9fa';
      tablaHTML += `<tr style="background-color: ${rowColor};"><td style="padding: 12px; border: 1px solid #dee2e6;">${nombreEvaluacion}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;"><b>${notaEstudiante.toFixed(2)}</b></td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${notaMaxima.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #dee2e6; text-align: center;">${promedioGrupo.toFixed(2)}</td></tr>`;
      rowIndex++;
    }
  });
  tablaHTML += `</tbody><tfoot style="background-color: #005A9C; color: #ffffff; font-weight: bold;"><tr><td style="padding: 12px; border: 1px solid #005A9C;">Total Acumulado Parcial</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeAcumulado.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">${puntajeMaximoPosible.toFixed(2)}</td><td style="padding: 12px; border: 1px solid #005A9C; text-align: center;">-</td></tr></tfoot></table>`;
  return tablaHTML;
}

function crearGraficoDeProgreso(hoja, datosFila, promediosPorColumna) {
    const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Evaluación").addColumn(Charts.ColumnType.NUMBER, "Tu Calificación").addColumn(Charts.ColumnType.NUMBER, "Promedio del Grupo");
    let hayDatos = false;
    CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
        const notaEstudiante = toNumero(datosFila[col - 1]); // Refactorizado
        if (!isNaN(notaEstudiante)) {
            const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
            const promedioGrupo = promediosPorColumna[col];
            dataTable.addRow([nombreEvaluacion, notaEstudiante, promedioGrupo]);
            hayDatos = true;
        }
    });
    if (!hayDatos) return null;
    const chart = Charts.newColumnChart().setTitle("Tu Progreso vs. el Promedio del Grupo").setDataTable(dataTable).setOption('legend', { position: 'top' }).setOption('vAxis', { title: 'Calificación', minValue: 0 }).setDimensions(600, 400).build();
    return chart.getAs('image/png');
}
