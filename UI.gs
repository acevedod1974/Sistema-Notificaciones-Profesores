/**
 * ARCHIVO: UI.gs
 * RESPONSABILIDAD: Interfaz de usuario, Menús, Triggers y Dashboard.
 */

function onOpen() {
  const esModoPrueba = obtenerModoPrueba();
  const icono = esModoPrueba ? "🟢" : "🔴";
  const estadoTexto = esModoPrueba ? "ACTIVO" : "INACTIVO";
  
  SpreadsheetApp.getUi()
      .createMenu('📧 Notificaciones')
      .addItem('Enviar Correos Pendientes', 'enviarCorreosPendientes')
      .addSeparator()
      .addItem('Enviar Reporte Semestral a Todos', 'enviarReporteSemestral')
      .addSeparator()
      .addItem('📊 Ver Estadísticas del Grupo', 'mostrarEstadisticasProfesor') 
      .addSeparator()
      .addItem(`🔄 Alternar Modo Prueba (${icono} ${estadoTexto})`, 'alternarModoPrueba')
      .addItem('🔐 Configurar Token Telegram', 'guardarTokenTelegram')
      .addToUi();
}

function onEdit(e) {
  const rango = e.range;
  const hoja = e.source.getActiveSheet();
  const filaEditada = rango.getRow();
  const colEditada = rango.getColumn();
  
  if (hoja.getName() !== CONFIG.NOMBRE_HOJA || filaEditada < CONFIG.FILA_INICIO_ESTUDIANTES || !CONFIG.COLUMNAS_CALIFICACIONES.includes(colEditada)) { return; }
  
  const valorEditado = rango.getDisplayValue();
  const celdaEstado = hoja.getRange(filaEditada, CONFIG.COL_ESTADO_CORREO);
  
  if (valorEditado === "") { 
    rango.setBackground(null); 
    celdaEstado.setValue('Actualizado'); 
    return; 
  }
  
  celdaEstado.setValue(`Pendiente:${rango.getA1Notation()}`);
  
  const notaObtenida = toNumero(valorEditado); // Refactorizado: Uso de toNumero
  if (isNaN(notaObtenida)) { return; }
  
  const notaMaximaEvaluacion = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, colEditada).getValue());
  
  if (isNaN(notaMaximaEvaluacion) || notaMaximaEvaluacion <= 0) { return; }
  if (notaObtenida > notaMaximaEvaluacion) { rango.setBackground('#FFC0CB'); } else { rango.setBackground(null); }
}

function doPost(e) {
  Logger.log("Llamada POST recibida de AppSheet.");
  try {
    const postData = JSON.parse(e.postData.contents);
    const numeroDeFila = postData.fila;
    if (!numeroDeFila) return ContentService.createTextOutput("Error: No se proporcionó el número de fila.");
    
    const resultado = enviarCorreoAUnaFila(numeroDeFila, null); 
    return ContentService.createTextOutput(resultado);
  } catch (error) {
    Logger.log("Error en doPost: " + error.toString());
    return ContentService.createTextOutput("Error en el script: " + error.toString());
  }
}

function mostrarEstadisticasProfesor() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  if (hoja.getLastRow() < CONFIG.FILA_INICIO_ESTUDIANTES) { SpreadsheetApp.getActive().toast("No hay datos suficientes.", "Aviso"); return; }
  
  const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_ESTUDIANTES, 1, hoja.getLastRow() - CONFIG.FILA_INICIO_ESTUDIANTES + 1, hoja.getLastColumn());
  const datos = rangoDatos.getValues();
  const promediosRaw = calcularPromedios(hoja); // Llamada al Helper
  
  let totalEstudiantes = 0;
  let estudiantesEnRiesgo = [];
  let maxPuntosEvaluadosHastaAhora = 0;
  let evaluacionesActivas = [];

  CONFIG.COLUMNAS_CALIFICACIONES.forEach(col => {
    const promedioCol = promediosRaw[col];
    const nombreEvaluacion = hoja.getRange(CONFIG.FILA_ENCABEZADOS, col).getValue();
    if (promedioCol > 0) {
       const maxNota = toNumero(hoja.getRange(CONFIG.FILA_NOTA_MAXIMA, col).getValue()); // Refactorizado
       maxPuntosEvaluadosHastaAhora += maxNota;
       evaluacionesActivas.push({nombre: nombreEvaluacion, promedio: promedioCol, max: maxNota});
    }
  });
  if (maxPuntosEvaluadosHastaAhora === 0) maxPuntosEvaluadosHastaAhora = 1; 

  let sumaPuntajesAcumulados = 0;
  datos.forEach(fila => {
    const nombreCompleto = `${fila[CONFIG.COL_NOMBRES - 1]} ${fila[CONFIG.COL_APELLIDOS - 1]}`;
    if (fila[CONFIG.COL_NOMBRES - 1]) { 
      totalEstudiantes++;
      const acumuladoTotal = toNumero(fila[CONFIG.COL_SUMA_100 - 1]); // Refactorizado
      const rendimientoReal = (acumuladoTotal / maxPuntosEvaluadosHastaAhora) * 100;
      sumaPuntajesAcumulados += acumuladoTotal;
      if (rendimientoReal < CONFIG.UMBRAL_APROBACION) {
        estudiantesEnRiesgo.push({ nombre: nombreCompleto, rendimiento: rendimientoReal.toFixed(1), puntos: acumuladoTotal.toFixed(1) });
      }
    }
  });

  const promedioClasePuntos = (totalEstudiantes > 0) ? (sumaPuntajesAcumulados / totalEstudiantes) : 0;
  const porcentajeRendimientoGlobal = (promedioClasePuntos / maxPuntosEvaluadosHastaAhora) * 100;
  
  let dataEvaluaciones = [['Evaluación', 'Promedio', { role: 'style' }]];
  evaluacionesActivas.forEach(ev => {
    const porcentaje = (ev.promedio / ev.max);
    const color = porcentaje < 0.5 ? '#dc3545' : '#005A9C';
    dataEvaluaciones.push([ev.nombre, ev.promedio, color]);
  });

  const htmlTemplate = HtmlService.createTemplate(`<html><head><script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script><script type="text/javascript">google.charts.load('current', {'packages':['corechart', 'bar', 'gauge']});google.charts.setOnLoadCallback(drawCharts);function drawCharts() {var dataGauge = google.visualization.arrayToDataTable([['Label', 'Value'],['Rendimiento', ${porcentajeRendimientoGlobal.toFixed(1)}]]);var optionsGauge = {width: 400, height: 200,redFrom: 0, redTo: 50,yellowFrom: 50, yellowTo: 75,greenFrom: 75, greenTo: 100,minorTicks: 5};var chartGauge = new google.visualization.Gauge(document.getElementById('chart_div_gauge'));chartGauge.draw(dataGauge, optionsGauge);var dataBar = google.visualization.arrayToDataTable(${JSON.stringify(dataEvaluaciones)});var optionsBar = {title: 'Promedios por Evaluación (Sobre nota máxima real)',legend: { position: 'none' },vAxis: {title: 'Puntos'},hAxis: {title: 'Evaluación'},animation: {startup: true, duration: 1000, easing: 'out'}};var chartBar = new google.visualization.ColumnChart(document.getElementById('chart_div_bar'));chartBar.draw(dataBar, optionsBar);}</script><style>body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f4f6f9; color: #333; } .header { text-align: center; margin-bottom: 20px; } .kpi-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; } .kpi-card { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; flex: 1; } .kpi-value { font-size: 28px; font-weight: bold; color: #005A9C; } .kpi-sub { font-size: 12px; color: #777; margin-top: 5px; } .chart-row { display: flex; gap: 20px; margin-bottom: 20px; } .chart-box { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex: 1; min-height: 250px; display: flex; justify-content: center; align-items: center; flex-direction: column; } .risk-table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); } .risk-table th { background: #dc3545; color: white; padding: 10px; text-align: left; } .risk-table td { padding: 10px; border-bottom: 1px solid #eee; } .btn-print { background: #005A9C; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; margin-right: 10px; } .btn-close { background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; } @media print { .no-print { display: none !important; } body { background: white; } .chart-box, .kpi-card, .risk-table { box-shadow: none; border: 1px solid #ddd; } }</style></head><body><div class="header"><h2>Reporte de Rendimiento: ${CONFIG.NOMBRE_ASIGNATURA}</h2><p>Profesor: ${CONFIG.NOMBRE_PROFESOR} | Fecha: ${new Date().toLocaleDateString()}</p></div><div class="kpi-container"><div class="kpi-card"><div class="kpi-value">${totalEstudiantes}</div><div>Estudiantes</div></div><div class="kpi-card"><div class="kpi-value">${maxPuntosEvaluadosHastaAhora} pts</div><div>Evaluado hasta hoy</div><div class="kpi-sub">Total acumulable a la fecha</div></div><div class="kpi-card"><div class="kpi-value">${promedioClasePuntos.toFixed(2)}</div><div>Promedio Puntos</div><div class="kpi-sub">Promedio real (Base 100)</div></div><div class="kpi-card"><div class="kpi-value">${porcentajeRendimientoGlobal.toFixed(1)}%</div><div>Efectividad Global</div><div class="kpi-sub">% de puntos obtenidos vs posibles</div></div></div><div class="chart-row"><div class="chart-box"><h4 style="margin:0 0 10px 0;">Salud General del Grupo</h4><div id="chart_div_gauge"></div></div><div class="chart-box"><div id="chart_div_bar" style="width:100%; height:100%;"></div></div></div><h3 style="color: #dc3545;">⚠️ Alerta: Estudiantes con Rendimiento Bajo (< 50%)</h3><table class="risk-table"><thead><tr><th>Estudiante</th><th>Puntos Acumulados</th><th>Rendimiento Real (%)</th></tr></thead><tbody>${estudiantesEnRiesgo.length > 0 ? estudiantesEnRiesgo.map(e => `<tr><td>${e.nombre}</td><td>${e.puntos} / ${maxPuntosEvaluadosHastaAhora}</td><td><strong>${e.rendimiento}%</strong></td></tr>`).join('') : '<tr><td colspan="3" style="text-align:center; padding:20px;">🎉 ¡Excelente! No hay estudiantes en zona crítica.</td></tr>'}</tbody></table><div class="no-print" style="text-align:center; margin-top:30px;"><button class="btn-print" onclick="window.print()">🖨️ Imprimir Reporte</button><button class="btn-close" onclick="google.script.host.close()">Cerrar</button></div></body></html>`);
  const htmlOutput = htmlTemplate.evaluate().setWidth(900).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Estadísticas del Grupo');
}
