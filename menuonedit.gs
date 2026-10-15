
var SIM_SHEET_NAME = 'SIMULADOR_ACCIONARIO';
var HIST_SHEET_NAME = 'HISTORIAL_SIMULACIONES';

var SIM_RANGE = {
  startRow: 12,
  endRow: 25,
  startCol: 10,
  numCols: 6,
};

// Paleta Cualli
var CUALLI = {
  dark: '#1F2D3D',
  yellow: '#FDB913',
  gray: '#515151',
  greenDark: '#1E8449',
  redDark: '#C0392B',
  blueDark: '#1A5276',
};

// ═══════════════════════════════════════════════════════════════════
//  onOpen — crea el menú al abrir el spreadsheet
// ═══════════════════════════════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Herramientas')
    .addItem('🔢 Calculadora Asamblea', 'mostrarCalculadora')
    .addToUi();
}

function mostrarCalculadora() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('calculadora.html')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}


// ═══════════════════════════════════════════════════════════════════
//  INICIALIZACIÓN
// ═══════════════════════════════════════════════════════════════════

function menuInicializarSistema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  if (!ss.getSheetByName(SIM_SHEET_NAME)) {
    ui.alert(
      '⚠️ Simulador no encontrado',
      'No se encontró la hoja "' + SIM_SHEET_NAME + '". Ejecuta primero buildSimulador().',
      ui.ButtonSet.OK
    );
    return;
  }

  var hist = ss.getSheetByName(HIST_SHEET_NAME);
  if (hist) {
    var resp = ui.alert(
      'Historial ya existe',
      'La hoja "' + HIST_SHEET_NAME + '" ya existe. ¿Quieres reinicializarla? (Esto borrará todas las simulaciones guardadas)',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(hist);
  }

  crearHojaHistorial_(ss);
  ui.alert(
    '✅ Sistema inicializado',
    'Se creó la hoja "' + HIST_SHEET_NAME + '". Ya puedes guardar simulaciones desde el menú.',
    ui.ButtonSet.OK
  );
}

function crearHojaHistorial_(ss) {
  var ws = ss.insertSheet(HIST_SHEET_NAME);

  var headers = [
    'ID', 'Fecha', 'Autor', 'Nombre escenario', 'Notas',
    'Total Base', 'Total SIM', 'Diferencia', '% Cobertura',
    '# Accionistas con cambio', 'Valores SIM (JSON)', 'Métricas (JSON)',
  ];

  ws.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRng = ws.getRange(1, 1, 1, headers.length);
  headerRng.setBackground(CUALLI.dark)
    .setFontColor(CUALLI.yellow)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  ws.setRowHeight(1, 36);

  var widths = [180, 140, 200, 220, 280, 100, 100, 100, 100, 100, 300, 300];
  widths.forEach(function(w, i) { ws.setColumnWidth(i + 1, w); });

  ws.getRange('F:H').setNumberFormat('#,##0');
  ws.getRange('I:I').setNumberFormat('0.00%');
  ws.getRange('J:J').setNumberFormat('0');

  ws.getRange(1, 1, 1, headers.length).createFilter();
  ws.setFrozenRows(1);
  ws.setTabColor(CUALLI.yellow);
  ws.hideColumns(11, 2);

  SpreadsheetApp.flush();
}

// ═══════════════════════════════════════════════════════════════════
//  GUARDAR SIMULACIÓN
// ═══════════════════════════════════════════════════════════════════

function menuGuardarSimulacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sim = ss.getSheetByName(SIM_SHEET_NAME);

  if (!sim) {
    ui.alert('⚠️ No se encontró la hoja del simulador.');
    return;
  }

  var hist = ss.getSheetByName(HIST_SHEET_NAME);
  if (!hist) {
    var resp = ui.alert(
      'Sistema no inicializado',
      'No existe la hoja "' + HIST_SHEET_NAME + '". ¿Crearla ahora?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    crearHojaHistorial_(ss);
    hist = ss.getSheetByName(HIST_SHEET_NAME);
  }

  var nombreResp = ui.prompt(
    '💾 Guardar simulación',
    'Nombre del escenario (ej. "Reestructura Q4 2026", "Salida Mazzoni"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (nombreResp.getSelectedButton() !== ui.Button.OK) return;
  var nombre = nombreResp.getResponseText().trim();
  if (!nombre) {
    ui.alert('Debes proporcionar un nombre.');
    return;
  }

  var notasResp = ui.prompt(
    'Notas (opcional)',
    'Describe brevemente el escenario, motivo, acuerdo de referencia, etc.:',
    ui.ButtonSet.OK_CANCEL
  );
  if (notasResp.getSelectedButton() !== ui.Button.OK) return;
  var notas = notasResp.getResponseText().trim();

  var snapshot = capturarEstadoSimulador_(sim);

  var id = Utilities.getUuid();
  var fecha = new Date();
  var autor = Session.getActiveUser().getEmail() || 'desconocido';

  var fila = [
    id, fecha, autor, nombre, notas,
    snapshot.totalBase,
    snapshot.totalSim,
    snapshot.totalBase - snapshot.totalSim,
    snapshot.totalBase > 0 ? snapshot.totalSim / snapshot.totalBase : 0,
    snapshot.numAccionistasConCambio,
    JSON.stringify(snapshot.valoresSim),
    JSON.stringify(snapshot.metricas),
  ];

  hist.appendRow(fila);
  var ultimaFila = hist.getLastRow();
  hist.getRange(ultimaFila, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Escenario "' + nombre + '" guardado correctamente.',
    '✅ Simulación guardada',
    5
  );
}

// ═══════════════════════════════════════════════════════════════════
//  CAPTURAR ESTADO BÁSICO 
// ═══════════════════════════════════════════════════════════════════

function capturarEstadoSimulador_(sim) {
  var startRow = SIM_RANGE.startRow;
  var endRow = SIM_RANGE.endRow;
  var startCol = SIM_RANGE.startCol;
  var numCols = SIM_RANGE.numCols;
  var numRows = endRow - startRow + 1;

  var valoresSim = sim.getRange(startRow, startCol, numRows, numCols).getValues();
  var valoresBase = sim.getRange(startRow, 2, numRows, 6).getValues();
  var nombres = sim.getRange(startRow, 1, numRows, 1).getValues().map(function(r) { return r[0]; });

  var metricas = [];
  var numAccionistasConCambio = 0;

  for (var i = 0; i < numRows; i++) {
    var totalBaseFila = valoresBase[i].reduce(function(a, b) { return a + Number(b || 0); }, 0);
    var totalSimFila  = valoresSim[i].reduce(function(a, b) { return a + Number(b || 0); }, 0);
    var diferencia = totalSimFila - totalBaseFila;

    if (diferencia !== 0) numAccionistasConCambio++;

    metricas.push({
      nombre: nombres[i],
      totalBase: totalBaseFila,
      totalSim: totalSimFila,
      diferencia: diferencia,
    });
  }

  var totalBase = metricas.reduce(function(a, m) { return a + m.totalBase; }, 0);
  var totalSim  = metricas.reduce(function(a, m) { return a + m.totalSim; }, 0);

  return {
    valoresSim: valoresSim,
    metricas: metricas,
    totalBase: totalBase,
    totalSim: totalSim,
    numAccionistasConCambio: numAccionistasConCambio,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  CARGAR SIMULACIÓN 
// ═══════════════════════════════════════════════════════════════════

function menuCargarSimulacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var hist = ss.getSheetByName(HIST_SHEET_NAME);

  if (!hist) {
    ui.alert('No hay historial. Guarda primero al menos una simulación.');
    return;
  }

  var lastRow = hist.getLastRow();
  if (lastRow < 2) {
    ui.alert('El historial está vacío.');
    return;
  }

  var numFilas = Math.min(20, lastRow - 1);
  var desdeFila = Math.max(2, lastRow - numFilas + 1);
  var datos = hist.getRange(desdeFila, 1, lastRow - desdeFila + 1, 12).getValues();

  var template = HtmlService.createTemplate(HTML_CARGAR_MODAL);
  template.simulaciones = datos.map(function(r) {
    return {
      id: r[0],
      fecha: Utilities.formatDate(new Date(r[1]), 'America/Mexico_City', 'yyyy-MM-dd HH:mm'),
      autor: r[2],
      nombre: r[3],
      notas: r[4],
      totalBase: r[5],
      totalSim: r[6],
      diferencia: r[7],
      numCambios: r[9],
    };
  }).reverse();

  var html = template.evaluate().setWidth(800).setHeight(600);
  ui.showModalDialog(html, '📂 Cargar simulación guardada');
}

// ═══════════════════════════════════════════════════════════════════
//  EJECUTOR DE CARGA
// ═══════════════════════════════════════════════════════════════════

function ejecutarCarga(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(HIST_SHEET_NAME);
  var sim = ss.getSheetByName(SIM_SHEET_NAME);

  var datos = hist.getDataRange().getValues();
  var fila = null;
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === id) {
      fila = datos[i];
      break;
    }
  }

  if (!fila) return { ok: false, error: 'No se encontró la simulación.' };

  try {
    var valoresSim = JSON.parse(fila[10]);
    sim.getRange(SIM_RANGE.startRow, SIM_RANGE.startCol, valoresSim.length, valoresSim[0].length)
       .setValues(valoresSim);
    SpreadsheetApp.flush();

    return {
      ok: true,
      nombre: fila[3],
      autor: fila[2],
      fecha: Utilities.formatDate(new Date(fila[1]), 'America/Mexico_City', 'yyyy-MM-dd HH:mm'),
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════════
//  VER HISTORIAL
// ═══════════════════════════════════════════════════════════════════

function menuVerHistorial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(HIST_SHEET_NAME);
  if (!hist) {
    SpreadsheetApp.getUi().alert('No existe el historial. Ejecuta "Inicializar sistema" primero.');
    return;
  }
  ss.setActiveSheet(hist);
}

// ═══════════════════════════════════════════════════════════════════
//  HTML del modal de carga
// ═══════════════════════════════════════════════════════════════════

var HTML_CARGAR_MODAL = '<!DOCTYPE html>\n' +
'<html>\n' +
'<head>\n' +
'<base target="_top">\n' +
'<style>\n' +
'  * { box-sizing: border-box; }\n' +
'  body { font-family: "IBM Plex Sans", -apple-system, system-ui, sans-serif; margin: 0; padding: 20px; background: #F8F9FA; color: #1F2D3D; }\n' +
'  h2 { margin: 0 0 16px 0; color: #1F2D3D; border-bottom: 3px solid #FDB913; padding-bottom: 8px; font-size: 18px; }\n' +
'  .info { background: #FFF9E6; border-left: 4px solid #FDB913; padding: 10px 14px; margin-bottom: 16px; font-size: 12px; color: #7D6608; }\n' +
'  .lista { display: flex; flex-direction: column; gap: 8px; }\n' +
'  .item { background: white; border: 1px solid #E5E7EB; border-radius: 8px; padding: 14px 16px; cursor: pointer; transition: all 0.15s; }\n' +
'  .item:hover { border-color: #FDB913; box-shadow: 0 2px 8px rgba(253,185,19,0.15); transform: translateY(-1px); }\n' +
'  .nombre { font-weight: 700; font-size: 14px; color: #1F2D3D; margin-bottom: 4px; }\n' +
'  .meta { font-size: 11px; color: #6B7280; margin-bottom: 8px; }\n' +
'  .notas { font-size: 12px; color: #4B5563; font-style: italic; margin-bottom: 8px; }\n' +
'  .metricas { display: flex; gap: 16px; font-size: 11px; flex-wrap: wrap; }\n' +
'  .metricas span { color: #6B7280; }\n' +
'  .metricas b { color: #1F2D3D; font-weight: 600; }\n' +
'  .diff-pos { color: #C0392B; font-weight: 700; }\n' +
'  .diff-neg { color: #1E8449; font-weight: 700; }\n' +
'  .diff-zero { color: #6B7280; font-weight: 700; }\n' +
'  .empty { text-align: center; padding: 40px; color: #9CA3AF; }\n' +
'  .footer { margin-top: 16px; padding-top: 12px; border-top: 1px solid #E5E7EB; text-align: right; }\n' +
'  button { background: #6B7280; color: white; border: none; padding: 8px 18px; border-radius: 6px; cursor: pointer; font-family: inherit; font-size: 13px; }\n' +
'  #status { padding: 10px; margin-top: 12px; border-radius: 6px; font-size: 12px; display: none; }\n' +
'  .status-loading { display: block !important; background: #DBEAFE; color: #1E40AF; }\n' +
'  .status-ok { display: block !important; background: #D1FAE5; color: #065F46; }\n' +
'  .status-err { display: block !important; background: #FEE2E2; color: #991B1B; }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n' +
'  <h2>📂 Simulaciones guardadas</h2>\n' +
'  <div class="info">Selecciona un escenario para cargarlo. Se sobrescribirán los valores actuales de las columnas SIM (J:O). Los valores BASE no se tocan.</div>\n' +
'  <div class="lista">\n' +
'    <? if (simulaciones.length === 0) { ?>\n' +
'      <div class="empty">No hay simulaciones guardadas todavía.</div>\n' +
'    <? } else { ?>\n' +
'      <? for (var i = 0; i < simulaciones.length; i++) { var s = simulaciones[i]; ?>\n' +
'        <div class="item" onclick="cargar(\'<?= s.id ?>\', \'<?!= s.nombre.replace(/\'/g, "\\\\\'") ?>\')">\n' +
'          <div class="nombre"><?= s.nombre ?></div>\n' +
'          <div class="meta"><?= s.fecha ?> · <?= s.autor ?></div>\n' +
'          <? if (s.notas) { ?><div class="notas">"<?= s.notas ?>"</div><? } ?>\n' +
'          <div class="metricas">\n' +
'            <span>Base: <b><?= s.totalBase.toLocaleString() ?></b></span>\n' +
'            <span>SIM: <b><?= s.totalSim.toLocaleString() ?></b></span>\n' +
'            <span>Diferencia: <b class="<?= s.diferencia > 0 ? \'diff-pos\' : s.diferencia < 0 ? \'diff-neg\' : \'diff-zero\' ?>"><?= s.diferencia.toLocaleString() ?></b></span>\n' +
'            <span>Cambios: <b><?= s.numCambios ?></b></span>\n' +
'          </div>\n' +
'        </div>\n' +
'      <? } ?>\n' +
'    <? } ?>\n' +
'  </div>\n' +
'  <div id="status"></div>\n' +
'  <div class="footer"><button onclick="google.script.host.close()">Cerrar</button></div>\n' +
'<script>\n' +
'function cargar(id, nombre) {\n' +
'  if (!confirm("¿Cargar el escenario \\"" + nombre + "\\"?\\n\\nEsto sobrescribirá los valores SIM actuales del simulador.")) return;\n' +
'  var s = document.getElementById("status");\n' +
'  s.className = "status-loading";\n' +
'  s.textContent = "⏳ Cargando " + nombre + "…";\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(res) {\n' +
'      if (res.ok) {\n' +
'        s.className = "status-ok";\n' +
'        s.textContent = "✅ Cargado: \\"" + res.nombre + "\\" (guardado " + res.fecha + " por " + res.autor + ")";\n' +
'        setTimeout(function() { google.script.host.close(); }, 1500);\n' +
'      } else {\n' +
'        s.className = "status-err";\n' +
'        s.textContent = "❌ Error: " + res.error;\n' +
'      }\n' +
'    })\n' +
'    .withFailureHandler(function(err) {\n' +
'      s.className = "status-err";\n' +
'      s.textContent = "❌ Error: " + err.message;\n' +
'    })\n' +
'    .ejecutarCarga(id);\n' +
'}\n' +
'</script>\n' +
'</body>\n' +
'</html>';
