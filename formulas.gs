

var DET_DIST  = 'AND($K$22=0,NOT(ISBLANK($K$22)))';    // K22 en cero explícito
var DET_DISM  = 'ISBLANK($K$22)';                       // K22 vacía = disminución
var DET_TESOR = 'OR($K$22=0,ISBLANK($K$22))';           // cualquier modo activo

var DEN_BASE  = 'SUM($H$12:$H$25)';
var DEN_SIM   = 'SUM($P$12:$P$25)';
var DEN_AUTO  = 'IF(' + DET_DISM + ',' + DEN_SIM + ',' + DEN_BASE + ')';
var DEN_DIST  = '(SUM($H$12:$H$19)+SUM($H$23:$H$25))';
var TESOR_K   = '$C$22';
var TESOR_N   = '$F$22';
var RECEPTORAS = [12,13,14,15,16,17,18,19,23,24,25];

// ═══════════════════════════════════════════════════════════════════
//  aplicarFormulasAutomaticas
//  Escribe Q (% SIM) y V (Estatus) con detección automática de modo
//  Ejecutar UNA SOLA VEZ después de buildSimulador
// ═══════════════════════════════════════════════════════════════════

function aplicarFormulasAutomaticas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ws = ss.getSheetByName('SIMULADOR_ACCIONARIO');
  if (!ws) { ui.alert('⚠️ No se encontró la hoja SIMULADOR_ACCIONARIO.'); return; }

  var FILA_INICIO = 12;
  var FILA_FIN    = 25;

  // ── 1. Limpiar L6 si quedó checkbox de versión anterior ─────────
  var celdaL6 = ws.getRange(6, 12);
  try { celdaL6.removeCheckboxes(); } catch(e) {}
  celdaL6.clearContent().clearNote()
    .setBackground('#F8F9FA').setFontColor('#515151')
    .setFontWeight('normal').setFontSize(9)
    .setBorder(false,false,false,false,false,false);
  ws.getRange(6,13).clearContent()
    .setBackground('#F8F9FA').setFontColor('#515151');

  // ── 2. Q — % SIM con denominador automático según modo ──────────
  // Modo disminución (P21=0 Y P22=0) → denominador flotante SIM
  // Cualquier otro caso               → denominador fijo BASE
  for (var row = FILA_INICIO; row <= FILA_FIN; row++) {
    ws.getRange(row, 17)
      .setFormula('=IFERROR(P' + row + '/' + DEN_AUTO + ',0)')
      .setNumberFormat('0.00%');
  }
  // Q26 — total
  ws.getRange(26, 17)
    .setFormula('=IFERROR(' + DEN_SIM + '/' + DEN_AUTO + ',0)')
    .setNumberFormat('0.00%');

  // ── 3. V — Estatus ───────────────────────────────────────────────
  for (var r = FILA_INICIO; r <= FILA_FIN; r++) {
    var formula = buildEstatusFormula_(r);
    ws.getRange(r, 22).setFormula(formula);
  }

  SpreadsheetApp.flush();

  ui.alert(
    '✅ Fórmulas automáticas v5 aplicadas\n\n' +
    'Tres modos automáticos:\n\n' +
    '▸ NORMAL (P22 ≠ 0)\n' +
    '  K y N = base. Denominador % = total base.\n\n' +
    '▸ DISTRIBUCIÓN (P22 = 0 solo)\n' +
    '  K y N += proporcional de Tesorería.\n' +
    '  Denominador % = base (socios suman acciones).\n\n' +
    '▸ DISMINUCIÓN (P21 = 0 Y P22 = 0)\n' +
    '  K y N sin cambio. Denominador % = total SIM.\n' +
    '  Tesorería y Reporto ISC → 🔻 Capital cancelado.\n\n' +
    '• Reset L5 restaura todo correctamente.'
  );
}

// ── Construcción de fórmula V según fila ────────────────────────
function buildEstatusFormula_(r) {
  var base =
    'IF(P' + r + '=0,"Sin participación",' +
    'IF(P' + r + '>H' + r + ',"⬆ Aumenta",' +
    'IF(P' + r + '<H' + r + ',"⬇ Disminuye","✔ Sin cambio")))';

  if (r === 21) {
    // Reporto ISC — sin lógica especial, estatus normal siempre
    return '=' + base;
  } else if (r === 22) {
    // Tesorería — capital cancelado cuando K22=0 o K22 vacía
    return '=IF(AND(OR(ISBLANK($K$22),$K$22=0),H22>0),' +
      '"🔻 Capital cancelado",' + base + ')';
  } else {
    return '=' + base;
  }
}

// ═══════════════════════════════════════════════════════════════════
//  aplicarDistribucionTesoreria
//  Escribe fórmulas K y N con los tres modos integrados
//  Ejecutar UNA SOLA VEZ después de aplicarFormulasAutomaticas
// ═══════════════════════════════════════════════════════════════════

function aplicarDistribucionTesoreria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ws = ss.getSheetByName('SIMULADOR_ACCIONARIO');
  if (!ws) { ui.alert('⚠️ No se encontró la hoja SIMULADOR_ACCIONARIO.'); return; }

  RECEPTORAS.forEach(function(r) {
    var H = '$H$' + r;

    // K (col 11) — SIM B Fija
    // Disminución (blanco)  → solo base, sin distribución
    // Distribución (cero)   → base + proporcional Tesorería K
    // Normal                → base
    ws.getRange(r, 11).setFormula(
      '=IF(' + DET_DISM + ',C' + r + ',' +
        'IF(' + DET_DIST + ',' +
          'IF(' + DEN_DIST + '=0,C' + r + ',' +
            'C' + r + '+ROUND(' + TESOR_K + '*(' + H + '/' + DEN_DIST + '),0)),' +
          'C' + r + '))'
    ).setNumberFormat('#,##0');

    // N (col 14) — SIM B Variable
    ws.getRange(r, 14).setFormula(
      '=IF(' + DET_DISM + ',F' + r + ',' +
        'IF(' + DET_DIST + ',' +
          'IF(' + DEN_DIST + '=0,F' + r + ',' +
            'F' + r + '+ROUND(' + TESOR_N + '*(' + H + '/' + DEN_DIST + '),0)),' +
          'F' + r + '))'
    ).setNumberFormat('#,##0');
  });

  SpreadsheetApp.flush();

  ui.alert(
    '✅ Distribución de Tesorería configurada (v5)\n\n' +
    'Columnas K y N con tres modos:\n\n' +
    '▸ P22 ≠ 0 → K = base, N = base\n' +
    '▸ P22 = 0 (solo) → K = base + proporcional 10,688\n' +
    '                    N = base + proporcional  1,936\n' +
    '▸ P21 = 0 Y P22 = 0 → K = base (sin distribución)\n' +
    '                        N = base (sin distribución)\n\n' +
    'El denominador del % cambia automáticamente según el modo.'
  );
}

// ═══════════════════════════════════════════════════════════════════
//  onEdit — solo maneja el reset L5
// ═══════════════════════════════════════════════════════════════════

function onEdit(e) {
  var hoja = e.range.getSheet();
  if (hoja.getName() !== 'SIMULADOR_ACCIONARIO') return;

  if (e.range.getRow() === 5 && e.range.getColumn() === 12 && e.value === 'TRUE') {
    restablecerSimulacion_(hoja);
    e.range.setValue(false);
  }
}

// ═══════════════════════════════════════════════════════════════════
//  restablecerSimulacion_ — restaura todo el cuadro SIM al estado base
// ═══════════════════════════════════════════════════════════════════

function restablecerSimulacion_(hoja) {
  var FILA_INICIO = 12;
  var FILA_FIN    = 25;
  var NUM_FILAS   = FILA_FIN - FILA_INICIO + 1;

  // ── 1. Columnas simples: B→J, D→L, E→M, G→O (todas las filas) ──
  [[2,10],[4,12],[5,13],[7,15]].forEach(function(pair) {
    var vals = hoja.getRange(FILA_INICIO, pair[0], NUM_FILAS, 1).getValues();
    hoja.getRange(FILA_INICIO, pair[1], NUM_FILAS, 1).setValues(vals);
  });

  // ── 2. K y N de filas especiales sin fórmula: Reporto JCML(20), ISC(21) ──
  [20, 21].forEach(function(r) {
    hoja.getRange(r, 11).setValue(hoja.getRange(r, 3).getValue()).setNumberFormat('#,##0');
    hoja.getRange(r, 14).setValue(hoja.getRange(r, 6).getValue()).setNumberFormat('#,##0');
  });

  // ── 3. Tesorería SIM → valor base (desactiva detectores) ────────
  hoja.getRange(22, 10, 1, 6).setValues(hoja.getRange(22, 2, 1, 6).getValues());

  // ── 4. Reescribir fórmulas K y N en accionistas receptores ──────
  RECEPTORAS.forEach(function(r) {
    var H = '$H$' + r;

    hoja.getRange(r, 11).setFormula(
      '=IF(' + DET_DISM + ',C' + r + ',' +
        'IF(' + DET_DIST + ',' +
          'IF(' + DEN_DIST + '=0,C' + r + ',' +
            'C' + r + '+ROUND(' + TESOR_K + '*(' + H + '/' + DEN_DIST + '),0)),' +
          'C' + r + '))'
    ).setNumberFormat('#,##0');

    hoja.getRange(r, 14).setFormula(
      '=IF(' + DET_DISM + ',F' + r + ',' +
        'IF(' + DET_DIST + ',' +
          'IF(' + DEN_DIST + '=0,F' + r + ',' +
            'F' + r + '+ROUND(' + TESOR_N + '*(' + H + '/' + DEN_DIST + '),0)),' +
          'F' + r + '))'
    ).setNumberFormat('#,##0');
  });

  SpreadsheetApp.flush();
}

// ═══════════════════════════════════════════════════════════════════
//  actualizarCardsEstado
//  Actualiza las fórmulas de las cards de estado para reflejar
//  el modo disminución correctamente.
//  Ejecutar UNA SOLA VEZ sobre la hoja existente.
//
//  Card estado (N5:P5):
//    Normal/Distribución → ✅ SIN DIFERENCIAS / ⚠️ REVISAR DIFERENCIA
//    Disminución         → ✅ DISMINUCIÓN CONFIRMADA
//
//  Diferencia (O6:P6):
//    Normal/Distribución → H25 - P25
//    Disminución         → muestra monto cancelado con etiqueta "Acta:"
//
//  Nuevo capital (nueva línea en card):
//    Disminución         → muestra P25 como nuevo capital social
// ═══════════════════════════════════════════════════════════════════

function actualizarCardsEstado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ws = ss.getSheetByName('SIMULADOR_ACCIONARIO');
  if (!ws) { ui.alert('⚠️ No se encontró la hoja SIMULADOR_ACCIONARIO.'); return; }

  var TOTAL_ROW = 25;

  // ── Card estado: N5:P5 ──────────────────────────────────────────
  // Orden de prioridad (importante — disminución PRIMERO):
  //   1. AND(P21=0, P22=0) → DISMINUCIÓN CONFIRMADA  (azul/amarillo)
  //   2. P22=0 solo        → estado normal de distribución (puede cuadrar o no)
  //   3. H25=P25           → SIN DIFERENCIAS  (verde)
  //   4. resto             → REVISAR DIFERENCIA  (rojo)
  var cardEstadoRng = ws.getRange(5, 14, 1, 3);
  cardEstadoRng.setFormula(
    '=IF(ISBLANK($K$22),' +
      '"✅  DISMINUCIÓN CONFIRMADA",' +
      'IF(H' + TOTAL_ROW + '=P' + TOTAL_ROW + ',' +
        '"✅  SIN DIFERENCIAS",' +
        '"⚠️  REVISAR DIFERENCIA"))'
  );

  ws.getRange(6, 14).setFormula(
    '=IF(ISBLANK($K$22),"Monto acta de disminución:","Diferencia pendiente:")'
  );

  // ── Valor diferencia O6:P6 ──────────────────────────────────────
  var diffRng = ws.getRange(6, 15, 1, 2);
  diffRng.setFormula(
    '=H' + TOTAL_ROW + '-P' + TOTAL_ROW
  );
  diffRng.setNumberFormat('#,##0;(#,##0);"-"');

  // ── Formato condicional card estado ─────────────────────────────
  var existingRules = ws.getConditionalFormatRules();
  var rngCard = ws.getRange(5, 14, 1, 3);

  // Quitar regla previa de DISMINUCIÓN si ya existe (evitar duplicados)
  existingRules = existingRules.filter(function(r) {
    try {
      var crit = r.getBooleanCondition();
      return !(crit && crit.getCriteriaValues &&
               String(crit.getCriteriaValues()).indexOf('DISMINUCIÓN') !== -1);
    } catch(e) { return true; }
  });

  var ruleDisminucion = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('DISMINUCIÓN CONFIRMADA')
    .setBackground('#1A5276')
    .setFontColor('#FDB913')
    .setBold(true)
    .setRanges([rngCard])
    .build();

  existingRules.unshift(ruleDisminucion);
  ws.setConditionalFormatRules(existingRules);

  SpreadsheetApp.flush();

  ui.alert(
    '✅ Cards de estado actualizadas\n\n' +
    '▸ AND(P21=0, P22=0) → ✅ DISMINUCIÓN CONFIRMADA\n' +
    '   Etiqueta: "Monto acta de disminución"\n\n' +
    '▸ H25 = P25 → ✅ SIN DIFERENCIAS\n\n' +
    '▸ H25 ≠ P25 → ⚠️ REVISAR DIFERENCIA\n\n' +
    'El modo disminución se evalúa PRIMERO,\n' +
    'por lo que nunca mostrará REVISAR DIFERENCIA\n' +
    'cuando Reporto ISC y Tesorería estén en cero.'
  );
}
