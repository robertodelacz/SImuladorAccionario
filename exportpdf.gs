

var PDF_FOLDER_NAME = 'Cualli - PDFs Simulador Accionario';
var PDF_FOLDER_ID   = '';
var LOGO_URL        = 'https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png';
var LOGO_CACHE_KEY  = 'cualli_logo_b64_v1';
var LOGO_CACHE_TTL  = 21600;

// ═══════════════════════════════════════════════════════════════════
//  PUNTO DE ENTRADA
// ═══════════════════════════════════════════════════════════════════

function menuExportarPDF() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var sim = ss.getSheetByName(SIM_SHEET_NAME);

  if (!sim) { ui.alert('No se encontró la hoja del simulador.'); return; }

  var nombreResp = ui.prompt(
    '📄 Exportar PDF ejecutivo',
    'Nombre del escenario para el PDF (opcional):',
    ui.ButtonSet.OK_CANCEL
  );
  if (nombreResp.getSelectedButton() !== ui.Button.OK) return;
  var nombreEscenario = nombreResp.getResponseText().trim();

  var notasResp = ui.prompt(
    'Notas (opcional)',
    'Descripción / contexto del escenario:',
    ui.ButtonSet.OK_CANCEL
  );
  if (notasResp.getSelectedButton() !== ui.Button.OK) return;
  var notas = notasResp.getResponseText().trim();

  try {
    ss.toast('Generando PDF…', '⏳ En proceso', -1);

    var snapshot = capturarEstadoDetallado_(sim);
    var html     = construirHTMLPDF_(snapshot, nombreEscenario, notas);
    var pdfBlob  = convertirHTMLaPDF_(html, nombreEscenario);
    var folder   = obtenerCarpetaPDFs_();
    var archivo  = folder.createFile(pdfBlob);

    ss.toast('PDF generado', '✅ Listo', 3);

    var url   = archivo.getUrl();
    var html2 = HtmlService.createHtmlOutput(
      '<div style="font-family:\'IBM Plex Sans\',sans-serif;padding:24px;">' +
        '<h2 style="color:#1F2D3D;border-bottom:3px solid #FDB913;padding-bottom:8px;">✅ PDF generado</h2>' +
        '<p style="font-size:13px;color:#515151;margin:14px 0;">' +
          'Documento guardado en <b>"' + PDF_FOLDER_NAME + '"</b>.' +
        '</p>' +
        '<p style="margin:18px 0;">' +
          '<a href="' + url + '" target="_blank" style="display:inline-block;background:#FDB913;' +
          'color:#1F2D3D;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:700;">' +
          '📄 Abrir PDF</a>' +
        '</p>' +
        '<p style="font-size:11px;color:#9CA3AF;margin-top:16px;">Archivo: <code>' +
          archivo.getName() + '</code></p>' +
      '</div>'
    ).setWidth(450).setHeight(260);
    ui.showModalDialog(html2, 'PDF Exportado');

  } catch (err) {
    ui.alert('❌ Error al generar el PDF:\n\n' + err.message + '\n\nStack:\n' + err.stack);
  }
}

// ═══════════════════════════════════════════════════════════════════
//  CAPTURA DETALLADA — lee L6 para respetar el modo activo
// ═══════════════════════════════════════════════════════════════════

function capturarEstadoDetallado_(sim) {
  var startRow = SIM_RANGE.startRow;
  var endRow   = SIM_RANGE.endRow;  // fila 25 (nueva estructura)
  var numRows  = endRow - startRow + 1;

  // Fila 21=ReportoISC (idx 9), Fila 22=Tesoreria (idx 10)
  var valoresSimPreview = sim.getRange(startRow, SIM_RANGE.startCol, numRows, SIM_RANGE.numCols).getValues();
  var idxReportoISC = 21 - startRow;
  var idxTesoreria  = 22 - startRow;
  var sumReportoISC = valoresSimPreview[idxReportoISC]
    ? valoresSimPreview[idxReportoISC].reduce(function(a,b){return a+Number(b||0);},0) : 0;
  var sumTesoreria  = valoresSimPreview[idxTesoreria]
    ? valoresSimPreview[idxTesoreria].reduce(function(a,b){return a+Number(b||0);},0)  : 0;
  var modoDistribucion = (sumTesoreria === 0 && sumReportoISC !== 0);
  var modoDisminucion  = (sumTesoreria === 0 && sumReportoISC === 0);

  var valoresSim  = sim.getRange(startRow, SIM_RANGE.startCol, numRows, SIM_RANGE.numCols).getValues();
  var valoresBase = sim.getRange(startRow, 2, numRows, 6).getValues();
  var nombres     = sim.getRange(startRow, 1, numRows, 1).getValues()
                       .map(function(r) { return String(r[0] || ''); });

  var detalle = [];
  var numAccionistasConCambio = 0;
  var totalBaseFija = 0, totalBaseVar = 0;
  var totalSimFija  = 0, totalSimVar  = 0;
  var totalAccionesMovidas = 0;

  for (var i = 0; i < numRows; i++) {
    var baseFija = { A: num_(valoresBase[i][0]), B: num_(valoresBase[i][1]), C: num_(valoresBase[i][2]) };
    var baseVar  = { A: num_(valoresBase[i][3]), B: num_(valoresBase[i][4]), C: num_(valoresBase[i][5]) };
    var simFija  = { A: num_(valoresSim[i][0]),  B: num_(valoresSim[i][1]),  C: num_(valoresSim[i][2]) };
    var simVar   = { A: num_(valoresSim[i][3]),  B: num_(valoresSim[i][4]),  C: num_(valoresSim[i][5]) };

    var totalBase = baseFija.A + baseFija.B + baseFija.C + baseVar.A + baseVar.B + baseVar.C;
    var totalSim  = simFija.A  + simFija.B  + simFija.C  + simVar.A  + simVar.B  + simVar.C;

    totalBaseFija += baseFija.A + baseFija.B + baseFija.C;
    totalBaseVar  += baseVar.A  + baseVar.B  + baseVar.C;
    totalSimFija  += simFija.A  + simFija.B  + simFija.C;
    totalSimVar   += simVar.A   + simVar.B   + simVar.C;

    if (totalSim !== totalBase) {
      numAccionistasConCambio++;
      if (totalSim > totalBase) totalAccionesMovidas += (totalSim - totalBase);
    }

    detalle.push({
      nombre: nombres[i],
      baseFija: baseFija, baseVar: baseVar,
      simFija: simFija,   simVar: simVar,
      totalBase: totalBase, totalSim: totalSim,
    });
  }

  var totalBaseGlobal = totalBaseFija + totalBaseVar;
  var totalSimGlobal  = totalSimFija  + totalSimVar;

  // Denominador: flotante solo en modo disminución
  var denominador = modoDisminucion ? totalSimGlobal : totalBaseGlobal;

  return {
    detalle: detalle,
    modoDistribucion: modoDistribucion,
    modoDisminucion:  modoDisminucion,
    totales: {
      base:        totalBaseGlobal,
      sim:         totalSimGlobal,
      baseFija:    totalBaseFija,
      baseVar:     totalBaseVar,
      simFija:     totalSimFija,
      simVar:      totalSimVar,
      denominador: denominador,
      cols:        calcularTotalesPorColumna_(detalle),
    },
    numAccionistasConCambio: numAccionistasConCambio,
    totalAccionesMovidas:    totalAccionesMovidas,
    pctCapitalMovimiento:    totalBaseGlobal > 0 ? totalAccionesMovidas / totalBaseGlobal : 0,
  };
}

function num_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function calcularTotalesPorColumna_(detalle) {
  var t = {
    baseFijaA:0, baseFijaB:0, baseFijaC:0,
    baseVarA:0,  baseVarB:0,  baseVarC:0,
    simFijaA:0,  simFijaB:0,  simFijaC:0,
    simVarA:0,   simVarB:0,   simVarC:0,
  };
  detalle.forEach(function(d) {
    t.baseFijaA += d.baseFija.A; t.baseFijaB += d.baseFija.B; t.baseFijaC += d.baseFija.C;
    t.baseVarA  += d.baseVar.A;  t.baseVarB  += d.baseVar.B;  t.baseVarC  += d.baseVar.C;
    t.simFijaA  += d.simFija.A;  t.simFijaB  += d.simFija.B;  t.simFijaC  += d.simFija.C;
    t.simVarA   += d.simVar.A;   t.simVarB   += d.simVar.B;   t.simVarC   += d.simVar.C;
  });
  return t;
}

// ═══════════════════════════════════════════════════════════════════
//  LOGO con caché
// ═══════════════════════════════════════════════════════════════════

function obtenerLogoBase64_() {
  var cache = CacheService.getScriptCache();
  var b64   = cache.get(LOGO_CACHE_KEY);
  if (b64) return b64;
  try {
    var resp = UrlFetchApp.fetch(LOGO_URL, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    b64 = Utilities.base64Encode(resp.getBlob().getBytes());
    if (b64.length < 95000) cache.put(LOGO_CACHE_KEY, b64, LOGO_CACHE_TTL);
    return b64;
  } catch(e) { return null; }
}

function obtenerCarpetaPDFs_() {
  if (PDF_FOLDER_ID) {
    try { return DriveApp.getFolderById(PDF_FOLDER_ID); } catch(e) {}
  }
  var iter = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(PDF_FOLDER_NAME);
}

function convertirHTMLaPDF_(html, nombreEscenario) {
  var fecha = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd_HHmm');
  var slug  = nombreEscenario
    ? nombreEscenario.replace(/[^a-zA-Z0-9 _-]/g,'').replace(/\s+/g,'_').substring(0,50)
    : 'estado_actual';
  var filename = 'Cualli_Simulador_' + slug + '_' + fecha + '.pdf';
  return Utilities.newBlob(html, 'text/html', filename + '.html')
    .getAs('application/pdf').setName(filename);
}

function limpiarCacheLogo() {
  CacheService.getScriptCache().remove(LOGO_CACHE_KEY);
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('El logo se descargará de nuevo en el próximo PDF.', '🔄 Caché limpiado', 4);
}

// ═══════════════════════════════════════════════════════════════════
//  CONSTRUCCIÓN DEL HTML
// ═══════════════════════════════════════════════════════════════════

function construirHTMLPDF_(snapshot, nombreEscenario, notas) {
  var ahora     = new Date();
  var fechaLarga = Utilities.formatDate(ahora, 'America/Mexico_City', "d 'de' MMMM 'de' yyyy");
  var horaCorta  = Utilities.formatDate(ahora, 'America/Mexico_City', 'HH:mm');
  var autor      = Session.getActiveUser().getEmail() || 'usuario';
  var idDoc      = Utilities.getUuid().substring(0, 13);

  var T          = snapshot.totales;
  var diferencia = T.base - T.sim;
  var cobertura  = T.base > 0 ? T.sim / T.base : 0;
  var numCambios = snapshot.numAccionistasConCambio;
  var totalPos   = snapshot.detalle.length;

  var logoB64  = obtenerLogoBase64_();
  var logoHTML = logoB64
    ? '<span class="logo-img"><img src="data:image/png;base64,' + logoB64 + '" alt="Cualli"/></span>'
    : '<span class="logo-fallback">cualli</span>';

  // Sección escenario
  var seccionEscenario = nombreEscenario
    ? '<div class="escenario-line">' +
        '<span class="esc-nombre">📋 Escenario: ' + escaparHTML_(nombreEscenario) + '</span>' +
        (notas ? '<span class="esc-notas">"' + escaparHTML_(notas) + '"</span>' : '') +
      '</div>'
    : '';

  // Banner según modo activo
  var bannerModo = '';
  if (snapshot.modoDisminucion) {
    bannerModo = '<div style="background:#1A5276;color:#FDB913;padding:4px 12px;font-weight:700;' +
      'font-size:8pt;margin-bottom:6px;border-radius:3px;letter-spacing:0.3px;">' +
      '⚖️  MODO: DISMINUCIÓN DE CAPITAL — Capital reducido. Reporto ISC y Tesorería cancelados. ' +
      'Los % reflejan el nuevo capital social.</div>';
  } else if (snapshot.modoDistribucion) {
    bannerModo = '<div style="background:#1E8449;color:#FFFFFF;padding:4px 12px;font-weight:700;' +
      'font-size:8pt;margin-bottom:6px;border-radius:3px;letter-spacing:0.3px;">' +
      '📤  MODO: DISTRIBUCIÓN — Las acciones de Tesorería se distribuyen proporcionalmente ' +
      'entre los accionistas activos. Capital total sin cambio.</div>';
  }

  // Filas de tablas — ahora reciben snapshot completo
  var filasSup = snapshot.detalle.map(function(d, idx) {
    return construirFilaSup_(d, idx, snapshot);
  }).join('');
  var filasInf = snapshot.detalle.map(function(d, idx) {
    return construirFilaInf_(d, idx, snapshot);
  }).join('');

  var cols = T.cols;

  // Texto de diferencia y etiquetas según modo
  var diffSubtexto, kpiDiffLabel, kpiDiffSub, coberturaDisplay;

  if (snapshot.modoDisminucion) {
    diffSubtexto   = '🔻 Capital cancelado';
    kpiDiffLabel   = 'Monto Acta';
    kpiDiffSub     = 'Disminución de capital';
    coberturaDisplay = '100.00%';
  } else if (snapshot.modoDistribucion) {
    diffSubtexto   = diferencia === 0 ? '✓ distribuido' : 'pendiente ' + formatearNum_(Math.abs(diferencia));
    kpiDiffLabel   = 'Acciones distribuidas';
    kpiDiffSub     = 'De Tesorería a socios';
    coberturaDisplay = formatearPct_(cobertura);
  } else {
    diffSubtexto   = diferencia === 0 ? '✓ balanceada'
      : (diferencia > 0 ? 'faltan ' + formatearNum_(diferencia) : 'sobran ' + formatearNum_(Math.abs(diferencia)));
    kpiDiffLabel   = 'Diferencia';
    kpiDiffSub     = diferencia === 0 ? 'Simulación balanceada' : 'Pendiente de asignar';
    coberturaDisplay = formatearPct_(cobertura);
  }

  return '<!DOCTYPE html>\n' +
'<html lang="es">\n' +
'<head>\n' +
'<meta charset="UTF-8">\n' +
'<title>Simulador Accionario - Cualli</title>\n' +
'<style>\n' +
'  * { box-sizing: border-box; margin: 0; padding: 0; }\n' +
'  @page { size: A4 landscape; margin: 7mm 9mm; }\n' +
'  html, body { font-family: "Helvetica Neue", Arial, sans-serif; color: #1F2D3D;\n' +
'    background: #FFFFFF; font-size: 8pt; line-height: 1.2;\n' +
'    -webkit-print-color-adjust: exact; print-color-adjust: exact; }\n' +
'  .header { display: table; width: 100%; padding-bottom: 5px;\n' +
'    border-bottom: 2px solid #FDB913; margin-bottom: 6px; }\n' +
'  .header-left, .header-right { display: table-cell; vertical-align: middle; }\n' +
'  .header-right { text-align: right; font-size: 7pt; color: #515151; }\n' +
'  .logo-img { display: inline-block; vertical-align: middle; background: #FFFFFF;\n' +
'    padding: 4px 10px; border: 2px solid #FDB913; border-radius: 6px;\n' +
'    margin-right: 11px; line-height: 0; }\n' +
'  .logo-img img { display: inline-block; height: 22px; width: auto; }\n' +
'  .logo-fallback { display: inline-block; vertical-align: middle; background: #FDB913;\n' +
'    color: #2E2E2E; padding: 5px 13px; border-radius: 6px; font-weight: 800;\n' +
'    font-size: 14pt; margin-right: 11px; font-style: italic; }\n' +
'  .titles { display: inline-block; vertical-align: middle; }\n' +
'  .titles h1 { font-size: 12pt; font-weight: 800; color: #2E2E2E; margin-bottom: 1px; }\n' +
'  .titles h2 { font-size: 7pt; font-weight: 500; color: #515151;\n' +
'    text-transform: uppercase; letter-spacing: 1.1px; }\n' +
'  .header-right .fecha { font-weight: 700; color: #2E2E2E; font-size: 8.5pt; margin-bottom: 1px; }\n' +
'  .header-right .meta-line { color: #6B7280; font-size: 6.8pt; }\n' +
'  .escenario-line { background: #FFF9E6; border-left: 3px solid #FDB913;\n' +
'    padding: 3px 9px; margin-bottom: 6px; font-size: 7.5pt; }\n' +
'  .esc-nombre { font-weight: 700; color: #1F2D3D; display: inline-block; margin-right: 7px; }\n' +
'  .esc-notas { font-style: italic; color: #7D6608; }\n' +
'  table.kpi-row { width: 100%; border-collapse: separate; border-spacing: 5px 0;\n' +
'    margin: 0 -5px 6px -5px; }\n' +
'  td.kpi-card { background: #F8F9FA; border: 1px solid #E5E7EB; border-radius: 6px;\n' +
'    padding: 5px 8px; vertical-align: middle; }\n' +
'  td.kpi-card.grande { width: 12.5%; }\n' +
'  td.kpi-card.grande.base   { border-left: 3px solid #FDB913; }\n' +
'  td.kpi-card.grande.sim    { border-left: 3px solid #2E2E2E; }\n' +
'  td.kpi-card.grande.diff   { border-left: 3px solid #FDB913; }\n' +
'  td.kpi-card.grande.cambios{ border-left: 3px solid #2E2E2E; }\n' +
'  td.kpi-card.grande .label { font-size: 6.3pt; color: #6B7280; text-transform: uppercase;\n' +
'    letter-spacing: 0.5px; font-weight: 600; margin-bottom: 1px; line-height: 1; }\n' +
'  td.kpi-card.grande .value { font-size: 13pt; font-weight: 800; color: #2E2E2E;\n' +
'    line-height: 1; margin: 1px 0; }\n' +
'  td.kpi-card.grande .sub   { font-size: 6pt; color: #9CA3AF; line-height: 1.1; }\n' +
'  td.kpi-card.grande.diff .value   { color: #C99100; }\n' +
'  td.kpi-card.mini { width: 16.66%; }\n' +
'  td.kpi-card.mini.movidas  { border-left: 3px solid #FDB913; }\n' +
'  td.kpi-card.mini.capital  { border-left: 3px solid #515151; }\n' +
'  td.kpi-card.mini.afectados{ border-left: 3px solid #FDB913; }\n' +
'  .mini-row { display: table; width: 100%; }\n' +
'  .mini-icon-cell { display: table-cell; vertical-align: middle; width: 24px; padding-right: 4px; }\n' +
'  .mini-icon { font-size: 14pt; line-height: 1; color: #6B7280; }\n' +
'  .mini-text-cell { display: table-cell; vertical-align: middle; }\n' +
'  td.kpi-card.mini .label { font-size: 6.3pt; color: #6B7280; text-transform: uppercase;\n' +
'    letter-spacing: 0.5px; font-weight: 600; line-height: 1; margin-bottom: 1px; }\n' +
'  td.kpi-card.mini .value { font-size: 11pt; font-weight: 800; line-height: 1;\n' +
'    margin: 1px 0; color: #2E2E2E; }\n' +
'  td.kpi-card.mini .sub { font-size: 6pt; color: #9CA3AF; line-height: 1.1; }\n' +
'  td.kpi-card.mini.movidas .value  { color: #C99100; }\n' +
'  td.kpi-card.mini.movidas .mini-icon { color: #C99100; }\n' +
'  td.kpi-card.mini.capital .value  { color: #515151; }\n' +
'  td.kpi-card.mini.afectados .value{ color: #C99100; }\n' +
'  td.kpi-card.mini.afectados .mini-icon { color: #C99100; }\n' +
'  .frac-secondary { font-size: 8pt; color: #9CA3AF; font-weight: 600; }\n' +
'  .seccion-titulo { font-size: 7.8pt; font-weight: 700; color: #2E2E2E;\n' +
'    text-transform: uppercase; letter-spacing: 0.7px; margin-bottom: 3px;\n' +
'    padding-bottom: 2px; border-bottom: 1px solid #E5E7EB; }\n' +
'  .seccion-titulo .sub { font-weight: 400; color: #6B7280; text-transform: none;\n' +
'    letter-spacing: 0; margin-left: 6px; font-size: 7pt; }\n' +
'  table.principal { width: 100%; border-collapse: collapse; font-size: 6.3pt;\n' +
'    margin-bottom: 7px; table-layout: fixed; }\n' +
'  table.principal col.col-acc   { width: 18%; }\n' +
'  table.principal col.col-tipo  { width: 5.2%; }\n' +
'  table.principal col.col-total { width: 6.4%; }\n' +
'  table.principal col.col-pct   { width: 5.6%; }\n' +
'  table.principal thead tr.macro th { padding: 2px; font-size: 6.2pt; font-weight: 700;\n' +
'    text-transform: uppercase; letter-spacing: 0.4px; text-align: center; }\n' +
'  table.principal thead tr.macro th.macro-acc  { background: #2E2E2E; color: #FDB913; }\n' +
'  table.principal thead tr.macro th.macro-base { background: #FDB913; color: #2E2E2E; }\n' +
'  table.principal thead tr.macro th.macro-sim  { background: #FDB913; color: #2E2E2E; }\n' +
'  table.principal thead tr.grupo th { padding: 2px 1px; font-size: 5.9pt; font-weight: 700;\n' +
'    text-align: center; text-transform: uppercase; letter-spacing: 0.3px; }\n' +
'  table.principal thead tr.grupo th.gr-fija-base  { background: #FFF4D6; }\n' +
'  table.principal thead tr.grupo th.gr-var-base   { background: #FFFBEC; }\n' +
'  table.principal thead tr.grupo th.gr-total-base { background: #2E2E2E; color: #FDB913; }\n' +
'  table.principal thead tr.grupo th.gr-fija-sim   { background: #FFF4D6; }\n' +
'  table.principal thead tr.grupo th.gr-var-sim    { background: #FFFBEC; }\n' +
'  table.principal thead tr.grupo th.gr-total-sim  { background: #2E2E2E; color: #FDB913; }\n' +
'  table.principal thead tr.cols th { background: #F5F5F5; color: #515151; padding: 1.5px 1px;\n' +
'    font-size: 5.7pt; font-weight: 700; text-transform: uppercase;\n' +
'    border-bottom: 2px solid #D1D5DB; text-align: center; }\n' +
'  table.principal th.sep, table.principal td.sep { border-left: 2px solid #FDB913; }\n' +
'  table.principal tbody td { padding: 1.5px 2px; border-bottom: 1px solid #F0F1F3;\n' +
'    text-align: center; vertical-align: middle; font-size: 6.6pt; }\n' +
'  table.principal tbody td.acc-name { text-align: left; padding-left: 5px;\n' +
'    font-weight: 600; color: #2E2E2E; font-size: 6.6pt; }\n' +
'  table.principal tbody td.acc-name .rfc { display: block; font-size: 5.3pt;\n' +
'    color: #9CA3AF; font-weight: 400; }\n' +
'  table.principal tbody td.fija-base  { background: #FFFFFF; }\n' +
'  table.principal tbody td.var-base   { background: #FAFAFA; font-style: italic; }\n' +
'  table.principal tbody td.total-base { background: #FDB913; font-weight: 700; }\n' +
'  table.principal tbody td.fija-sim   { background: #FFFFFF; }\n' +
'  table.principal tbody td.var-sim    { background: #FAFAFA; font-style: italic; }\n' +
'  table.principal tbody td.total-sim  { background: #FDB913; font-weight: 700; }\n' +
'  table.principal tbody td.pct-base   { background: #FFF4D6; font-weight: 600; }\n' +
'  table.principal tbody td.pct-sim    { background: #FFF4D6; font-weight: 600; }\n' +
'  table.principal tbody tr.especial td { background: #F5F5F5; font-style: italic; }\n' +
'  table.principal tbody tr.especial td.acc-name { color: #515151; }\n' +
'  table.principal tbody tr.especial td.total-base { background: #C99100; color: #FFFFFF; }\n' +
'  table.principal tbody tr.especial td.total-sim  { background: #C99100; color: #FFFFFF; }\n' +
'  table.principal tbody tr.especial td.pct-base   { background: #FFE5A3; color: #515151; }\n' +
'  table.principal tbody tr.especial td.pct-sim    { background: #FFE5A3; color: #515151; }\n' +
'  table.principal tfoot tr.total td { background: #2E2E2E; color: #FDB913; font-weight: 800;\n' +
'    padding: 3px 2px; font-size: 6.8pt; border-top: 2px solid #FDB913; }\n' +
'  table.principal tfoot tr.total td.acc-name { text-align: left; padding-left: 5px;\n' +
'    text-transform: uppercase; }\n' +
'  table.detalle-variacion { width: 100%; border-collapse: collapse; font-size: 6.8pt;\n' +
'    table-layout: fixed; }\n' +
'  table.detalle-variacion col.col-acc     { width: 20%; }\n' +
'  table.detalle-variacion col.col-num     { width: 8%; }\n' +
'  table.detalle-variacion col.col-var-acc { width: 11%; }\n' +
'  table.detalle-variacion col.col-var-pct { width: 8%; }\n' +
'  table.detalle-variacion col.col-pos     { width: 14%; }\n' +
'  table.detalle-variacion col.col-pct-fin { width: 14%; }\n' +
'  table.detalle-variacion col.col-estatus { width: 12%; }\n' +
'  table.detalle-variacion thead th { background: #1F2D3D; color: #FDB913; padding: 3px 4px;\n' +
'    font-size: 6.3pt; font-weight: 700; text-transform: uppercase;\n' +
'    letter-spacing: 0.5px; border-bottom: 2px solid #FDB913; }\n' +
'  table.detalle-variacion thead th.left    { text-align: left; padding-left: 7px; }\n' +
'  table.detalle-variacion thead th.var-acc { background: #1E8449; color: white; }\n' +
'  table.detalle-variacion thead th.pos-fin { background: #2E75B6; color: white; }\n' +
'  table.detalle-variacion tbody td { padding: 2px 4px; border-bottom: 1px solid #E5E7EB;\n' +
'    vertical-align: middle; text-align: center; font-size: 6.8pt; }\n' +
'  table.detalle-variacion tbody td.acc-name { text-align: left; padding-left: 7px;\n' +
'    font-weight: 600; color: #1F2D3D; }\n' +
'  table.detalle-variacion tbody td.acc-name .rfc { display: inline-block; font-size: 5.6pt;\n' +
'    color: #9CA3AF; font-weight: 400; margin-left: 3px; }\n' +
'  table.detalle-variacion tbody td.acc-name .tag-especial { display: inline-block;\n' +
'    background: #455A64; color: white; font-size: 5.3pt; padding: 1px 4px;\n' +
'    border-radius: 2px; margin-left: 4px; font-weight: 700; vertical-align: middle; }\n' +
'  table.detalle-variacion tbody td.num      { color: #1F2D3D; }\n' +
'  table.detalle-variacion tbody td.num.muted{ color: #9CA3AF; }\n' +
'  table.detalle-variacion tbody td.var-acc-cell { font-weight: 800; font-size: 7.7pt; }\n' +
'  table.detalle-variacion tbody td.var-acc-cell.up   { background: #D5F5E3; color: #1E8449; }\n' +
'  table.detalle-variacion tbody td.var-acc-cell.down { background: #FADBD8; color: #C0392B; }\n' +
'  table.detalle-variacion tbody td.var-acc-cell.eq   { background: #F8F9FA; color: #9CA3AF; }\n' +
'  table.detalle-variacion tbody td.var-pct-cell { font-weight: 700; }\n' +
'  table.detalle-variacion tbody td.var-pct-cell.up   { color: #1E8449; }\n' +
'  table.detalle-variacion tbody td.var-pct-cell.down { color: #C0392B; }\n' +
'  table.detalle-variacion tbody td.var-pct-cell.eq   { color: #9CA3AF; }\n' +
'  table.detalle-variacion tbody td.pos-final  { color: #1F2D3D; font-weight: 600; }\n' +
'  table.detalle-variacion tbody td.pct-final  { color: #1F2D3D; font-weight: 600; }\n' +
'  .arrow-up   { color: #1E8449; font-weight: 800; margin: 0 3px; }\n' +
'  .arrow-down { color: #C0392B; font-weight: 800; margin: 0 3px; }\n' +
'  .arrow-eq   { color: #9CA3AF; font-weight: 600; margin: 0 3px; }\n' +
'  table.detalle-variacion tbody td.estatus-cell span { display: inline-block;\n' +
'    padding: 1.5px 7px; border-radius: 9px; font-size: 6.2pt; font-weight: 700; }\n' +
'  .badge-aumenta   { background: #D5F5E3; color: #1E8449; border: 1px solid #A9DFBF; }\n' +
'  .badge-disminuye { background: #FADBD8; color: #C0392B; border: 1px solid #F1948A; }\n' +
'  .badge-sincambio { background: #EBF5FB; color: #2E75B6; border: 1px solid #AED6F1; }\n' +
'  .badge-sinpart   { background: #F2F3F4; color: #7F8C8D; border: 1px solid #D5DBDB; }\n' +
'  .badge-cancelado { background: #1A5276; color: #FFFFFF; border: 1px solid #154360; }\n' +
'  table.detalle-variacion tbody tr.especial td { background: #F8F9FA; }\n' +
'  table.detalle-variacion tbody tr.especial td.acc-name { color: #455A64; }\n' +
'  table.detalle-variacion tfoot tr.total td { background: #1F2D3D; color: #FDB913;\n' +
'    font-weight: 800; padding: 3px 4px; font-size: 7.2pt; border-top: 2px solid #FDB913; }\n' +
'  table.detalle-variacion tfoot tr.total td.acc-name { text-align: left; padding-left: 7px;\n' +
'    text-transform: uppercase; }\n' +
'  .footer { margin-top: 5px; padding-top: 3px; border-top: 1px solid #E5E7EB;\n' +
'    display: table; width: 100%; font-size: 6.2pt; color: #9CA3AF; }\n' +
'  .footer-l { display: table-cell; text-align: left;   font-weight: 600; }\n' +
'  .footer-c { display: table-cell; text-align: center; font-weight: 600; color: #515151; }\n' +
'  .footer-r { display: table-cell; text-align: right;  font-style: italic; }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n' +
'<div class="header">\n' +
'  <div class="header-left">' + logoHTML +
'    <div class="titles"><h1>Simulador de Cuadro Accionario</h1>' +
'    <h2>Financiera Cualli · Gobierno Corporativo</h2></div>\n' +
'  </div>\n' +
'  <div class="header-right">\n' +
'    <div class="fecha">' + fechaLarga + '</div>\n' +
'    <div class="meta-line">Generado: ' + fechaLarga + ' ' + horaCorta + ' · Por: ' + escaparHTML_(autor) + '</div>\n' +
'  </div>\n' +
'</div>\n' +
seccionEscenario + '\n' +
bannerModo + '\n' +
'<table class="kpi-row"><tr>\n' +
'  <td class="kpi-card grande base">\n' +
'    <div class="label">Total Base</div>\n' +
'    <div class="value">' + formatearNum_(T.base) + '</div>\n' +
'    <div class="sub">F: ' + formatearNum_(T.baseFija) + ' · V: ' + formatearNum_(T.baseVar) + '</div>\n' +
'  </td>\n' +
'  <td class="kpi-card grande sim">\n' +
'    <div class="label">Total Simulado</div>\n' +
'    <div class="value">' + formatearNum_(T.sim) + '</div>\n' +
'    <div class="sub">F: ' + formatearNum_(T.simFija) + ' · V: ' + formatearNum_(T.simVar) + '</div>\n' +
'  </td>\n' +
'  <td class="kpi-card grande diff">\n' +
'    <div class="label">' + kpiDiffLabel + '</div>\n' +
'    <div class="value">' + formatearNum_(Math.abs(diferencia)) + '</div>\n' +
'    <div class="sub">' + kpiDiffSub + '</div>\n' +
'  </td>\n' +
'  <td class="kpi-card grande cambios">\n' +
'    <div class="label">Con cambio</div>\n' +
'    <div class="value">' + numCambios + '</div>\n' +
'    <div class="sub">de ' + totalPos + ' posiciones</div>\n' +
'  </td>\n' +
'  <td class="kpi-card mini movidas">\n' +
'    <div class="mini-row"><div class="mini-icon-cell"><span class="mini-icon">⇄</span></div>\n' +
'    <div class="mini-text-cell"><div class="label">Acciones movidas</div>\n' +
'    <div class="value">' + formatearNum_(snapshot.totalAccionesMovidas) + '</div>\n' +
'    <div class="sub">Movimiento absoluto</div></div></div>\n' +
'  </td>\n' +
'  <td class="kpi-card mini capital">\n' +
'    <div class="mini-row"><div class="mini-icon-cell"><span class="mini-icon">%</span></div>\n' +
'    <div class="mini-text-cell"><div class="label">Capital en movimiento</div>\n' +
'    <div class="value">' + formatearPct_(snapshot.pctCapitalMovimiento) + '</div>\n' +
'    <div class="sub">Sobre el total base</div></div></div>\n' +
'  </td>\n' +
'  <td class="kpi-card mini afectados">\n' +
'    <div class="mini-row"><div class="mini-icon-cell"><span class="mini-icon">⊜</span></div>\n' +
'    <div class="mini-text-cell"><div class="label">Accionistas afectados</div>\n' +
'    <div class="value">' + numCambios + ' <span class="frac-secondary">/ ' + totalPos + '</span></div>\n' +
'    <div class="sub">Cambio de participación</div></div></div>\n' +
'  </td>\n' +
'</tr></table>\n' +
'<div class="seccion-titulo">Cuadro Base y Escenario Simulado' +
'<span class="sub">Series A · B · C · Fijas y Variables</span></div>\n' +
'<table class="principal">\n' +
'  <colgroup>\n' +
'    <col class="col-acc">\n' +
'    <col class="col-tipo"><col class="col-tipo"><col class="col-tipo">\n' +
'    <col class="col-tipo"><col class="col-tipo"><col class="col-tipo">\n' +
'    <col class="col-total"><col class="col-pct">\n' +
'    <col class="col-tipo"><col class="col-tipo"><col class="col-tipo">\n' +
'    <col class="col-tipo"><col class="col-tipo"><col class="col-tipo">\n' +
'    <col class="col-total"><col class="col-pct">\n' +
'  </colgroup>\n' +
'  <thead>\n' +
'    <tr class="macro">\n' +
'      <th class="macro-acc" rowspan="3">Accionista</th>\n' +
'      <th class="macro-base sep" colspan="8">◀  Cuadro Base  ▶</th>\n' +
'      <th class="macro-sim sep" colspan="8">◀  Escenario Simulado  ▶</th>\n' +
'    </tr>\n' +
'    <tr class="grupo">\n' +
'      <th class="gr-fija-base sep" colspan="3">Acciones Fijas</th>\n' +
'      <th class="gr-var-base" colspan="3">Acciones Variables</th>\n' +
'      <th class="gr-total-base" rowspan="2">Total</th>\n' +
'      <th class="gr-total-base" rowspan="2">%</th>\n' +
'      <th class="gr-fija-sim sep" colspan="3">Fijas (SIM)</th>\n' +
'      <th class="gr-var-sim" colspan="3">Variables (SIM)</th>\n' +
'      <th class="gr-total-sim" rowspan="2">Total</th>\n' +
'      <th class="gr-total-sim" rowspan="2">%</th>\n' +
'    </tr>\n' +
'    <tr class="cols">\n' +
'      <th class="sep">A</th><th>B</th><th>C</th>\n' +
'      <th>A</th><th>B</th><th>C</th>\n' +
'      <th class="sep">A</th><th>B</th><th>C</th>\n' +
'      <th>A</th><th>B</th><th>C</th>\n' +
'    </tr>\n' +
'  </thead>\n' +
'  <tbody>\n' + filasSup + '\n  </tbody>\n' +
'  <tfoot>\n' +
'    <tr class="total">\n' +
'      <td class="acc-name">Total</td>\n' +
'      <td class="sep">' + formatearNum_(cols.baseFijaA) + '</td>\n' +
'      <td>' + formatearNum_(cols.baseFijaB) + '</td>\n' +
'      <td>' + formatearNum_(cols.baseFijaC) + '</td>\n' +
'      <td>' + formatearNum_(cols.baseVarA)  + '</td>\n' +
'      <td>' + formatearNum_(cols.baseVarB)  + '</td>\n' +
'      <td>' + formatearNum_(cols.baseVarC)  + '</td>\n' +
'      <td>' + formatearNum_(T.base) + '</td>\n' +
'      <td>100.00%</td>\n' +
'      <td class="sep">' + formatearNum_(cols.simFijaA) + '</td>\n' +
'      <td>' + formatearNum_(cols.simFijaB) + '</td>\n' +
'      <td>' + formatearNum_(cols.simFijaC) + '</td>\n' +
'      <td>' + formatearNum_(cols.simVarA)  + '</td>\n' +
'      <td>' + formatearNum_(cols.simVarB)  + '</td>\n' +
'      <td>' + formatearNum_(cols.simVarC)  + '</td>\n' +
'      <td>' + formatearNum_(T.sim) + '</td>\n' +
'      <td>' + coberturaDisplay + '</td>\n' +
'    </tr>\n' +
'  </tfoot>\n' +
'</table>\n' +
'<div class="seccion-titulo">Análisis de Variación y Posición Final' +
'<span class="sub">Resultados de la simulación</span></div>\n' +
'<table class="detalle-variacion">\n' +
'  <colgroup>\n' +
'    <col class="col-acc"><col class="col-num"><col class="col-num">\n' +
'    <col class="col-var-acc"><col class="col-var-pct">\n' +
'    <col class="col-pos"><col class="col-pct-fin"><col class="col-estatus">\n' +
'  </colgroup>\n' +
'  <thead><tr>\n' +
'    <th class="left">Accionista</th>\n' +
'    <th>Total Base</th><th>Total SIM</th>\n' +
'    <th class="var-acc">Variación Accionaria</th>\n' +
'    <th class="var-acc">% Var.</th>\n' +
'    <th class="pos-fin">Posición Final Accionaria</th>\n' +
'    <th class="pos-fin">% Final</th>\n' +
'    <th>Estatus</th>\n' +
'  </tr></thead>\n' +
'  <tbody>\n' + filasInf + '\n  </tbody>\n' +
'  <tfoot>\n' +
'    <tr class="total">\n' +
'      <td class="acc-name">Total General</td>\n' +
'      <td>' + formatearNum_(T.base) + '</td>\n' +
'      <td>' + formatearNum_(T.sim)  + '</td>\n' +
'      <td colspan="2">Movimiento neto: ' + formatearNum_(snapshot.totalAccionesMovidas) + ' acciones</td>\n' +
'      <td colspan="2">Cobertura: ' + coberturaDisplay + '</td>\n' +
'      <td>' + (diferencia === 0 ? 'Balanceada' : (snapshot.modoDisminucion ? 'Disminución' : 'Pendiente')) + '</td>\n' +
'    </tr>\n' +
'  </tfoot>\n' +
'</table>\n' +
'<div class="footer">\n' +
'  <div class="footer-l">Financiera Cualli · S.A.P.I. de C.V., SOFOM E.N.R.</div>\n' +
'  <div class="footer-c">' + (snapshot.modoDisminucion ? '⚖️ Modo: Disminución de Capital' : snapshot.modoDistribucion ? '📤 Modo: Distribución de Acciones' : '') + '</div>\n' +
'  <div class="footer-r">ID: ' + idDoc + ' · Página 1 de 1</div>\n' +
'</div>\n' +
'</body></html>';
}

// ═══════════════════════════════════════════════════════════════════
//  FILAS TABLA SUPERIOR
// ═══════════════════════════════════════════════════════════════════

function construirFilaSup_(d, idx, snapshot) {
  // Especiales: Reporto JCML(idx 8), Reporto ISC(idx 9), Tesorería(idx 10)
  var ESPECIALES_IDX = [8, 9, 10];
  var esEspecial     = ESPECIALES_IDX.indexOf(idx) !== -1;
  var den            = snapshot.totales.denominador;
  var pctBase        = snapshot.totales.base > 0 ? d.totalBase / snapshot.totales.base : 0;
  var pctSim         = den > 0 ? d.totalSim / den : 0;
  var partes         = String(d.nombre || '').split('\n');
  var nombre         = partes[0] || '';
  var rfc            = partes[1] || '';

  return '\n    <tr' + (esEspecial ? ' class="especial"' : '') + '>' +
    '<td class="acc-name">' + escaparHTML_(nombre) +
      (rfc ? '<span class="rfc">' + escaparHTML_(rfc) + '</span>' : '') + '</td>' +
    '<td class="fija-base sep">' + formatearNum_(d.baseFija.A) + '</td>' +
    '<td class="fija-base">'     + formatearNum_(d.baseFija.B) + '</td>' +
    '<td class="fija-base">'     + formatearNum_(d.baseFija.C) + '</td>' +
    '<td class="var-base">'      + formatearNum_(d.baseVar.A)  + '</td>' +
    '<td class="var-base">'      + formatearNum_(d.baseVar.B)  + '</td>' +
    '<td class="var-base">'      + formatearNum_(d.baseVar.C)  + '</td>' +
    '<td class="total-base">'    + formatearNum_(d.totalBase)  + '</td>' +
    '<td class="pct-base">'      + formatearPct_(pctBase)      + '</td>' +
    '<td class="fija-sim sep">'  + formatearNum_(d.simFija.A)  + '</td>' +
    '<td class="fija-sim">'      + formatearNum_(d.simFija.B)  + '</td>' +
    '<td class="fija-sim">'      + formatearNum_(d.simFija.C)  + '</td>' +
    '<td class="var-sim">'       + formatearNum_(d.simVar.A)   + '</td>' +
    '<td class="var-sim">'       + formatearNum_(d.simVar.B)   + '</td>' +
    '<td class="var-sim">'       + formatearNum_(d.simVar.C)   + '</td>' +
    '<td class="total-sim">'     + formatearNum_(d.totalSim)   + '</td>' +
    '<td class="pct-sim">'       + formatearPct_(pctSim)       + '</td>' +
    '</tr>';
}

// ═══════════════════════════════════════════════════════════════════
//  FILAS TABLA INFERIOR
// ═══════════════════════════════════════════════════════════════════

function construirFilaInf_(d, idx, snapshot) {
  var ESPECIALES_IDX = [8, 9, 10];
  var esEspecial     = ESPECIALES_IDX.indexOf(idx) !== -1;
  var den            = snapshot.totales.denominador;
  var pctBase        = snapshot.totales.base > 0 ? d.totalBase / snapshot.totales.base : 0;
  var pctSim         = den > 0 ? d.totalSim / den : 0;
  var deltaAcc = d.totalSim - d.totalBase;
  var deltaPct = pctSim - pctBase;

  // El estatus SIEMPRE se basa en acciones absolutas (deltaAcc),
  // tanto en modo normal como en modo disminución.
  // La tenencia no cambia por efecto del denominador — solo cambia
  // si el accionista realmente cedió o adquirió acciones.
  //
  // Caso especial en modo disminución: fila especial (Reporto/Tesorería)
  // con SIM en cero → "🔻 Capital cancelado" en lugar de "⬇ Disminuye"

  var varAccTxt, varPctTxt, varClass, arrow, arrowClass, badgeClass, badgeContent;

  if (deltaAcc > 0) {
    varAccTxt  = '▲ ' + formatearNum_(Math.abs(deltaAcc));
    varPctTxt  = '▲ ' + formatearPct_(Math.abs(deltaPct));
    varClass   = 'up'; arrow = '▲'; arrowClass = 'arrow-up';
    badgeClass = 'badge-aumenta'; badgeContent = '⬆ Aumenta';

  } else if (deltaAcc < 0) {
    varAccTxt = '▼ ' + formatearNum_(Math.abs(deltaAcc));
    varPctTxt = '▼ ' + formatearPct_(Math.abs(deltaPct));
    varClass  = 'down'; arrow = '▼'; arrowClass = 'arrow-down';

    // Capital cancelado:
    // - Tesorería (idx 10): en modo distribución O disminución
    // - Reporto ISC (idx 9): solo en modo disminución
    var esCancelado =
      (idx === 10 && d.totalSim === 0 && (snapshot.modoDistribucion || snapshot.modoDisminucion)) ||
      (idx === 9  && d.totalSim === 0 && snapshot.modoDisminucion);

    if (esCancelado) {
      badgeClass   = 'badge-cancelado';
      badgeContent = '🔻 Capital cancelado';
    } else {
      badgeClass   = 'badge-disminuye';
      badgeContent = '⬇ Disminuye';
    }

  } else {
    // deltaAcc === 0 → tenencia sin cambio
    varAccTxt = '— 0'; varPctTxt = '— 0.00%';
    varClass  = 'eq'; arrow = '→'; arrowClass = 'arrow-eq';
    badgeClass   = d.totalSim === 0 ? 'badge-sinpart' : 'badge-sincambio';
    badgeContent = d.totalSim === 0 ? 'Sin participación' : '✔ Sin cambio';
  }

  var partes = String(d.nombre || '').split('\n');
  var nombre = partes[0] || '';
  var rfc    = (partes[1] || '').replace(/^RFC:\s*/i, '');
  var tagEsp = esEspecial ? '<span class="tag-especial">ESPECIAL</span>' : '';

  return '\n    <tr' + (esEspecial ? ' class="especial"' : '') + '>' +
    '<td class="acc-name">' + escaparHTML_(nombre) +
      (rfc ? '<span class="rfc">' + escaparHTML_(rfc) + '</span>' : '') + tagEsp + '</td>' +
    '<td class="' + (d.totalBase === 0 ? 'num muted' : 'num') + '">' + formatearNum_(d.totalBase) + '</td>' +
    '<td class="' + (d.totalSim  === 0 ? 'num muted' : 'num') + '">' + formatearNum_(d.totalSim)  + '</td>' +
    '<td class="var-acc-cell ' + varClass + '">' + varAccTxt + '</td>' +
    '<td class="var-pct-cell ' + varClass + '">' + varPctTxt + '</td>' +
    '<td class="pos-final">' + formatearNum_(d.totalBase) +
      ' <span class="' + arrowClass + '">' + arrow + '</span> ' + formatearNum_(d.totalSim) + '</td>' +
    '<td class="pct-final">' + formatearPct_(pctBase) +
      ' <span class="' + arrowClass + '">' + arrow + '</span> ' + formatearPct_(pctSim) + '</td>' +
    '<td class="estatus-cell"><span class="' + badgeClass + '">' + badgeContent + '</span></td>' +
    '</tr>';
}

// ═══════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════

function formatearNum_(n) {
  if (n == null || isNaN(n)) return '0';
  return Math.round(Number(n)).toLocaleString('es-MX');
}

function formatearPct_(n) {
  if (n == null || isNaN(n)) return '0.00%';
  return (Number(n) * 100).toFixed(2) + '%';
}

function escaparHTML_(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
