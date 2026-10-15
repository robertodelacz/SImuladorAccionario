
function buildSimulador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName("SIMULADOR_ACCIONARIO");
  if (existing) ss.deleteSheet(existing);

  var ws = ss.insertSheet("SIMULADOR_ACCIONARIO");
  ss.setActiveSheet(ws);

  // ── PALETA ──────────────────────────────────────────────────────
  var C = {
    // ─── INSTITUCIONALES (aplicados a encabezados y cuadros base/sim) ───
    cualli_amarillo:     "#FDB913",
    amarillo_palido:     "#FFF4D6",
    amarillo_muy_palido: "#FFFBEC",
    amarillo_oscuro:     "#C99100",
    cualli_gris:         "#515151",
    gris_muy_oscuro:     "#2E2E2E",
    gris_oscuro:         "#3F3F3F",
    gris_medio:          "#757575",
    gris_borde:          "#D1D5DB",
    gris_muy_claro:      "#F5F5F5",
    gris_palido:         "#FAFAFA",
    white:               "#FFFFFF",
    black:               "#000000",

    cualli_dark:   "#1F2D3D",
    cualli_yellow: "#FDB913",   
    cualli_gray:   "#515151",   
    blue_mid:      "#2E75B6",
    blue_dark:     "#1A5276",
    blue_light:    "#D6EAF8",
    sim_header:    "#E67E22",
    sim_bg:        "#FFF3CD",
    sim_text:      "#7D6608",
    result_dark:   "#1E8449",
    result_mid:    "#27AE60",
    result_bg:     "#D5F5E3",
    total_bg:      "#2C3E50",
    sep_bg:        "#EAECEE",
    card_green:    "#D5F5E3",
    label_gray:    "#7F8C8D",
    instruc_bg:    "#FFF9E6",
    instruc_text:  "#7D6608",
    var_bg:        "#EBF5FB",
    var_text:      "#1B2631",
    purple:        "#6C3483",
    orange_dark:   "#D35400",
  };

  // ── MAPA DE FILAS ────────────────────────────────────────────────
  var R = {
    title:    2,
    card_top: 4, card_mid: 5, card_bot: 6,
    instruc:  8,
    grp_hdr:  10,
    col_hdr:  11,
    start:    12,
    fundus:   12, ricardo: 13, gsf: 14, inmob: 15,
    sergio:   16, carlos: 17, adolfo: 18, mazzoni: 19,
    reporto:  20, tesor: 21, nuevo: 22, nuevo2: 23, nuevo3: 24,
    total:    25,
    chk_row:  5,
    chk_col:  12,
  };

  // ── ANCHOS DE COLUMNA ────────────────────────────────────────────
  var colWidths = [
    [1,  260], [2,   80], [3,   80], [4,   80], [5,   80],
    [6,   80], [7,   80], [8,   90], [9,   72], [10,  82],
    [11,  82], [12,  82], [13,  82], [14,  82], [15,  82],
    [16,  90], [17,  72], [18, 108], [19,  80], [20, 195],
    [21, 105], [22, 115], [23, 180],
  ];
  colWidths.forEach(function(item) { ws.setColumnWidth(item[0], item[1]); });

  // ── ALTOS DE FILA ────────────────────────────────────────────────
  var rowHeights = [
    [1, 10], [2, 48], [3, 10],
    [4, 34], [5, 34], [6, 34],
    [7, 10], [8, 22], [9, 10],
    [10, 22], [11, 36],
  ];
  for (var r = R.start; r <= R.total; r++) {
    rowHeights.push([r, r === R.total ? 32 : 26]);
  }
  rowHeights.forEach(function(item) { ws.setRowHeight(item[0], item[1]); });

  // ── HELPERS ──────────────────────────────────────────────────────
  function cell(row, col) { return ws.getRange(row, col); }
  function range(r1, c1, r2, c2) {
    return ws.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  function styleRange(rng, opts) {
    if (opts.bg   != null) rng.setBackground(opts.bg);
    if (opts.fg   != null) rng.setFontColor(opts.fg);
    if (opts.bold != null) rng.setFontWeight(opts.bold ? "bold" : "normal");
    if (opts.italic)       rng.setFontStyle("italic");
    if (opts.size != null) rng.setFontSize(opts.size);
    if (opts.h)            rng.setHorizontalAlignment(opts.h);
    if (opts.v)            rng.setVerticalAlignment(opts.v || "middle");
    if (opts.wrap != null) rng.setWrap(opts.wrap);
    if (opts.font)         rng.setFontFamily(opts.font);
  }
  function solidBorder(rng, color, thickness) {
    var style = thickness === "thick"
      ? SpreadsheetApp.BorderStyle.SOLID_THICK
      : thickness === "medium"
      ? SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      : SpreadsheetApp.BorderStyle.SOLID;
    rng.setBorder(true, true, true, true, false, false, color || "#CCCCCC", style);
  }

  // ════════════════════════════════════════════════════════════════
  // 1. TÍTULO  ★ INSTITUCIONAL: gris oscuro sobre amarillo ★
  // ════════════════════════════════════════════════════════════════
  var titleRng = range(R.title, 1, R.title, 23);
  titleRng.merge();
  titleRng.setValue("SIMULADOR ACCIONARIO  ·  FINANCIERA CUALLI");
  styleRange(titleRng, {
    bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true,
    size: 18, h: "center", v: "middle", font: "Montserrat",
  });
  titleRng.setBorder(false, false, true, false, false, false,
    C.gris_muy_oscuro, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ════════════════════════════════════════════════════════════════
  // 2. ÁREA DE CARDS (filas 4-6) — sin cambios respecto al original
  // ════════════════════════════════════════════════════════════════

  // Card reset / checkbox
  range(R.card_top, 1, R.card_top, 13).merge().setValue("🔄  RESTABLECER SIMULACIÓN");
  styleRange(range(R.card_top, 1, R.card_top, 13),
    { bg: C.cualli_dark, fg: C.cualli_yellow, bold: true, size: 10, h: "center" });

  range(R.card_mid, 1, R.card_mid, 11).merge()
    .setValue("Marcar para restablecer todos los valores originales →");
  styleRange(range(R.card_mid, 1, R.card_mid, 11),
    { bg: "#F8F9FA", fg: C.cualli_gray, italic: true, size: 9, h: "right" });

  var chkCell = ws.getRange(R.chk_row, R.chk_col);
  chkCell.insertCheckboxes();
  chkCell.setValue(false);
  styleRange(chkCell, { bg: "#FFF9E6", size: 14, h: "center", v: "middle" });
  chkCell.setBorder(true, true, true, true, false, false,
    C.cualli_yellow, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  chkCell.setNote("Marca aquí para restablecer el escenario simulado a los valores del cuadro base.");

  styleRange(cell(R.card_mid, 13),
    { bg: "#FFF9E6", fg: "#7D6608", bold: true, size: 8, h: "left" });
  cell(R.card_mid, 13).setValue("← toca");

  range(R.card_bot, 1, R.card_bot, 13).merge()
    .setValue("⚠️  Copia los valores del Cuadro Base al Escenario Simulado");
  styleRange(range(R.card_bot, 1, R.card_bot, 13),
    { bg: "#FFF3CD", fg: "#7D6608", italic: true, size: 8, h: "center" });

  solidBorder(range(R.card_top, 1, R.card_bot, 13), C.cualli_yellow, "medium");

  // Card 1: Estado de simulación
  range(R.card_top, 14, R.card_top, 16).merge().setValue("⚡  ESTADO DE SIMULACIÓN");
  styleRange(range(R.card_top, 14, R.card_top, 16),
    { bg: C.blue_mid, fg: C.white, bold: true, size: 9, h: "center" });

  range(R.card_mid, 14, R.card_mid, 16).merge()
    .setFormula('=IF(H' + R.total + '=P' + R.total + ',"✅  SIN DIFERENCIAS","⚠️  REVISAR DIFERENCIA")');
  styleRange(range(R.card_mid, 14, R.card_mid, 16),
    { bg: C.card_green, fg: C.result_dark, bold: true, size: 10, h: "center" });

  cell(R.card_bot, 14).setValue("Diferencia pendiente:");
  styleRange(cell(R.card_bot, 14), { bg: "#F8F9FA", fg: C.label_gray, size: 8, h: "right" });
  range(R.card_bot, 15, R.card_bot, 16).merge()
    .setFormula('=H' + R.total + '-P' + R.total);
  ws.getRange(R.card_bot, 15, 1, 2).setNumberFormat('#,##0;(#,##0);"-"');
  styleRange(range(R.card_bot, 15, R.card_bot, 16),
    { bg: "#F8F9FA", fg: C.result_dark, bold: true, size: 10, h: "center" });

  solidBorder(range(R.card_top, 14, R.card_bot, 16), "#BDC3C7", "medium");

  // Card 2: Acciones
  range(R.card_top, 17, R.card_top, 19).merge().setValue("📊  ACCIONES");
  styleRange(range(R.card_top, 17, R.card_top, 19),
    { bg: C.blue_dark, fg: C.white, bold: true, size: 9, h: "center" });

  cell(R.card_mid, 17).setValue("Total Base:");
  styleRange(cell(R.card_mid, 17), { bg: "#F8F9FA", fg: C.label_gray, size: 8, h: "right" });
  range(R.card_mid, 18, R.card_mid, 19).merge().setFormula('=H' + R.total);
  ws.getRange(R.card_mid, 18, 1, 2).setNumberFormat("#,##0");
  styleRange(range(R.card_mid, 18, R.card_mid, 19),
    { bg: "#F8F9FA", fg: C.black, bold: true, size: 10, h: "center" });

  cell(R.card_bot, 17).setValue("Total Simulado:");
  styleRange(cell(R.card_bot, 17), { bg: "#F8F9FA", fg: C.label_gray, size: 8, h: "right" });
  range(R.card_bot, 18, R.card_bot, 19).merge().setFormula('=P' + R.total);
  ws.getRange(R.card_bot, 18, 1, 2).setNumberFormat("#,##0");
  styleRange(range(R.card_bot, 18, R.card_bot, 19),
    { bg: "#F8F9FA", fg: C.result_dark, bold: true, size: 10, h: "center" });

  solidBorder(range(R.card_top, 17, R.card_bot, 19), "#BDC3C7", "medium");

  // Card 3: Especiales
  range(R.card_top, 20, R.card_top, 23).merge().setValue("📋  ACCIONES ESPECIALES");
  styleRange(range(R.card_top, 20, R.card_top, 23),
    { bg: C.purple, fg: C.white, bold: true, size: 9, h: "center" });

  var especialesCards = [
    [R.card_mid, "En Reporto:",   '=SUM(J' + R.reporto + ':O' + R.reporto + ')'],
    [R.card_bot, "En Tesorería:", '=SUM(J' + R.tesor + ':O' + R.tesor + ')'],
  ];
  especialesCards.forEach(function(item) {
    var row = item[0], label = item[1], formula = item[2];
    cell(row, 20).setValue(label);
    styleRange(cell(row, 20), { bg: "#F8F9FA", fg: C.label_gray, size: 8, h: "right" });
    range(row, 21, row, 22).merge().setFormula(formula);
    ws.getRange(row, 21, 1, 2).setNumberFormat("#,##0");
    styleRange(range(row, 21, row, 22),
      { bg: "#F8F9FA", fg: C.purple, bold: true, size: 10, h: "center" });
    styleRange(cell(row, 23), { bg: "#F8F9FA" });
  });

  solidBorder(range(R.card_top, 20, R.card_bot, 23), "#BDC3C7", "medium");

  // ════════════════════════════════════════════════════════════════
  // 3. INSTRUCCIONES — sin cambios
  // ════════════════════════════════════════════════════════════════
  range(R.instruc, 1, R.instruc, 23).merge()
    .setValue("  ✏️   Modifique únicamente las columnas  SIM A, SIM B, SIM C  (fondo amarillo)   —   El resto del cuadro se recalcula automáticamente.");
  styleRange(range(R.instruc, 1, R.instruc, 23),
    { bg: C.instruc_bg, fg: C.instruc_text, italic: true, size: 9, h: "left" });
  range(R.instruc, 1, R.instruc, 23)
    .setBorder(false, false, false, true, false, false,
      C.cualli_yellow, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ════════════════════════════════════════════════════════════════
  // 4. ENCABEZADOS DE GRUPO  ★ INSTITUCIONAL ★
  //    Gris oscuro sobre amarillo
  // ════════════════════════════════════════════════════════════════
  range(R.grp_hdr, 1, R.grp_hdr, 9).merge()
    .setValue("◀  CUADRO BASE (solo lectura)  ▶");
  styleRange(range(R.grp_hdr, 1, R.grp_hdr, 9),
    { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 9, h: "center" });

  range(R.grp_hdr, 10, R.grp_hdr, 17).merge()
    .setValue("✏  ESCENARIO SIMULADO — edite aquí  ✏");
  styleRange(range(R.grp_hdr, 10, R.grp_hdr, 17),
    { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 9, h: "center" });

  // El tercer grupo (Variación/Posición/Estatus) NO se toca: queda como antes
  range(R.grp_hdr, 18, R.grp_hdr, 23).merge()
    .setValue("📈  ANÁLISIS DE VARIACIÓN Y POSICIÓN FINAL");
  styleRange(range(R.grp_hdr, 18, R.grp_hdr, 23),
    { bg: C.result_dark, fg: C.white, bold: true, size: 9, h: "center" });

  // ════════════════════════════════════════════════════════════════
  // 5. ENCABEZADOS DE COLUMNA   INSTITUCIONAL 
  // ════════════════════════════════════════════════════════════════
  var colHeaders = [
    // Col, label, bg, fg
    [1,  "ACCIONISTA / RFC",          C.gris_muy_oscuro, C.cualli_amarillo],
    [2,  "ACC. A\n(Fija)",            C.cualli_amarillo, C.gris_muy_oscuro],
    [3,  "ACC. B\n(Fija)",            C.cualli_amarillo, C.gris_muy_oscuro],
    [4,  "ACC. C\n(Fija)",            C.cualli_amarillo, C.gris_muy_oscuro],
    [5,  "ACC. A\n(Variable)",        C.cualli_amarillo, C.gris_muy_oscuro],
    [6,  "ACC. B\n(Variable)",        C.cualli_amarillo, C.gris_muy_oscuro],
    [7,  "ACC. C\n(Variable)",        C.cualli_amarillo, C.gris_muy_oscuro],
    [8,  "TOTAL\nBASE",               C.gris_muy_oscuro, C.cualli_amarillo],
    [9,  "% BASE",                    C.gris_muy_oscuro, C.cualli_amarillo],
    [10, "SIM A\n(Fija)",             C.cualli_amarillo, C.gris_muy_oscuro],
    [11, "SIM B\n(Fija)",             C.cualli_amarillo, C.gris_muy_oscuro],
    [12, "SIM C\n(Fija)",             C.cualli_amarillo, C.gris_muy_oscuro],
    [13, "SIM A\n(Variable)",         C.cualli_amarillo, C.gris_muy_oscuro],
    [14, "SIM B\n(Variable)",         C.cualli_amarillo, C.gris_muy_oscuro],
    [15, "SIM C\n(Variable)",         C.cualli_amarillo, C.gris_muy_oscuro],
    [16, "TOTAL\nSIM",                C.gris_muy_oscuro, C.cualli_amarillo],
    [17, "% SIM",                     C.gris_muy_oscuro, C.cualli_amarillo],
    [18, "VARIACIÓN\nACCIONARIA",     C.var_text,     C.white],
    [19, "% VAR.",                    C.var_text,     C.white],
    [20, "POSICIÓN FINAL ACCIONARIA", C.var_text,     C.white],
    [21, "% FINAL",                   C.var_text,     C.white],
    [22, "ESTATUS",                   C.cualli_gris,  C.white],
    [23, "ACTO CORPORATIVO",          C.cualli_gris,  C.white],
  ];
  colHeaders.forEach(function(item) {
    var col = item[0], label = item[1], bg = item[2], fg = item[3];
    var hdr = cell(R.col_hdr, col);
    hdr.setValue(label);
    styleRange(hdr, { bg: bg, fg: fg, bold: true, size: 8, h: "center", v: "middle", wrap: true });
    hdr.setBorder(false, false, true, false, false, false,
      C.white, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // ════════════════════════════════════════════════════════════════
  // 6. FILAS DE DATOS
  // ════════════════════════════════════════════════════════════════
  var accionistas = [
    [R.fundus,  "Fundus Capital S.A.P.I. de C.V.\nRFC: FCA211028892",
      45000,0,2000,6256,0,278,     0,0,2000,6256,0,278],
    [R.ricardo, "Ricardo Madrazo Langle\nRFC: MALR720123GB6",
      0,6965,0,0,968,0,            0,6965,0,0,968,0],
    [R.gsf,     "Grupo GSF, S.A. de C.V.\nRFC: GGS001214JE1",
      0,6965,11870,0,968,1650,     0,6965,11870,0,968,1650],
    [R.inmob,   "Inmobiliaria Sur Centro, S.A.P.I. de C.V.\nRFC: ISC821208DN4",
      0,0,0,0,0,0,                 0,0,0,0,0,0],
    [R.sergio,  "Sergio Madrazo Langle\nRFC: MAL5671022EZ1",
      0,0,7770,0,0,1080,           0,0,7770,0,0,1080],
    [R.carlos,  "Carlos Mauricio Ballesteros Franco\nRFC: BAFC480629I91",
      0,0,3000,0,0,1366,           0,0,3000,0,0,1366],
    [R.adolfo,  "Adolfo González Olhovich\nRFC: GOOA730326K31",
      0,0,2500,0,0,348,            0,0,2500,0,0,348],
    [R.mazzoni, "Juan Carlos Mazzoni López\nRFC: MALJ700220MR0",
      0,0,0,0,0,1947,              0,0,0,0,0,1947],
  ];
  var especiales = [
    [R.reporto, "Acciones en Reporto JCML",  0,0,0,0,0,9443, 0,13930,0,0,0,9443],
    [R.reporto, "Acciones en Reporto ISC",  0,3242,0,0,0,9443, 0,3242,0,0,0,0],
    [R.tesor,   "Tesorería",            0,0,0,0,0,0,           0,0,0,0,0,0],
    [R.nuevo,   "Accionista Nuevo 1",     0,0,0,0,0,0,           0,0,0,0,0,0],
    [R.nuevo2,  "Accionista Nuevo 2",     0,0,0,0,0,0,           0,0,0,0,0,0],
    [R.nuevo3,  "Accionista Nuevo 3",     0,0,0,0,0,0,           0,0,0,0,0,0],
  ];

  // ★ Fondos institucionales para filas alternas: blanco + gris muy claro 
  function writeDataRow(row, name, b, c, d, e, f, g, j, k, l, m, n, o, isSpecial) {
    var bgBase = isSpecial
      ? C.gris_muy_claro  // filas especiales: gris claro (institucional)
      : (row % 2 === 0 ? C.white : C.gris_palido);  // alternancia blanco / gris muy claro
    var fmtNum = "#,##0";
    var fmtPct = "0.00%";

    // Columna A: nombre / RFC   institucional 
    var nc = cell(row, 1);
    nc.setValue(name);
    styleRange(nc, {
      bg: isSpecial ? C.gris_muy_oscuro : C.gris_palido,
      fg: isSpecial ? C.white : C.gris_muy_oscuro,
      bold: isSpecial, size: 9, h: "left", v: "middle", wrap: true,
    });
    nc.setBorder(false, true, true, false, false, false,
      C.gris_muy_oscuro, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Cols B-G: Cuadro Base  ★ institucional 
    [b,c,d,e,f,g].forEach(function(val, i) {
      var cr = cell(row, 2 + i);
      cr.setValue(val).setNumberFormat(fmtNum);
      styleRange(cr, { bg: bgBase, fg: C.gris_muy_oscuro, size: 9, h: "center" });
      cr.setBorder(true, false, true, false, false, false,
        C.gris_borde, SpreadsheetApp.BorderStyle.SOLID);
    });

    // Col H: Total Base (sin cambios — celeste como antes)
    cell(row, 8).setFormula('=SUM(B' + row + ':G' + row + ')').setNumberFormat(fmtNum);
    styleRange(cell(row, 8), { bg: C.blue_light, fg: C.blue_dark, bold: true, size: 9, h: "center" });
    cell(row, 8).setBorder(false, true, true, true, false, false,
      "#1A5276", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Col I: % Base (sin cambios)
    cell(row, 9).setFormula(
      '=IFERROR(H' + row + '/SUM($H$' + R.start + ':$H$' + R.nuevo + '),0)'
    ).setNumberFormat(fmtPct);
    styleRange(cell(row, 9), { bg: C.blue_light, fg: C.blue_dark, bold: true, size: 9, h: "center" });
    cell(row, 9).setBorder(false, true, true, false, false, false,
      C.sim_header, SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Cols J-O: Cuadro Simulado (editable)   institucional: amarillo pálido 
    [j,k,l,m,n,o].forEach(function(val, i) {
      var cr = cell(row, 10 + i);
      cr.setValue(val).setNumberFormat(fmtNum);
      styleRange(cr, {
        bg: isSpecial ? C.amarillo_muy_palido : C.amarillo_palido,
        fg: C.amarillo_oscuro,
        bold: false, size: 9, h: "center"
      });
      cr.setBorder(true, false, true, false, false, false,
        C.cualli_amarillo, SpreadsheetApp.BorderStyle.SOLID);
    });

    // Col P: Total SIM (sin cambios — verde)
    cell(row, 16).setFormula('=SUM(J' + row + ':O' + row + ')').setNumberFormat(fmtNum);
    styleRange(cell(row, 16), { bg: C.card_green, fg: C.result_dark, bold: true, size: 9, h: "center" });
    cell(row, 16).setBorder(false, true, true, true, false, false,
      C.result_dark, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Col Q: % SIM (sin cambios)
    cell(row, 17).setFormula(
      '=IFERROR(P' + row + '/SUM($H$' + R.start + ':$H$' + R.nuevo + '),0)'
    ).setNumberFormat(fmtPct);
    styleRange(cell(row, 17), { bg: C.card_green, fg: C.result_dark, bold: true, size: 9, h: "center" });
    cell(row, 17).setBorder(false, true, true, false, false, false,
      C.var_text, SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Col R: Variación Accionaria (sin cambios)
    cell(row, 18).setFormula(
      '=IF(P' + row + '-H' + row + '>0,"▲ "&TEXT(P' + row + '-H' + row + ',"#,##0"),' +
      'IF(P' + row + '-H' + row + '<0,"▼ "&TEXT(ABS(P' + row + '-H' + row + '),"#,##0"),"— 0"))'
    );
    styleRange(cell(row, 18), { bg: C.var_bg, fg: C.var_text, bold: true, size: 9, h: "center" });

    // Col S: % Variación (sin cambios)
    cell(row, 19).setFormula(
      '=IF(Q' + row + '-I' + row + '>0,"▲ "&TEXT(Q' + row + '-I' + row + ',"0.00%"),' +
      'IF(Q' + row + '-I' + row + '<0,"▼ "&TEXT(ABS(Q' + row + '-I' + row + '),"0.00%"),"— 0.00%"))'
    );
    styleRange(cell(row, 19), { bg: C.var_bg, fg: C.var_text, bold: true, size: 9, h: "center" });

    // Col T: Posición Final (sin cambios)
    cell(row, 20).setFormula(
      '=TEXT(H' + row + ',"#,##0")&" → "&TEXT(P' + row + ',"#,##0")'
    );
    styleRange(cell(row, 20), { bg: C.var_bg, fg: C.var_text, bold: true, size: 9, h: "center" });

    // Col U: % Final (sin cambios)
    cell(row, 21).setFormula(
      '=TEXT(I' + row + ',"0.00%")&IF(Q' + row + '>I' + row + '," ▲ ",IF(Q' + row + '<I' + row + '," ▼ "," → "))&TEXT(Q' + row + ',"0.00%")'
    );
    styleRange(cell(row, 21), { bg: C.var_bg, fg: C.var_text, bold: true, size: 9, h: "center" });

    // Col V: Estatus (sin cambios)
    cell(row, 22).setFormula(
      '=IF(P' + row + '=0,"Sin participación",' +
      'IF(P' + row + '-H' + row + '>0,"⬆ Aumenta",' +
      'IF(P' + row + '-H' + row + '<0,"⬇ Disminuye","✔ Sin cambio")))'
    );
    styleRange(cell(row, 22), { bg: "#F4F6F7", fg: C.cualli_gray, size: 9, h: "center" });

    // Col W: Acto Corporativo (sin cambios)
    styleRange(cell(row, 23), { bg: "#F9F9F9", fg: C.cualli_gray, size: 9, h: "left" });
    cell(row, 23).setBorder(false, true, true, false, false, false,
      "#BBBBBB", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  accionistas.forEach(function(d) {
    writeDataRow(d[0],d[1],d[2],d[3],d[4],d[5],d[6],d[7],d[8],d[9],d[10],d[11],d[12],d[13],false);
  });
  especiales.forEach(function(d) {
    writeDataRow(d[0],d[1],d[2],d[3],d[4],d[5],d[6],d[7],d[8],d[9],d[10],d[11],d[12],d[13],true);
  });

  // ════════════════════════════════════════════════════════════════
  // 7. FILA TOTAL  ★ INSTITUCIONAL ★
  // ════════════════════════════════════════════════════════════════
  var tr = R.total;
  styleRange(range(tr, 1, tr, 17),
    { bg: C.gris_muy_oscuro, fg: C.cualli_amarillo, bold: true, size: 10 });
  range(tr, 1, tr, 17).setBorder(true, true, true, true, false, false,
    C.cualli_amarillo, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Cols 18-23: variación, sin cambios
  styleRange(range(tr, 18, tr, 23),
    { bg: C.cualli_dark, fg: C.cualli_yellow, bold: true, size: 10 });
  range(tr, 18, tr, 23).setBorder(true, true, true, true, false, false,
    C.cualli_yellow, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  cell(tr, 1).setValue("TOTAL GENERAL");
  styleRange(cell(tr, 1), { h: "center", size: 11 });

  ["B","C","D","E","F","G"].forEach(function(col, i) {
    cell(tr, 2 + i).setFormula('=SUM(' + col + R.start + ':' + col + R.reporto + ')').setNumberFormat("#,##0");
    styleRange(cell(tr, 2 + i), { h: "center", size: 9 });
  });

  cell(tr, 8).setFormula('=SUM(H' + R.start + ':H' + R.nuevo + ')').setNumberFormat("#,##0");
  styleRange(cell(tr, 8), { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 11, h: "center" });

  cell(tr, 9).setFormula('=SUM(I' + R.start + ':I' + R.nuevo + ')').setNumberFormat("0.00%");
  styleRange(cell(tr, 9), { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 10, h: "center" });

  ["J","K","L","M","N","O"].forEach(function(col, i) {
    cell(tr, 10 + i).setFormula('=SUM(' + col + R.start + ':' + col + R.reporto + ')').setNumberFormat("#,##0");
    styleRange(cell(tr, 10 + i),
      { bg: C.gris_muy_oscuro, fg: C.cualli_amarillo, bold: true, size: 9, h: "center" });
  });

  cell(tr, 16).setFormula('=SUM(P' + R.start + ':P' + R.nuevo + ')').setNumberFormat("#,##0");
  styleRange(cell(tr, 16),
    { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 11, h: "center" });

  cell(tr, 17).setFormula(
    '=IFERROR(SUM(P' + R.start + ':P' + R.nuevo + ')/SUM(H' + R.start + ':H' + R.nuevo + '),0)'
  ).setNumberFormat("0.00%");
  styleRange(cell(tr, 17),
    { bg: C.cualli_amarillo, fg: C.gris_muy_oscuro, bold: true, size: 10, h: "center" });

  // ════════════════════════════════════════════════════════════════
  // 8. VALIDACIÓN DE DATOS en SIM J:O
  // ════════════════════════════════════════════════════════════════
  range(R.start, 10, R.nuevo, 15).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setHelpText("Ingrese el número de acciones simuladas (entero ≥ 0).")
      .build()
  );

  // ════════════════════════════════════════════════════════════════
  // 9. FORMATO CONDICIONAL 
  // ════════════════════════════════════════════════════════════════
  var cfRules = [];
  var startRow = R.start;
  var endRow   = R.nuevo;
  var numRows  = endRow - startRow + 1;

  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint("#FFFFFF").setGradientMaxpoint("#1A5276")
    .setRanges([ws.getRange(startRow, 8, numRows, 1)]).build());

  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint("#FFFFFF").setGradientMaxpoint("#2E75B6")
    .setRanges([ws.getRange(startRow, 9, numRows, 1)]).build());

  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint("#FFFFFF").setGradientMaxpoint("#1E8449")
    .setRanges([ws.getRange(startRow, 16, numRows, 1)]).build());

  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint("#FFFFFF").setGradientMaxpoint("#27AE60")
    .setRanges([ws.getRange(startRow, 17, numRows, 1)]).build());

  var simComparisons = [
    { simCol: 10, baseCol: "B" }, { simCol: 11, baseCol: "C" },
    { simCol: 12, baseCol: "D" }, { simCol: 13, baseCol: "E" },
    { simCol: 14, baseCol: "F" }, { simCol: 15, baseCol: "G" },
  ];
  simComparisons.forEach(function(item) {
    var simCol = item.simCol, baseCol = item.baseCol;
    var simLetter = columnToLetter(simCol);
    for (var row = startRow; row <= endRow; row++) {
      var rng = ws.getRange(row, simCol, 1, 1);
      cfRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=' + simLetter + row + '>' + baseCol + row)
          .setBackground("#C8F7C5").setFontColor("#1E8449")
          .setRanges([rng]).build()
      );
      cfRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(' + simLetter + row + '<' + baseCol + row + ',' + baseCol + row + '>0)')
          .setBackground("#FADBD8").setFontColor("#C0392B")
          .setRanges([rng]).build()
      );
      cfRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(' + simLetter + row + '=' + baseCol + row + ',' + baseCol + row + '>0)')
          .setBackground("#D6EAF8").setFontColor("#1A5276")
          .setRanges([rng]).build()
      );
    }
  });

  var rngR = ws.getRange(startRow, 18, numRows, 1);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("▲",R' + startRow + '))')
      .setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true)
      .setRanges([rngR]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("▼",R' + startRow + '))')
      .setBackground("#FADBD8").setFontColor("#C0392B").setBold(true)
      .setRanges([rngR]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("—",R' + startRow + '))')
      .setBackground("#F2F3F4").setFontColor("#7F8C8D")
      .setRanges([rngR]).build()
  );

  var rngS = ws.getRange(startRow, 19, numRows, 1);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("▲",S' + startRow + '))')
      .setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true)
      .setRanges([rngS]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("▼",S' + startRow + '))')
      .setBackground("#FADBD8").setFontColor("#C0392B").setBold(true)
      .setRanges([rngS]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISNUMBER(SEARCH("—",S' + startRow + '))')
      .setBackground("#F2F3F4").setFontColor("#7F8C8D")
      .setRanges([rngS]).build()
  );

  var rngV = ws.getRange(startRow, 22, numRows, 1);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Aumenta")
      .setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true)
      .setRanges([rngV]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Disminuye")
      .setBackground("#FADBD8").setFontColor("#C0392B").setBold(true)
      .setRanges([rngV]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Sin cambio")
      .setBackground("#EBF5FB").setFontColor("#2E75B6")
      .setRanges([rngV]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Sin participación")
      .setBackground("#F2F3F4").setFontColor("#AAB7B8")
      .setRanges([rngV]).build()
  );

  var rngCard = ws.getRange(R.card_mid, 14, 1, 3);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("SIN DIFERENCIAS")
      .setBackground("#A9DFBF").setFontColor("#1E8449").setBold(true)
      .setRanges([rngCard]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("REVISAR")
      .setBackground("#F1948A").setFontColor("#922B21").setBold(true)
      .setRanges([rngCard]).build()
  );

  ws.setConditionalFormatRules(cfRules);

  // ════════════════════════════════════════════════════════════════
  // 10. INMOVILIZAR / TAB / OCULTAR FILAS DE PADDING
  // ════════════════════════════════════════════════════════════════
  ws.setFrozenRows(R.col_hdr);
  ws.setTabColor(C.cualli_amarillo);
  ws.hideRows(1);
  ws.hideRows(3);
  ws.hideRows(7);
  ws.hideRows(9);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "✅ Simulador construido exitosamente.\n\n" +
    "• Edite las columnas SIM A / B / C (fondo amarillo pálido).\n" +
    "• El checkbox en L5 restablece los valores al cuadro base.\n" +
    "• Asigne 'menuExportarPDF' a su imagen para generar el PDF."
  );
}

// ── HELPER ───────────────────────────────────────────────────────────
function columnToLetter(col) {
  var letter = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

