const ExcelJS = require("exceljs");
const path = require("path");

async function generateBudgetTracker() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Budget Tracker Pro";
  workbook.created = new Date();

  // ============================================================
  // LUXURY GREEN ANALYTICS COLOR PALETTE (YT Analytics inspired)
  // ============================================================
  const c = {
    // Dark backgrounds
    bgDark: "0D1117",       // Darkest - main bg
    bgCard: "161B22",       // Card background
    bgCardHover: "1C2333",  // Slightly lighter
    bgElevated: "21262D",   // Elevated surfaces
    bgInput: "0D1117",      // Input fields

    // Accent greens (YT Analytics style)
    green1: "00C853",       // Primary bright green
    green2: "00E676",       // Lighter green
    green3: "69F0AE",       // Soft green
    green4: "B9F6CA",       // Very light green
    greenDark: "003D1A",    // Dark green bg
    greenMid: "0A4D2E",     // Medium green bg
    greenGlow: "00E676",    // Glow effect green

    // Accent colors
    blue: "58A6FF",         // Links / info
    blueGlow: "1F6FEB",     // Blue accent
    purple: "BC8CFF",       // Purple accent
    purpleBg: "2D1B4E",     // Purple background
    cyan: "56D4DD",         // Cyan accent
    cyanBg: "0A3D3D",       // Cyan background
    orange: "F0883E",       // Warning
    orangeBg: "3D2200",     // Warning bg
    red: "F85149",          // Danger/expense
    redBg: "3D0C0C",        // Red bg
    redSoft: "FF7B72",      // Softer red
    yellow: "E3B341",       // Gold accent
    yellowBg: "3D2E00",     // Yellow bg
    pink: "F778BA",         // Pink accent

    // Text
    textBright: "F0F6FC",   // Primary text (almost white)
    textPrimary: "C9D1D9",  // Normal text
    textSecondary: "8B949E", // Muted text
    textTertiary: "484F58",  // Very muted
    textGreen: "3FB950",     // Green text

    // Borders
    border: "30363D",       // Default border
    borderLight: "21262D",  // Subtle border
    borderGreen: "238636",  // Green border

    white: "FFFFFF",
    black: "000000",
  };

  // ============================================================
  // SHARED STYLES
  // ============================================================
  const currencyFmt = '#,##0.00 "€"';
  const pctFmt = "0.0%";
  const dateFmt = "DD.MM.YYYY";

  const darkBorder = {
    top:    { style: "thin", color: { argb: c.border } },
    left:   { style: "thin", color: { argb: c.border } },
    bottom: { style: "thin", color: { argb: c.border } },
    right:  { style: "thin", color: { argb: c.border } },
  };

  const noBorder = {
    top: { style: "thin", color: { argb: c.bgDark } },
    left: { style: "thin", color: { argb: c.bgDark } },
    bottom: { style: "thin", color: { argb: c.bgDark } },
    right: { style: "thin", color: { argb: c.bgDark } },
  };

  const greenBorderBottom = {
    top: { style: "thin", color: { argb: c.border } },
    left: { style: "thin", color: { argb: c.border } },
    bottom: { style: "medium", color: { argb: c.green1 } },
    right: { style: "thin", color: { argb: c.border } },
  };

  // Helper: fill entire row with dark background
  function fillRowDark(ws, rowNum, cols, bg = c.bgDark) {
    const row = ws.getRow(rowNum);
    for (let i = 1; i <= cols; i++) {
      row.getCell(i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      row.getCell(i).border = noBorder;
    }
  }

  // Helper: style header row (green glow header)
  function luxuryHeader(ws, rowNum, headers, colCount) {
    const row = ws.getRow(rowNum);
    row.height = 36;
    for (let i = 1; i <= colCount; i++) {
      const cell = row.getCell(i);
      cell.value = headers[i - 1] || "";
      cell.font = { bold: true, size: 10, color: { argb: c.green1 }, name: "Segoe UI" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgCard } };
      cell.border = greenBorderBottom;
      cell.alignment = { vertical: "middle", horizontal: "left" };
    }
  }

  // Helper: style data cell
  function luxuryCell(cell, opts = {}) {
    const { isAlt, font, numFmt, align, fg } = opts;
    const bg = fg || (isAlt ? c.bgCardHover : c.bgCard);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    cell.font = font || { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" };
    cell.border = darkBorder;
    cell.alignment = { vertical: "middle", horizontal: align || "left", ...opts.extra };
    if (numFmt) cell.numFmt = numFmt;
  }

  // Helper: stat card (merged cells with value)
  function createStatCard(ws, startRow, startCol, endCol, label, formula, numFmt, labelColor, valueColor, bgColor) {
    // Label row
    ws.mergeCells(startRow, startCol, startRow, endCol);
    const labelCell = ws.getCell(startRow, startCol);
    labelCell.value = label;
    labelCell.font = { bold: true, size: 9, color: { argb: labelColor }, name: "Segoe UI", letterSpacing: 200 };
    labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
    labelCell.alignment = { horizontal: "center", vertical: "middle" };
    labelCell.border = { top: { style: "medium", color: { argb: labelColor } }, left: { style: "thin", color: { argb: c.border } }, right: { style: "thin", color: { argb: c.border } }, bottom: { style: "thin", color: { argb: c.border } } };
    for (let col = startCol + 1; col <= endCol; col++) {
      ws.getCell(startRow, col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
      ws.getCell(startRow, col).border = labelCell.border;
    }

    // Value row
    ws.mergeCells(startRow + 1, startCol, startRow + 1, endCol);
    const valCell = ws.getCell(startRow + 1, startCol);
    valCell.value = formula;
    valCell.numFmt = numFmt;
    valCell.font = { bold: true, size: 24, color: { argb: valueColor }, name: "Segoe UI" };
    valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
    valCell.alignment = { horizontal: "center", vertical: "middle" };
    valCell.border = { top: { style: "thin", color: { argb: c.border } }, left: { style: "thin", color: { argb: c.border } }, right: { style: "thin", color: { argb: c.border } }, bottom: { style: "medium", color: { argb: labelColor } } };
    for (let col = startCol + 1; col <= endCol; col++) {
      ws.getCell(startRow + 1, col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
      ws.getCell(startRow + 1, col).border = valCell.border;
    }

    // Subtext row
    ws.mergeCells(startRow + 2, startCol, startRow + 2, endCol);
    const subCell = ws.getCell(startRow + 2, startCol);
    subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
    subCell.border = { top: { style: "thin", color: { argb: c.border } }, left: { style: "thin", color: { argb: c.border } }, right: { style: "thin", color: { argb: c.border } }, bottom: { style: "thin", color: { argb: c.border } } };
    for (let col = startCol + 1; col <= endCol; col++) {
      ws.getCell(startRow + 2, col).fill = subCell.fill;
      ws.getCell(startRow + 2, col).border = subCell.border;
    }
  }

  // ============================================================
  // 1. DASHBOARD
  // ============================================================
  const dash = workbook.addWorksheet("Dashboard", {
    properties: { tabColor: { argb: c.green1 } },
  });
  dash.properties.defaultRowHeight = 22;

  // Fill entire background dark
  for (let r = 1; r <= 55; r++) fillRowDark(dash, r, 12);

  // Column widths
  [3, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 3].forEach((w, i) => {
    dash.getColumn(i + 1).width = w;
  });

  // ---- TITLE ----
  dash.mergeCells("B2:K2");
  const title = dash.getCell("B2");
  title.value = "BUDGET TRACKER";
  title.font = { bold: true, size: 28, color: { argb: c.green1 }, name: "Segoe UI" };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  title.alignment = { horizontal: "left", vertical: "middle" };
  dash.getRow(2).height = 50;

  dash.mergeCells("B3:K3");
  const subtitle = dash.getCell("B3");
  subtitle.value = "Financial Analytics Dashboard";
  subtitle.font = { size: 12, color: { argb: c.textSecondary }, name: "Segoe UI" };
  subtitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getRow(3).height = 25;

  // Green accent line under title
  dash.mergeCells("B4:K4");
  dash.getCell("B4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getCell("B4").border = { bottom: { style: "medium", color: { argb: c.green1 } }, top: noBorder.top, left: noBorder.left, right: noBorder.right };
  dash.getRow(4).height = 4;

  // ---- STAT CARDS (Row 6-8) ----
  dash.getRow(6).height = 22;
  dash.getRow(7).height = 42;
  dash.getRow(8).height = 18;

  // Card 1: Gesamtguthaben
  createStatCard(dash, 6, 2, 3,
    "GESAMTGUTHABEN",
    { formula: "Konten!F15" },
    currencyFmt, c.green1, c.green2, c.greenDark
  );

  // Card 2: Einnahmen
  createStatCard(dash, 6, 4, 5,
    "EINNAHMEN · MONAT",
    { formula: 'SUMPRODUCT((Transaktionen!C3:C500="Einnahme")*(MONTH(Transaktionen!B3:B500)=MONTH(TODAY()))*(YEAR(Transaktionen!B3:B500)=YEAR(TODAY()))*Transaktionen!E3:E500)' },
    currencyFmt, c.cyan, c.cyan, c.cyanBg
  );

  // Card 3: Ausgaben
  createStatCard(dash, 6, 6, 7,
    "AUSGABEN · MONAT",
    { formula: 'SUMPRODUCT((Transaktionen!C3:C500="Ausgabe")*(MONTH(Transaktionen!B3:B500)=MONTH(TODAY()))*(YEAR(Transaktionen!B3:B500)=YEAR(TODAY()))*Transaktionen!E3:E500)' },
    currencyFmt, c.red, c.redSoft, c.redBg
  );

  // Card 4: Bilanz
  createStatCard(dash, 6, 8, 9,
    "BILANZ · MONAT",
    { formula: "D7-F7" },
    currencyFmt, c.purple, c.purple, c.purpleBg
  );

  // Card 5: Sparquote
  createStatCard(dash, 6, 10, 11,
    "SPARQUOTE · MONAT",
    { formula: "IF(D7=0,0,(D7-F7)/D7)" },
    pctFmt, c.yellow, c.yellow, c.yellowBg
  );

  // ---- SECTION: MONATSÜBERSICHT (Row 10+) ----
  dash.mergeCells("B10:F10");
  dash.getCell("B10").value = "MONATSÜBERSICHT";
  dash.getCell("B10").font = { bold: true, size: 14, color: { argb: c.textBright }, name: "Segoe UI" };
  dash.getCell("B10").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getRow(10).height = 32;

  // Chart data hint
  dash.mergeCells("B11:F11");
  dash.getCell("B11").value = "Einnahmen vs. Ausgaben pro Monat";
  dash.getCell("B11").font = { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" };
  dash.getCell("B11").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  const months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];

  luxuryHeader(dash, 12, ["", "Monat", "Einnahmen", "Ausgaben", "Bilanz", "Sparquote"], 6);

  for (let i = 0; i < 12; i++) {
    const row = dash.getRow(13 + i);
    const r = 13 + i;
    const isAlt = i % 2 === 1;

    // Row num
    luxuryCell(row.getCell(1), { isAlt, fg: c.bgDark });

    // Month
    luxuryCell(row.getCell(2), {
      isAlt,
      font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" },
    });
    row.getCell(2).value = months[i];

    // Einnahmen
    luxuryCell(row.getCell(3), {
      isAlt,
      numFmt: currencyFmt,
      font: { size: 11, color: { argb: c.green3 }, name: "Segoe UI" },
    });
    row.getCell(3).value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$500="Einnahme")*(MONTH(Transaktionen!B$3:B$500)=${i + 1})*(YEAR(Transaktionen!B$3:B$500)=YEAR(TODAY()))*Transaktionen!E$3:E$500)` };

    // Ausgaben
    luxuryCell(row.getCell(4), {
      isAlt,
      numFmt: currencyFmt,
      font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" },
    });
    row.getCell(4).value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$500="Ausgabe")*(MONTH(Transaktionen!B$3:B$500)=${i + 1})*(YEAR(Transaktionen!B$3:B$500)=YEAR(TODAY()))*Transaktionen!E$3:E$500)` };

    // Bilanz
    luxuryCell(row.getCell(5), {
      isAlt,
      numFmt: currencyFmt,
      font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" },
    });
    row.getCell(5).value = { formula: `C${r}-D${r}` };

    // Sparquote
    luxuryCell(row.getCell(6), {
      isAlt,
      numFmt: pctFmt,
      font: { size: 11, color: { argb: c.yellow }, name: "Segoe UI" },
    });
    row.getCell(6).value = { formula: `IF(C${r}=0,0,E${r}/C${r})` };
  }

  // Total row
  const totRow = dash.getRow(25);
  totRow.height = 32;
  for (let col = 1; col <= 6; col++) {
    const cell = totRow.getCell(col);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.greenDark } };
    cell.border = { top: { style: "medium", color: { argb: c.green1 } }, bottom: { style: "medium", color: { argb: c.green1 } }, left: darkBorder.left, right: darkBorder.right };
    cell.font = { bold: true, size: 12, color: { argb: c.green1 }, name: "Segoe UI" };
    cell.alignment = { vertical: "middle" };
  }
  totRow.getCell(2).value = "GESAMT";
  totRow.getCell(3).value = { formula: "SUM(C13:C24)" };
  totRow.getCell(3).numFmt = currencyFmt;
  totRow.getCell(4).value = { formula: "SUM(D13:D24)" };
  totRow.getCell(4).numFmt = currencyFmt;
  totRow.getCell(5).value = { formula: "C25-D25" };
  totRow.getCell(5).numFmt = currencyFmt;
  totRow.getCell(6).value = { formula: "IF(C25=0,0,E25/C25)" };
  totRow.getCell(6).numFmt = pctFmt;

  // ---- SECTION: BUDGET STATUS (Row 10, Cols H-K) ----
  dash.mergeCells("H10:K10");
  dash.getCell("H10").value = "BUDGET STATUS";
  dash.getCell("H10").font = { bold: true, size: 14, color: { argb: c.textBright }, name: "Segoe UI" };
  dash.getCell("H10").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  dash.mergeCells("H11:K11");
  dash.getCell("H11").value = "Ausgaben vs. Budget (aktueller Monat)";
  dash.getCell("H11").font = { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" };
  dash.getCell("H11").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  luxuryHeader(dash, 12, [null, null, null, null, null, null, null, "Kategorie", "Budget", "Ausgegeben", "Status"], 11);

  const expCats = ["Lebensmittel", "Transport", "Wohnung", "Unterhaltung", "Gesundheit", "Bildung", "Kleidung", "Restaurant"];
  for (let i = 0; i < 8; i++) {
    const row = dash.getRow(13 + i);
    const r = 13 + i;
    const isAlt = i % 2 === 1;

    // Kategorie
    luxuryCell(row.getCell(8), {
      isAlt,
      font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" },
    });
    row.getCell(8).value = { formula: `IF(Budgets!A${3 + i}="","",Budgets!A${3 + i})` };

    // Budget
    luxuryCell(row.getCell(9), {
      isAlt,
      numFmt: currencyFmt,
      font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" },
    });
    row.getCell(9).value = { formula: `IF(H${r}="","",Budgets!B${3 + i})` };

    // Ausgegeben
    luxuryCell(row.getCell(10), {
      isAlt,
      numFmt: currencyFmt,
      font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" },
    });
    row.getCell(10).value = { formula: `IF(H${r}="","",SUMPRODUCT((Transaktionen!C$3:C$500="Ausgabe")*(Transaktionen!D$3:D$500=H${r})*(MONTH(Transaktionen!B$3:B$500)=MONTH(TODAY()))*(YEAR(Transaktionen!B$3:B$500)=YEAR(TODAY()))*Transaktionen!E$3:E$500))` };

    // Status (visual bar using repeat character)
    luxuryCell(row.getCell(11), {
      isAlt,
      font: { bold: true, size: 10, color: { argb: c.green1 }, name: "Segoe UI" },
      align: "center",
    });
    row.getCell(11).value = { formula: `IF(OR(H${r}="",I${r}=0),"",IF(J${r}/I${r}>1,"OVER",IF(J${r}/I${r}>0.8,"WARN",TEXT(J${r}/I${r},"0%"))))` };
  }

  // Budget total
  const budTotDash = dash.getRow(21);
  for (let col = 8; col <= 11; col++) {
    budTotDash.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.greenDark } };
    budTotDash.getCell(col).border = { top: { style: "medium", color: { argb: c.green1 } }, bottom: { style: "medium", color: { argb: c.green1 } }, left: darkBorder.left, right: darkBorder.right };
    budTotDash.getCell(col).font = { bold: true, size: 11, color: { argb: c.green1 }, name: "Segoe UI" };
    budTotDash.getCell(col).alignment = { vertical: "middle" };
  }
  budTotDash.getCell(8).value = "TOTAL";
  budTotDash.getCell(9).value = { formula: "SUM(I13:I20)" };
  budTotDash.getCell(9).numFmt = currencyFmt;
  budTotDash.getCell(10).value = { formula: "SUM(J13:J20)" };
  budTotDash.getCell(10).numFmt = currencyFmt;
  budTotDash.getCell(11).value = { formula: 'IF(I21=0,"",TEXT(J21/I21,"0%"))' };

  // Conditional formatting for budget status
  dash.addConditionalFormatting({
    ref: "K13:K20",
    rules: [
      { type: "containsText", operator: "containsText", text: "OVER",
        style: { font: { bold: true, color: { argb: c.red } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.redBg } } } },
      { type: "containsText", operator: "containsText", text: "WARN",
        style: { font: { bold: true, color: { argb: c.orange } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.orangeBg } } } },
    ],
  });

  // ---- SECTION: KONTEN OVERVIEW (Row 27+) ----
  dash.mergeCells("B27:K27");
  dash.getCell("B27").value = "KONTENÜBERSICHT";
  dash.getCell("B27").font = { bold: true, size: 14, color: { argb: c.textBright }, name: "Segoe UI" };
  dash.getCell("B27").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getRow(27).height = 32;

  dash.mergeCells("B28:K28");
  dash.getCell("B28").value = "Aktueller Stand aller Konten";
  dash.getCell("B28").font = { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" };
  dash.getCell("B28").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  luxuryHeader(dash, 29, ["", "Konto", "Typ", "Startguthaben", "Einnahmen", "Ausgaben", "Aktueller Stand", "", "", "", ""], 11);

  for (let i = 0; i < 4; i++) {
    const row = dash.getRow(30 + i);
    const r = 30 + i;
    const isAlt = i % 2 === 1;
    const kr = 3 + i; // Konten sheet row

    luxuryCell(row.getCell(1), { fg: c.bgDark });
    luxuryCell(row.getCell(2), { isAlt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = { formula: `IF(Konten!A${kr}="","",Konten!A${kr})` };

    luxuryCell(row.getCell(3), { isAlt, font: { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" } });
    row.getCell(3).value = { formula: `IF(B${r}="","",Konten!B${kr})` };

    luxuryCell(row.getCell(4), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" } });
    row.getCell(4).value = { formula: `IF(B${r}="","",Konten!C${kr})` };

    luxuryCell(row.getCell(5), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.green3 }, name: "Segoe UI" } });
    row.getCell(5).value = { formula: `IF(B${r}="","",Konten!D${kr})` };

    luxuryCell(row.getCell(6), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" } });
    row.getCell(6).value = { formula: `IF(B${r}="","",Konten!E${kr})` };

    luxuryCell(row.getCell(7), { isAlt, numFmt: currencyFmt, font: { bold: true, size: 13, color: { argb: c.green1 }, name: "Segoe UI" } });
    row.getCell(7).value = { formula: `IF(B${r}="","",Konten!F${kr})` };
  }

  // ---- CATEGORY BREAKDOWN (Row 36+) ----
  dash.mergeCells("B36:G36");
  dash.getCell("B36").value = "AUSGABEN NACH KATEGORIE";
  dash.getCell("B36").font = { bold: true, size: 14, color: { argb: c.textBright }, name: "Segoe UI" };
  dash.getCell("B36").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getRow(36).height = 32;

  dash.mergeCells("B37:G37");
  dash.getCell("B37").value = "Verteilung der Ausgaben im aktuellen Monat";
  dash.getCell("B37").font = { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" };
  dash.getCell("B37").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  luxuryHeader(dash, 38, ["", "Kategorie", "Betrag", "Anteil", "Visualisierung", "", ""], 7);

  for (let i = 0; i < 8; i++) {
    const row = dash.getRow(39 + i);
    const r = 39 + i;
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });

    luxuryCell(row.getCell(2), { isAlt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = { formula: `IF(H${13 + i}="","",H${13 + i})` };

    luxuryCell(row.getCell(3), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" } });
    row.getCell(3).value = { formula: `IF(B${r}="","",J${13 + i})` };

    luxuryCell(row.getCell(4), { isAlt, numFmt: pctFmt, font: { size: 11, color: { argb: c.yellow }, name: "Segoe UI" } });
    row.getCell(4).value = { formula: `IF(OR(B${r}="",SUM(C$39:C$46)=0),"",C${r}/SUM(C$39:C$46))` };

    // Visual bar using repeated character
    luxuryCell(row.getCell(5), { isAlt, font: { size: 10, color: { argb: c.green1 }, name: "Segoe UI" } });
    row.getCell(5).value = { formula: `IF(D${r}="","",REPT("█",ROUND(D${r}*30,0)))` };
    dash.mergeCells(r, 5, r, 7);
  }

  // ---- TOP TRANSACTIONS (Row 36+, Cols H-K) ----
  dash.mergeCells("H36:K36");
  dash.getCell("H36").value = "LETZTE TRANSAKTIONEN";
  dash.getCell("H36").font = { bold: true, size: 14, color: { argb: c.textBright }, name: "Segoe UI" };
  dash.getCell("H36").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  dash.mergeCells("H37:K37");
  dash.getCell("H37").value = "Die letzten 8 Buchungen";
  dash.getCell("H37").font = { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" };
  dash.getCell("H37").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };

  luxuryHeader(dash, 38, [null, null, null, null, null, null, null, "Beschreibung", "Datum", "Betrag", "Typ"], 11);

  for (let i = 0; i < 8; i++) {
    const row = dash.getRow(39 + i);
    const r = 39 + i;
    const isAlt = i % 2 === 1;
    const txr = 3 + i;

    luxuryCell(row.getCell(8), { isAlt, font: { size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(8).value = { formula: `IF(Transaktionen!A${txr}="","",Transaktionen!A${txr})` };

    luxuryCell(row.getCell(9), { isAlt, numFmt: dateFmt, font: { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" } });
    row.getCell(9).value = { formula: `IF(H${r}="","",Transaktionen!B${txr})` };

    luxuryCell(row.getCell(10), { isAlt, numFmt: currencyFmt, font: { bold: true, size: 11, color: { argb: c.green3 }, name: "Segoe UI" } });
    row.getCell(10).value = { formula: `IF(H${r}="","",Transaktionen!E${txr})` };

    luxuryCell(row.getCell(11), { isAlt, font: { bold: true, size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" }, align: "center" });
    row.getCell(11).value = { formula: `IF(H${r}="","",Transaktionen!C${txr})` };
  }

  // Conditional formatting for transaction types
  dash.addConditionalFormatting({
    ref: "K39:K46",
    rules: [
      { type: "containsText", operator: "containsText", text: "Einnahme",
        style: { font: { bold: true, color: { argb: c.green1 } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.greenDark } } } },
      { type: "containsText", operator: "containsText", text: "Ausgabe",
        style: { font: { bold: true, color: { argb: c.red } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.redBg } } } },
    ],
  });

  // Footer
  dash.mergeCells("B49:K49");
  dash.getCell("B49").value = "Powered by Budget Tracker Pro · All data updates automatically";
  dash.getCell("B49").font = { size: 9, color: { argb: c.textTertiary }, name: "Segoe UI", italic: true };
  dash.getCell("B49").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  dash.getCell("B49").alignment = { horizontal: "center" };

  // Protect dashboard
  dash.protect("", { selectLockedCells: true, selectUnlockedCells: true });

  // ============================================================
  // 2. TRANSAKTIONEN
  // ============================================================
  const tx = workbook.addWorksheet("Transaktionen", {
    properties: { tabColor: { argb: c.green1 } },
  });

  // Dark fill for all
  for (let r = 1; r <= 110; r++) fillRowDark(tx, r, 8);

  [3, 30, 16, 14, 20, 18, 18, 28].forEach((w, i) => tx.getColumn(i + 1).width = w);

  // Title
  tx.mergeCells("B1:G1");
  tx.getCell("B1").value = "TRANSAKTIONEN";
  tx.getCell("B1").font = { bold: true, size: 22, color: { argb: c.green1 }, name: "Segoe UI" };
  tx.getCell("B1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  tx.getRow(1).height = 45;

  tx.mergeCells("B2:G2");
  tx.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  tx.getCell("B2").border = { bottom: { style: "medium", color: { argb: c.green1 } }, top: noBorder.top, left: noBorder.left, right: noBorder.right };
  tx.getRow(2).height = 4;

  luxuryHeader(tx, 3, ["", "Beschreibung", "Typ", "Kategorie", "Betrag", "Konto", "Notizen", ""], 8);

  const sampleTx = [
    ["Gehalt Februar", "Einnahme", "Gehalt", 3500, "Hauptkonto", "Monatsgehalt"],
    ["Freelance Projekt", "Einnahme", "Freelance", 800, "Hauptkonto", "Webdesign"],
    ["Miete", "Ausgabe", "Wohnung", 950, "Hauptkonto", "Monatsmiete"],
    ["Einkauf REWE", "Ausgabe", "Lebensmittel", 67.50, "Bargeld", "Wocheneinkauf"],
    ["Tankstelle", "Ausgabe", "Transport", 55, "Kreditkarte", "Benzin"],
    ["Netflix", "Ausgabe", "Unterhaltung", 12.99, "Hauptkonto", "Abo"],
    ["Arztbesuch", "Ausgabe", "Gesundheit", 30, "Hauptkonto", "Zuzahlung"],
    ["Restaurant Besuch", "Ausgabe", "Restaurant", 45.80, "Kreditkarte", "Abendessen"],
    ["Online Kurs", "Ausgabe", "Bildung", 29.99, "Kreditkarte", "Udemy"],
    ["Einkauf LIDL", "Ausgabe", "Lebensmittel", 43.20, "Bargeld", "Wocheneinkauf"],
    ["T-Shirt", "Ausgabe", "Kleidung", 25, "Kreditkarte", "H&M"],
    ["Dividende", "Einnahme", "Investitionen", 120, "Sparkonto", "ETF Ausschüttung"],
  ];

  // Dates for samples (spread across Feb 2026)
  const sampleDates = [1,3,1,5,7,8,10,12,14,16,18,20].map(d => new Date(2026, 1, d));

  sampleTx.forEach((t, i) => {
    const row = tx.getRow(4 + i);
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });

    // Beschreibung
    luxuryCell(row.getCell(2), { isAlt, font: { size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = t[0];

    // Datum (stored in column B but displayed)
    // Actually we need a date column. Let me adjust - Date in B, rest shifted
    // Wait, the header has: Beschreibung, Typ, Kategorie, Betrag, Konto, Notizen
    // Let me add Datum between Beschreibung and Typ
  });

  // Actually, let me redo with proper date column
  // Clear and redo
  luxuryHeader(tx, 3, ["", "Beschreibung", "Datum", "Typ", "Betrag", "Kategorie", "Konto", "Notizen"], 8);

  sampleTx.forEach((t, i) => {
    const row = tx.getRow(4 + i);
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });

    luxuryCell(row.getCell(2), { isAlt, font: { size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = t[0]; // Beschreibung

    luxuryCell(row.getCell(3), { isAlt, numFmt: dateFmt, font: { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" } });
    row.getCell(3).value = sampleDates[i]; // Datum

    // Typ
    const isIncome = t[1] === "Einnahme";
    luxuryCell(row.getCell(4), {
      isAlt,
      font: { bold: true, size: 10, color: { argb: isIncome ? c.green1 : c.red }, name: "Segoe UI" },
      fg: isIncome ? c.greenDark : c.redBg,
      align: "center",
    });
    row.getCell(4).value = t[1];

    // Betrag
    luxuryCell(row.getCell(5), {
      isAlt,
      numFmt: currencyFmt,
      font: { bold: true, size: 12, color: { argb: isIncome ? c.green1 : c.redSoft }, name: "Segoe UI" },
    });
    row.getCell(5).value = t[3];

    // Kategorie
    luxuryCell(row.getCell(6), { isAlt, font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" } });
    row.getCell(6).value = t[2];

    // Konto
    luxuryCell(row.getCell(7), { isAlt, font: { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" } });
    row.getCell(7).value = t[4];

    // Notizen
    luxuryCell(row.getCell(8), { isAlt, font: { size: 10, color: { argb: c.textTertiary }, name: "Segoe UI", italic: true } });
    row.getCell(8).value = t[5];
  });

  // Now the column mapping for formulas is:
  // A=spacer, B=Beschreibung, C=Datum, D=Typ, E=Betrag, F=Kategorie, G=Konto, H=Notizen

  // Empty rows for user data
  for (let i = sampleTx.length; i < 100; i++) {
    const row = tx.getRow(4 + i);
    const isAlt = i % 2 === 1;
    luxuryCell(row.getCell(1), { fg: c.bgDark });
    for (let col = 2; col <= 8; col++) {
      luxuryCell(row.getCell(col), { isAlt });
      if (col === 3) row.getCell(col).numFmt = dateFmt;
      if (col === 5) row.getCell(col).numFmt = currencyFmt;
    }
  }

  // Data validation
  for (let i = 4; i <= 103; i++) {
    tx.getCell(`D${i}`).dataValidation = { type: "list", allowBlank: true, formulae: ['"Einnahme,Ausgabe"'] };
    tx.getCell(`F${i}`).dataValidation = { type: "list", allowBlank: true, formulae: ["Kategorien!$A$4:$A$20"] };
    tx.getCell(`G${i}`).dataValidation = { type: "list", allowBlank: true, formulae: ["Konten!$B$4:$B$12"] };
  }

  // Conditional formatting for type
  tx.addConditionalFormatting({
    ref: "D4:D103",
    rules: [
      { type: "containsText", operator: "containsText", text: "Einnahme",
        style: { font: { bold: true, color: { argb: c.green1 } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.greenDark } } } },
      { type: "containsText", operator: "containsText", text: "Ausgabe",
        style: { font: { bold: true, color: { argb: c.red } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.redBg } } } },
    ],
  });

  tx.autoFilter = { from: "B3", to: "H103" };
  tx.views = [{ state: "frozen", ySplit: 3, activeCell: "B4" }];

  // ============================================================
  // 3. KATEGORIEN
  // ============================================================
  const cats = workbook.addWorksheet("Kategorien", {
    properties: { tabColor: { argb: c.orange } },
  });

  for (let r = 1; r <= 30; r++) fillRowDark(cats, r, 6);
  [3, 24, 16, 18, 16, 3].forEach((w, i) => cats.getColumn(i + 1).width = w);

  cats.mergeCells("B1:E1");
  cats.getCell("B1").value = "KATEGORIEN";
  cats.getCell("B1").font = { bold: true, size: 22, color: { argb: c.green1 }, name: "Segoe UI" };
  cats.getCell("B1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  cats.getRow(1).height = 45;

  cats.mergeCells("B2:E2");
  cats.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  cats.getCell("B2").border = { bottom: { style: "medium", color: { argb: c.green1 } }, top: noBorder.top, left: noBorder.left, right: noBorder.right };
  cats.getRow(2).height = 4;

  luxuryHeader(cats, 3, ["", "Kategorie", "Typ", "Monatsbudget", "Farbe", ""], 6);

  const defaultCats = [
    ["Gehalt", "Einnahme", 0, "Grün"],
    ["Freelance", "Einnahme", 0, "Blau"],
    ["Investitionen", "Einnahme", 0, "Lila"],
    ["Sonstiges Einkommen", "Einnahme", 0, "Grau"],
    ["Lebensmittel", "Ausgabe", 400, "Rot"],
    ["Transport", "Ausgabe", 150, "Orange"],
    ["Wohnung", "Ausgabe", 1000, "Gelb"],
    ["Unterhaltung", "Ausgabe", 100, "Pink"],
    ["Gesundheit", "Ausgabe", 50, "Türkis"],
    ["Bildung", "Ausgabe", 50, "Indigo"],
    ["Kleidung", "Ausgabe", 80, "Rosa"],
    ["Restaurant", "Ausgabe", 100, "Lila"],
  ];

  defaultCats.forEach((cat, i) => {
    const row = cats.getRow(4 + i);
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });
    luxuryCell(row.getCell(2), { isAlt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = cat[0];

    const isInc = cat[1] === "Einnahme";
    luxuryCell(row.getCell(3), { isAlt, font: { bold: true, size: 10, color: { argb: isInc ? c.green1 : c.red }, name: "Segoe UI" }, fg: isInc ? c.greenDark : c.redBg, align: "center" });
    row.getCell(3).value = cat[1];

    luxuryCell(row.getCell(4), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" } });
    row.getCell(4).value = cat[2];

    luxuryCell(row.getCell(5), { isAlt, font: { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" } });
    row.getCell(5).value = cat[3];

    luxuryCell(row.getCell(6), { fg: c.bgDark });
  });

  for (let i = defaultCats.length; i < 16; i++) {
    const row = cats.getRow(4 + i);
    const isAlt = i % 2 === 1;
    luxuryCell(row.getCell(1), { fg: c.bgDark });
    for (let col = 2; col <= 5; col++) {
      luxuryCell(row.getCell(col), { isAlt });
      if (col === 4) row.getCell(col).numFmt = currencyFmt;
    }
    luxuryCell(row.getCell(6), { fg: c.bgDark });
  }

  for (let i = 4; i <= 19; i++) {
    cats.getCell(`C${i}`).dataValidation = { type: "list", allowBlank: true, formulae: ['"Einnahme,Ausgabe"'] };
  }

  cats.views = [{ state: "frozen", ySplit: 3, activeCell: "B4" }];

  // ============================================================
  // 4. BUDGETS
  // ============================================================
  const bud = workbook.addWorksheet("Budgets", {
    properties: { tabColor: { argb: c.purple } },
  });

  for (let r = 1; r <= 25; r++) fillRowDark(bud, r, 9);
  [3, 22, 18, 18, 18, 14, 16, 28, 3].forEach((w, i) => bud.getColumn(i + 1).width = w);

  bud.mergeCells("B1:H1");
  bud.getCell("B1").value = "MONATLICHE BUDGETS";
  bud.getCell("B1").font = { bold: true, size: 22, color: { argb: c.green1 }, name: "Segoe UI" };
  bud.getCell("B1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  bud.getRow(1).height = 45;

  bud.mergeCells("B2:H2");
  bud.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  bud.getCell("B2").border = { bottom: { style: "medium", color: { argb: c.green1 } }, top: noBorder.top, left: noBorder.left, right: noBorder.right };
  bud.getRow(2).height = 4;

  luxuryHeader(bud, 3, ["", "Kategorie", "Budget", "Ausgegeben", "Verbleibend", "Auslastung", "Status", "Fortschritt", ""], 9);

  const expCategories = defaultCats.filter(ct => ct[1] === "Ausgabe");
  expCategories.forEach((cat, i) => {
    const row = bud.getRow(4 + i);
    const r = 4 + i;
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });

    luxuryCell(row.getCell(2), { isAlt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = cat[0];

    luxuryCell(row.getCell(3), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" } });
    row.getCell(3).value = cat[2];

    luxuryCell(row.getCell(4), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" } });
    row.getCell(4).value = { formula: `SUMPRODUCT((Transaktionen!D$4:D$500="Ausgabe")*(Transaktionen!F$4:F$500=B${r})*(MONTH(Transaktionen!C$4:C$500)=MONTH(TODAY()))*(YEAR(Transaktionen!C$4:C$500)=YEAR(TODAY()))*Transaktionen!E$4:E$500)` };

    luxuryCell(row.getCell(5), { isAlt, numFmt: currencyFmt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(5).value = { formula: `C${r}-D${r}` };

    luxuryCell(row.getCell(6), { isAlt, numFmt: pctFmt, font: { size: 11, color: { argb: c.yellow }, name: "Segoe UI" }, align: "center" });
    row.getCell(6).value = { formula: `IF(C${r}=0,0,D${r}/C${r})` };

    luxuryCell(row.getCell(7), { isAlt, font: { bold: true, size: 10, name: "Segoe UI" }, align: "center" });
    row.getCell(7).value = { formula: `IF(C${r}=0,"—",IF(F${r}>1,"OVER",IF(F${r}>0.8,"WARN","OK")))` };

    luxuryCell(row.getCell(8), { isAlt, font: { size: 9, color: { argb: c.green1 }, name: "Segoe UI" } });
    row.getCell(8).value = { formula: `IF(C${r}=0,"",REPT("█",MIN(ROUND(F${r}*20,0),20))&REPT("░",MAX(20-ROUND(F${r}*20,0),0)))` };

    luxuryCell(row.getCell(9), { fg: c.bgDark });
  });

  // Budget conditional formatting
  bud.addConditionalFormatting({
    ref: "G4:G14",
    rules: [
      { type: "containsText", operator: "containsText", text: "OVER",
        style: { font: { bold: true, color: { argb: c.red } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.redBg } } } },
      { type: "containsText", operator: "containsText", text: "WARN",
        style: { font: { bold: true, color: { argb: c.orange } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.orangeBg } } } },
      { type: "containsText", operator: "containsText", text: "OK",
        style: { font: { bold: true, color: { argb: c.green1 } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.greenDark } } } },
    ],
  });

  bud.addConditionalFormatting({
    ref: "F4:F14",
    rules: [
      { type: "colorScale",
        cfvo: [{ type: "num", value: 0 }, { type: "num", value: 0.5 }, { type: "num", value: 1 }],
        color: [{ argb: c.greenDark }, { argb: c.yellowBg }, { argb: c.redBg }] },
    ],
  });

  // Total
  const budTotal = bud.getRow(4 + expCategories.length);
  const btr = 4 + expCategories.length;
  budTotal.height = 32;
  for (let col = 1; col <= 9; col++) {
    budTotal.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.greenDark } };
    budTotal.getCell(col).border = { top: { style: "medium", color: { argb: c.green1 } }, bottom: { style: "medium", color: { argb: c.green1 } }, left: darkBorder.left, right: darkBorder.right };
    budTotal.getCell(col).font = { bold: true, size: 12, color: { argb: c.green1 }, name: "Segoe UI" };
    budTotal.getCell(col).alignment = { vertical: "middle" };
  }
  budTotal.getCell(2).value = "TOTAL";
  budTotal.getCell(3).value = { formula: `SUM(C4:C${btr - 1})` };
  budTotal.getCell(3).numFmt = currencyFmt;
  budTotal.getCell(4).value = { formula: `SUM(D4:D${btr - 1})` };
  budTotal.getCell(4).numFmt = currencyFmt;
  budTotal.getCell(5).value = { formula: `C${btr}-D${btr}` };
  budTotal.getCell(5).numFmt = currencyFmt;
  budTotal.getCell(6).value = { formula: `IF(C${btr}=0,0,D${btr}/C${btr})` };
  budTotal.getCell(6).numFmt = pctFmt;

  bud.views = [{ state: "frozen", ySplit: 3, activeCell: "B4" }];

  // ============================================================
  // 5. KONTEN
  // ============================================================
  const kon = workbook.addWorksheet("Konten", {
    properties: { tabColor: { argb: c.yellow } },
  });

  for (let r = 1; r <= 25; r++) fillRowDark(kon, r, 8);
  [3, 22, 16, 20, 18, 18, 22, 3].forEach((w, i) => kon.getColumn(i + 1).width = w);

  kon.mergeCells("B1:G1");
  kon.getCell("B1").value = "KONTEN";
  kon.getCell("B1").font = { bold: true, size: 22, color: { argb: c.green1 }, name: "Segoe UI" };
  kon.getCell("B1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  kon.getRow(1).height = 45;

  kon.mergeCells("B2:G2");
  kon.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  kon.getCell("B2").border = { bottom: { style: "medium", color: { argb: c.green1 } }, top: noBorder.top, left: noBorder.left, right: noBorder.right };
  kon.getRow(2).height = 4;

  luxuryHeader(kon, 3, ["", "Kontoname", "Typ", "Startguthaben", "Einnahmen", "Ausgaben", "Aktueller Stand", ""], 8);

  const defAccounts = [
    ["Hauptkonto", "Bank", 2500],
    ["Bargeld", "Bargeld", 200],
    ["Kreditkarte", "Kreditkarte", 0],
    ["Sparkonto", "Sparkonto", 5000],
  ];

  defAccounts.forEach((acc, i) => {
    const row = kon.getRow(4 + i);
    const r = 4 + i;
    const isAlt = i % 2 === 1;

    luxuryCell(row.getCell(1), { fg: c.bgDark });

    luxuryCell(row.getCell(2), { isAlt, font: { bold: true, size: 11, color: { argb: c.textBright }, name: "Segoe UI" } });
    row.getCell(2).value = acc[0];

    luxuryCell(row.getCell(3), { isAlt, font: { size: 10, color: { argb: c.textSecondary }, name: "Segoe UI" }, align: "center" });
    row.getCell(3).value = acc[1];

    luxuryCell(row.getCell(4), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" } });
    row.getCell(4).value = acc[2];

    luxuryCell(row.getCell(5), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.green3 }, name: "Segoe UI" } });
    row.getCell(5).value = { formula: `SUMPRODUCT((Transaktionen!D$4:D$500="Einnahme")*(Transaktionen!G$4:G$500=B${r})*Transaktionen!E$4:E$500)` };

    luxuryCell(row.getCell(6), { isAlt, numFmt: currencyFmt, font: { size: 11, color: { argb: c.redSoft }, name: "Segoe UI" } });
    row.getCell(6).value = { formula: `SUMPRODUCT((Transaktionen!D$4:D$500="Ausgabe")*(Transaktionen!G$4:G$500=B${r})*Transaktionen!E$4:E$500)` };

    luxuryCell(row.getCell(7), { isAlt, numFmt: currencyFmt, font: { bold: true, size: 14, color: { argb: c.green1 }, name: "Segoe UI" } });
    row.getCell(7).value = { formula: `D${r}+E${r}-F${r}` };

    luxuryCell(row.getCell(8), { fg: c.bgDark });
  });

  // Empty rows
  for (let i = defAccounts.length; i < 8; i++) {
    const row = kon.getRow(4 + i);
    const r = 4 + i;
    const isAlt = i % 2 === 1;
    luxuryCell(row.getCell(1), { fg: c.bgDark });
    for (let col = 2; col <= 7; col++) {
      luxuryCell(row.getCell(col), { isAlt });
      if (col >= 4) row.getCell(col).numFmt = currencyFmt;
    }
    row.getCell(5).value = { formula: `IF(B${r}="","",SUMPRODUCT((Transaktionen!D$4:D$500="Einnahme")*(Transaktionen!G$4:G$500=B${r})*Transaktionen!E$4:E$500))` };
    row.getCell(6).value = { formula: `IF(B${r}="","",SUMPRODUCT((Transaktionen!D$4:D$500="Ausgabe")*(Transaktionen!G$4:G$500=B${r})*Transaktionen!E$4:E$500))` };
    row.getCell(7).value = { formula: `IF(B${r}="","",D${r}+E${r}-F${r})` };
    luxuryCell(row.getCell(8), { fg: c.bgDark });
  }

  // Account type validation
  for (let i = 4; i <= 11; i++) {
    kon.getCell(`C${i}`).dataValidation = { type: "list", allowBlank: true, formulae: ['"Bank,Bargeld,Kreditkarte,Sparkonto"'] };
  }

  // Total
  const accTotal = kon.getRow(15);
  accTotal.height = 32;
  for (let col = 1; col <= 8; col++) {
    accTotal.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.greenDark } };
    accTotal.getCell(col).border = { top: { style: "medium", color: { argb: c.green1 } }, bottom: { style: "medium", color: { argb: c.green1 } }, left: darkBorder.left, right: darkBorder.right };
    accTotal.getCell(col).font = { bold: true, size: 12, color: { argb: c.green1 }, name: "Segoe UI" };
    accTotal.getCell(col).alignment = { vertical: "middle" };
  }
  accTotal.getCell(2).value = "TOTAL";
  accTotal.getCell(4).value = { formula: "SUM(D4:D11)" };
  accTotal.getCell(4).numFmt = currencyFmt;
  accTotal.getCell(5).value = { formula: "SUM(E4:E11)" };
  accTotal.getCell(5).numFmt = currencyFmt;
  accTotal.getCell(6).value = { formula: "SUM(F4:F11)" };
  accTotal.getCell(6).numFmt = currencyFmt;
  accTotal.getCell(7).value = { formula: "SUM(G4:G11)" };
  accTotal.getCell(7).numFmt = currencyFmt;

  kon.addConditionalFormatting({
    ref: "G4:G11",
    rules: [
      { type: "cellIs", operator: "lessThan", formulae: [0],
        style: { font: { bold: true, color: { argb: c.red } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: c.redBg } } } },
    ],
  });

  kon.views = [{ state: "frozen", ySplit: 3, activeCell: "B4" }];

  // ============================================================
  // 6. ANLEITUNG
  // ============================================================
  const help = workbook.addWorksheet("Anleitung", {
    properties: { tabColor: { argb: c.textTertiary } },
  });

  for (let r = 1; r <= 40; r++) fillRowDark(help, r, 2);
  help.getColumn(1).width = 3;
  help.getColumn(2).width = 90;

  help.getCell("B1").value = "BUDGET TRACKER PRO";
  help.getCell("B1").font = { bold: true, size: 24, color: { argb: c.green1 }, name: "Segoe UI" };
  help.getCell("B1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
  help.getRow(1).height = 50;

  const helpLines = [
    ["Anleitung & Schnellstart", { bold: true, size: 12, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["SO FUNKTIONIERT ES:", { bold: true, size: 14, color: { argb: c.green1 }, name: "Segoe UI" }],
    ["", null],
    ["1. TRANSAKTIONEN", { bold: true, size: 13, color: { argb: c.cyan }, name: "Segoe UI" }],
    ["   Trage jede Einnahme und Ausgabe im Tab 'Transaktionen' ein.", { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" }],
    ["   Nutze die Dropdown-Menüs für Typ, Kategorie und Konto.", { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["2. KATEGORIEN", { bold: true, size: 13, color: { argb: c.cyan }, name: "Segoe UI" }],
    ["   Passe deine Kategorien an oder füge neue hinzu.", { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" }],
    ["   Setze Monatsbudgets für Ausgabe-Kategorien.", { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["3. BUDGETS", { bold: true, size: 13, color: { argb: c.cyan }, name: "Segoe UI" }],
    ["   Sieh auf einen Blick, ob du im Budget bist.", { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" }],
    ["   Fortschrittsbalken und Ampel-Status zeigen deinen Stand.", { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["4. KONTEN", { bold: true, size: 13, color: { argb: c.cyan }, name: "Segoe UI" }],
    ["   Verwalte Bank, Bargeld, Kreditkarte und Sparkonto.", { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" }],
    ["   Salden werden automatisch berechnet.", { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["5. DASHBOARD", { bold: true, size: 13, color: { argb: c.cyan }, name: "Segoe UI" }],
    ["   Alles aktualisiert sich automatisch!", { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" }],
    ["   Stat-Cards, Monatsübersicht, Budget-Status und mehr.", { size: 11, color: { argb: c.textSecondary }, name: "Segoe UI" }],
    ["", null],
    ["TIPP: Lösche die Beispieldaten und starte mit deinen eigenen!", { bold: true, size: 12, color: { argb: c.yellow }, name: "Segoe UI" }],
  ];

  helpLines.forEach(([text, font], i) => {
    const cell = help.getCell(`B${2 + i}`);
    cell.value = text;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.bgDark } };
    if (font) cell.font = font;
    else cell.font = { size: 11, color: { argb: c.textPrimary }, name: "Segoe UI" };
  });

  // ============================================================
  // UPDATE DASHBOARD FORMULAS - fix column references
  // ============================================================
  // Transaktionen columns: B=Beschreibung, C=Datum, D=Typ, E=Betrag, F=Kategorie, G=Konto
  // Dashboard formulas need to reference: D=Typ (Einnahme/Ausgabe), C=Datum, E=Betrag, F=Kategorie

  // Fix stat card formulas
  dash.getCell("D7").value = { formula: 'SUMPRODUCT((Transaktionen!D4:D500="Einnahme")*(MONTH(Transaktionen!C4:C500)=MONTH(TODAY()))*(YEAR(Transaktionen!C4:C500)=YEAR(TODAY()))*Transaktionen!E4:E500)' };
  dash.getCell("F7").value = { formula: 'SUMPRODUCT((Transaktionen!D4:D500="Ausgabe")*(MONTH(Transaktionen!C4:C500)=MONTH(TODAY()))*(YEAR(Transaktionen!C4:C500)=YEAR(TODAY()))*Transaktionen!E4:E500)' };

  // Fix monthly overview formulas
  for (let i = 0; i < 12; i++) {
    const r = 13 + i;
    dash.getCell(`C${r}`).value = { formula: `SUMPRODUCT((Transaktionen!D$4:D$500="Einnahme")*(MONTH(Transaktionen!C$4:C$500)=${i + 1})*(YEAR(Transaktionen!C$4:C$500)=YEAR(TODAY()))*Transaktionen!E$4:E$500)` };
    dash.getCell(`D${r}`).value = { formula: `SUMPRODUCT((Transaktionen!D$4:D$500="Ausgabe")*(MONTH(Transaktionen!C$4:C$500)=${i + 1})*(YEAR(Transaktionen!C$4:C$500)=YEAR(TODAY()))*Transaktionen!E$4:E$500)` };
  }

  // Fix budget status formulas (column F=Kategorie in Transaktionen)
  for (let i = 0; i < 8; i++) {
    const r = 13 + i;
    dash.getCell(`J${r}`).value = { formula: `IF(H${r}="","",SUMPRODUCT((Transaktionen!D$4:D$500="Ausgabe")*(Transaktionen!F$4:F$500=H${r})*(MONTH(Transaktionen!C$4:C$500)=MONTH(TODAY()))*(YEAR(Transaktionen!C$4:C$500)=YEAR(TODAY()))*Transaktionen!E$4:E$500))` };
  }

  // Fix Konten reference on dashboard
  dash.getCell("B7").value = { formula: "Konten!G15" };

  // ============================================================
  // FINAL
  // ============================================================
  workbook.views = [{ activeTab: 0 }];

  const outputPath = path.join(__dirname, "Budget-Tracker.xlsx");
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Premium Budget Tracker erstellt: ${outputPath}`);
}

generateBudgetTracker().catch(console.error);
