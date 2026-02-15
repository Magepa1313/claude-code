const ExcelJS = require("exceljs");
const path = require("path");

async function generateBudgetTracker() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Budget Tracker";
  workbook.created = new Date();

  // ============================================================
  // Color palette
  // ============================================================
  const colors = {
    primary: "2563EB",
    primaryLight: "DBEAFE",
    green: "16A34A",
    greenLight: "DCFCE7",
    greenBg: "F0FDF4",
    red: "DC2626",
    redLight: "FEE2E2",
    redBg: "FEF2F2",
    orange: "EA580C",
    orangeLight: "FFEDD5",
    purple: "7C3AED",
    purpleLight: "EDE9FE",
    yellow: "CA8A04",
    yellowLight: "FEF9C3",
    gray: "6B7280",
    grayLight: "F3F4F6",
    grayBorder: "E5E7EB",
    dark: "1F2937",
    white: "FFFFFF",
    headerBg: "1E3A5F",
    headerText: "FFFFFF",
    altRow: "F8FAFC",
  };

  const headerFont = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
  const titleFont = { bold: true, color: { argb: colors.dark }, size: 14, name: "Calibri" };
  const subtitleFont = { bold: true, color: { argb: colors.gray }, size: 10, name: "Calibri" };
  const normalFont = { size: 11, name: "Calibri", color: { argb: colors.dark } };
  const currencyFormat = '#,##0.00 €';
  const percentFormat = '0.0%';
  const dateFormat = 'DD.MM.YYYY';

  const thinBorder = {
    top: { style: "thin", color: { argb: colors.grayBorder } },
    left: { style: "thin", color: { argb: colors.grayBorder } },
    bottom: { style: "thin", color: { argb: colors.grayBorder } },
    right: { style: "thin", color: { argb: colors.grayBorder } },
  };

  function styleHeader(row, colCount) {
    row.height = 30;
    for (let i = 1; i <= colCount; i++) {
      const cell = row.getCell(i);
      cell.font = headerFont;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = thinBorder;
    }
  }

  function styleDataCell(cell, isAlt = false) {
    cell.font = normalFont;
    cell.border = thinBorder;
    cell.alignment = { vertical: "middle" };
    if (isAlt) {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.altRow } };
    }
  }

  // ============================================================
  // 1. DASHBOARD Sheet
  // ============================================================
  const dashboard = workbook.addWorksheet("Dashboard", {
    properties: { tabColor: { argb: colors.primary } },
  });

  dashboard.properties.defaultRowHeight = 20;

  // Title
  dashboard.mergeCells("A1:H1");
  const dashTitle = dashboard.getCell("A1");
  dashTitle.value = "BUDGET TRACKER - DASHBOARD";
  dashTitle.font = { bold: true, size: 20, color: { argb: colors.primary }, name: "Calibri" };
  dashTitle.alignment = { horizontal: "center", vertical: "middle" };
  dashboard.getRow(1).height = 45;

  dashboard.mergeCells("A2:H2");
  const dashSubtitle = dashboard.getCell("A2");
  dashSubtitle.value = "Deine Finanzen im Überblick";
  dashSubtitle.font = { size: 11, color: { argb: colors.gray }, name: "Calibri" };
  dashSubtitle.alignment = { horizontal: "center" };

  // Summary boxes - Row 4-7
  // Box 1: Gesamtguthaben
  dashboard.mergeCells("A4:B4");
  dashboard.getCell("A4").value = "GESAMTGUTHABEN";
  dashboard.getCell("A4").font = { bold: true, size: 9, color: { argb: colors.primary }, name: "Calibri" };
  dashboard.getCell("A4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.primaryLight } };
  dashboard.getCell("B4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.primaryLight } };

  dashboard.mergeCells("A5:B5");
  const balanceCell = dashboard.getCell("A5");
  balanceCell.value = { formula: "Konten!E15" };
  balanceCell.numFmt = currencyFormat;
  balanceCell.font = { bold: true, size: 22, color: { argb: colors.primary }, name: "Calibri" };
  balanceCell.alignment = { horizontal: "center" };
  balanceCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.primaryLight } };
  dashboard.getCell("B5").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.primaryLight } };

  // Box 2: Einnahmen (diesen Monat)
  dashboard.mergeCells("C4:D4");
  dashboard.getCell("C4").value = "EINNAHMEN (MONAT)";
  dashboard.getCell("C4").font = { bold: true, size: 9, color: { argb: colors.green }, name: "Calibri" };
  dashboard.getCell("C4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.greenLight } };
  dashboard.getCell("D4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.greenLight } };

  dashboard.mergeCells("C5:D5");
  const incomeCell = dashboard.getCell("C5");
  incomeCell.value = { formula: 'SUMPRODUCT((Transaktionen!C3:C1000="Einnahme")*(MONTH(Transaktionen!B3:B1000)=MONTH(TODAY()))*(YEAR(Transaktionen!B3:B1000)=YEAR(TODAY()))*Transaktionen!E3:E1000)' };
  incomeCell.numFmt = currencyFormat;
  incomeCell.font = { bold: true, size: 22, color: { argb: colors.green }, name: "Calibri" };
  incomeCell.alignment = { horizontal: "center" };
  incomeCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.greenLight } };
  dashboard.getCell("D5").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.greenLight } };

  // Box 3: Ausgaben (diesen Monat)
  dashboard.mergeCells("E4:F4");
  dashboard.getCell("E4").value = "AUSGABEN (MONAT)";
  dashboard.getCell("E4").font = { bold: true, size: 9, color: { argb: colors.red }, name: "Calibri" };
  dashboard.getCell("E4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.redLight } };
  dashboard.getCell("F4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.redLight } };

  dashboard.mergeCells("E5:F5");
  const expenseCell = dashboard.getCell("E5");
  expenseCell.value = { formula: 'SUMPRODUCT((Transaktionen!C3:C1000="Ausgabe")*(MONTH(Transaktionen!B3:B1000)=MONTH(TODAY()))*(YEAR(Transaktionen!B3:B1000)=YEAR(TODAY()))*Transaktionen!E3:E1000)' };
  expenseCell.numFmt = currencyFormat;
  expenseCell.font = { bold: true, size: 22, color: { argb: colors.red }, name: "Calibri" };
  expenseCell.alignment = { horizontal: "center" };
  expenseCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.redLight } };
  dashboard.getCell("F5").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.redLight } };

  // Box 4: Bilanz
  dashboard.mergeCells("G4:H4");
  dashboard.getCell("G4").value = "BILANZ (MONAT)";
  dashboard.getCell("G4").font = { bold: true, size: 9, color: { argb: colors.purple }, name: "Calibri" };
  dashboard.getCell("G4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.purpleLight } };
  dashboard.getCell("H4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.purpleLight } };

  dashboard.mergeCells("G5:H5");
  const bilanzCell = dashboard.getCell("G5");
  bilanzCell.value = { formula: "C5-E5" };
  bilanzCell.numFmt = currencyFormat;
  bilanzCell.font = { bold: true, size: 22, color: { argb: colors.purple }, name: "Calibri" };
  bilanzCell.alignment = { horizontal: "center" };
  bilanzCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.purpleLight } };
  dashboard.getCell("H5").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.purpleLight } };

  // Add borders to summary boxes
  for (const row of [4, 5]) {
    for (let col = 1; col <= 8; col++) {
      dashboard.getRow(row).getCell(col).border = thinBorder;
    }
  }

  // Monthly overview section
  dashboard.getRow(7).height = 30;
  dashboard.mergeCells("A7:H7");
  dashboard.getCell("A7").value = "MONATSÜBERSICHT";
  dashboard.getCell("A7").font = { bold: true, size: 12, color: { argb: colors.dark }, name: "Calibri" };
  dashboard.getCell("A7").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.grayLight } };
  dashboard.getCell("A7").alignment = { vertical: "middle" };
  dashboard.getCell("A7").border = thinBorder;

  const months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

  // Month headers
  const monthHeaderRow = dashboard.getRow(8);
  dashboard.getCell("A8").value = "Monat";
  dashboard.getCell("B8").value = "Einnahmen";
  dashboard.getCell("C8").value = "Ausgaben";
  dashboard.getCell("D8").value = "Bilanz";
  dashboard.getCell("E8").value = "Sparquote";
  styleHeader(monthHeaderRow, 5);

  for (let i = 0; i < 12; i++) {
    const row = dashboard.getRow(9 + i);
    const isAlt = i % 2 === 1;

    const monthCell = row.getCell(1);
    monthCell.value = months[i];
    styleDataCell(monthCell, isAlt);
    monthCell.font = { ...normalFont, bold: true };

    const incCell = row.getCell(2);
    incCell.value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$1000="Einnahme")*(MONTH(Transaktionen!B$3:B$1000)=${i + 1})*(YEAR(Transaktionen!B$3:B$1000)=YEAR(TODAY()))*Transaktionen!E$3:E$1000)` };
    incCell.numFmt = currencyFormat;
    styleDataCell(incCell, isAlt);
    incCell.font = { ...normalFont, color: { argb: colors.green } };

    const expCell = row.getCell(3);
    expCell.value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$1000="Ausgabe")*(MONTH(Transaktionen!B$3:B$1000)=${i + 1})*(YEAR(Transaktionen!B$3:B$1000)=YEAR(TODAY()))*Transaktionen!E$3:E$1000)` };
    expCell.numFmt = currencyFormat;
    styleDataCell(expCell, isAlt);
    expCell.font = { ...normalFont, color: { argb: colors.red } };

    const balCell = row.getCell(4);
    balCell.value = { formula: `B${9 + i}-C${9 + i}` };
    balCell.numFmt = currencyFormat;
    styleDataCell(balCell, isAlt);
    balCell.font = { ...normalFont, bold: true };

    const rateCell = row.getCell(5);
    rateCell.value = { formula: `IF(B${9 + i}=0,0,D${9 + i}/B${9 + i})` };
    rateCell.numFmt = percentFormat;
    styleDataCell(rateCell, isAlt);
  }

  // Year totals
  const totalRow = dashboard.getRow(21);
  totalRow.height = 28;
  const totalLabelCell = totalRow.getCell(1);
  totalLabelCell.value = "GESAMT";
  totalLabelCell.font = { bold: true, size: 11, color: { argb: colors.headerText }, name: "Calibri" };
  totalLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
  totalLabelCell.border = thinBorder;
  totalLabelCell.alignment = { vertical: "middle" };

  for (let col = 2; col <= 5; col++) {
    const cell = totalRow.getCell(col);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
    cell.border = thinBorder;
    cell.alignment = { vertical: "middle" };
    cell.font = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
  }
  totalRow.getCell(2).value = { formula: "SUM(B9:B20)" };
  totalRow.getCell(2).numFmt = currencyFormat;
  totalRow.getCell(3).value = { formula: "SUM(C9:C20)" };
  totalRow.getCell(3).numFmt = currencyFormat;
  totalRow.getCell(4).value = { formula: "B21-C21" };
  totalRow.getCell(4).numFmt = currencyFormat;
  totalRow.getCell(5).value = { formula: "IF(B21=0,0,D21/B21)" };
  totalRow.getCell(5).numFmt = percentFormat;

  // Budget status section
  dashboard.getRow(23).height = 30;
  dashboard.mergeCells("A23:H23");
  dashboard.getCell("A23").value = "BUDGET STATUS (AKTUELLER MONAT)";
  dashboard.getCell("A23").font = { bold: true, size: 12, color: { argb: colors.dark }, name: "Calibri" };
  dashboard.getCell("A23").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.grayLight } };
  dashboard.getCell("A23").alignment = { vertical: "middle" };

  const budgetHeaders = ["Kategorie", "Budget", "Ausgegeben", "Verbleibend", "Auslastung"];
  const bRow = dashboard.getRow(24);
  budgetHeaders.forEach((h, i) => {
    bRow.getCell(i + 1).value = h;
  });
  styleHeader(bRow, 5);

  for (let i = 0; i < 10; i++) {
    const row = dashboard.getRow(25 + i);
    const isAlt = i % 2 === 1;
    const r = 25 + i;

    // Category from Budgets sheet
    const catCell = row.getCell(1);
    catCell.value = { formula: `IF(Budgets!A${3 + i}="","",Budgets!A${3 + i})` };
    styleDataCell(catCell, isAlt);

    // Budget amount
    const budCell = row.getCell(2);
    budCell.value = { formula: `IF(A${r}="","",Budgets!B${3 + i})` };
    budCell.numFmt = currencyFormat;
    styleDataCell(budCell, isAlt);

    // Spent
    const spentCell = row.getCell(3);
    spentCell.value = { formula: `IF(A${r}="","",SUMPRODUCT((Transaktionen!C$3:C$1000="Ausgabe")*(Transaktionen!D$3:D$1000=A${r})*(MONTH(Transaktionen!B$3:B$1000)=MONTH(TODAY()))*(YEAR(Transaktionen!B$3:B$1000)=YEAR(TODAY()))*Transaktionen!E$3:E$1000))` };
    spentCell.numFmt = currencyFormat;
    styleDataCell(spentCell, isAlt);

    // Remaining
    const remCell = row.getCell(4);
    remCell.value = { formula: `IF(A${r}="","",B${r}-C${r})` };
    remCell.numFmt = currencyFormat;
    styleDataCell(remCell, isAlt);

    // Usage %
    const usageCell = row.getCell(5);
    usageCell.value = { formula: `IF(OR(A${r}="",B${r}=0),"",C${r}/B${r})` };
    usageCell.numFmt = percentFormat;
    styleDataCell(usageCell, isAlt);
  }

  // Set column widths for dashboard
  dashboard.getColumn(1).width = 16;
  dashboard.getColumn(2).width = 18;
  dashboard.getColumn(3).width = 18;
  dashboard.getColumn(4).width = 18;
  dashboard.getColumn(5).width = 14;
  dashboard.getColumn(6).width = 14;
  dashboard.getColumn(7).width = 14;
  dashboard.getColumn(8).width = 14;

  // ============================================================
  // 2. TRANSAKTIONEN Sheet
  // ============================================================
  const transactions = workbook.addWorksheet("Transaktionen", {
    properties: { tabColor: { argb: colors.green } },
  });

  // Title
  transactions.mergeCells("A1:G1");
  transactions.getCell("A1").value = "TRANSAKTIONEN";
  transactions.getCell("A1").font = { bold: true, size: 16, color: { argb: colors.primary }, name: "Calibri" };
  transactions.getCell("A1").alignment = { horizontal: "left", vertical: "middle" };
  transactions.getRow(1).height = 35;

  // Headers
  const txHeaders = ["Beschreibung", "Datum", "Typ", "Kategorie", "Betrag", "Konto", "Notizen"];
  const txHeaderRow = transactions.getRow(2);
  txHeaders.forEach((h, i) => {
    txHeaderRow.getCell(i + 1).value = h;
  });
  styleHeader(txHeaderRow, 7);

  // Sample data
  const sampleTransactions = [
    ["Gehalt Februar", new Date(2026, 1, 1), "Einnahme", "Gehalt", 3500, "Hauptkonto", "Monatsgehalt"],
    ["Freelance Projekt", new Date(2026, 1, 5), "Einnahme", "Freelance", 800, "Hauptkonto", "Webdesign"],
    ["Miete", new Date(2026, 1, 1), "Ausgabe", "Wohnung", 950, "Hauptkonto", "Monatsmiete"],
    ["Einkauf REWE", new Date(2026, 1, 3), "Ausgabe", "Lebensmittel", 67.50, "Bargeld", "Wocheneinkauf"],
    ["Tankstelle", new Date(2026, 1, 4), "Ausgabe", "Transport", 55, "Kreditkarte", "Benzin"],
    ["Netflix", new Date(2026, 1, 5), "Ausgabe", "Unterhaltung", 12.99, "Hauptkonto", "Abo"],
    ["Arztbesuch", new Date(2026, 1, 6), "Ausgabe", "Gesundheit", 30, "Hauptkonto", "Zuzahlung"],
    ["Restaurant", new Date(2026, 1, 7), "Ausgabe", "Restaurant", 45.80, "Kreditkarte", "Abendessen"],
    ["Online Kurs", new Date(2026, 1, 8), "Ausgabe", "Bildung", 29.99, "Kreditkarte", "Udemy"],
    ["Einkauf LIDL", new Date(2026, 1, 10), "Ausgabe", "Lebensmittel", 43.20, "Bargeld", ""],
    ["T-Shirt", new Date(2026, 1, 11), "Ausgabe", "Kleidung", 25, "Kreditkarte", "H&M"],
    ["Dividende", new Date(2026, 1, 15), "Einnahme", "Investitionen", 120, "Sparkonto", "ETF Ausschüttung"],
  ];

  sampleTransactions.forEach((tx, i) => {
    const row = transactions.getRow(3 + i);
    const isAlt = i % 2 === 1;

    for (let col = 1; col <= 7; col++) {
      const cell = row.getCell(col);
      cell.value = tx[col - 1];
      styleDataCell(cell, isAlt);
    }

    // Date format
    row.getCell(2).numFmt = dateFormat;

    // Type coloring
    const typeCell = row.getCell(3);
    if (tx[2] === "Einnahme") {
      typeCell.font = { ...normalFont, bold: true, color: { argb: colors.green } };
      typeCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.greenBg } };
    } else {
      typeCell.font = { ...normalFont, bold: true, color: { argb: colors.red } };
      typeCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.redBg } };
    }

    // Amount format + color
    const amountCell = row.getCell(5);
    amountCell.numFmt = currencyFormat;
    if (tx[2] === "Einnahme") {
      amountCell.font = { ...normalFont, bold: true, color: { argb: colors.green } };
    } else {
      amountCell.font = { ...normalFont, bold: true, color: { argb: colors.red } };
    }
  });

  // Add empty rows for user input (up to row 1000)
  for (let i = sampleTransactions.length; i < 100; i++) {
    const row = transactions.getRow(3 + i);
    const isAlt = i % 2 === 1;
    for (let col = 1; col <= 7; col++) {
      const cell = row.getCell(col);
      styleDataCell(cell, isAlt);
      if (col === 2) cell.numFmt = dateFormat;
      if (col === 5) cell.numFmt = currencyFormat;
    }
  }

  // Data validation for Type column
  for (let i = 3; i <= 102; i++) {
    transactions.getCell(`C${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ['"Einnahme,Ausgabe"'],
      showErrorMessage: true,
      errorTitle: "Ungültiger Typ",
      error: "Bitte wähle Einnahme oder Ausgabe",
    };
  }

  // Data validation for Category column
  for (let i = 3; i <= 102; i++) {
    transactions.getCell(`D${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ["Kategorien!$A$3:$A$20"],
      showErrorMessage: true,
      errorTitle: "Ungültige Kategorie",
      error: "Bitte wähle eine Kategorie aus der Liste",
    };
  }

  // Data validation for Account column
  for (let i = 3; i <= 102; i++) {
    transactions.getCell(`F${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ["Konten!$A$3:$A$12"],
      showErrorMessage: true,
      errorTitle: "Ungültiges Konto",
      error: "Bitte wähle ein Konto aus der Liste",
    };
  }

  // Summary row at bottom
  transactions.getRow(103).height = 5;
  const txSumRow = transactions.getRow(104);
  txSumRow.height = 28;

  transactions.mergeCells("A104:D104");
  txSumRow.getCell(1).value = "ZUSAMMENFASSUNG";
  txSumRow.getCell(1).font = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
  txSumRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
  txSumRow.getCell(1).alignment = { vertical: "middle" };
  for (let c = 2; c <= 4; c++) {
    txSumRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
  }
  txSumRow.getCell(5).value = { formula: "SUM(E3:E102)" };
  txSumRow.getCell(5).numFmt = currencyFormat;
  txSumRow.getCell(5).font = { bold: true, color: { argb: colors.headerText }, size: 12, name: "Calibri" };
  txSumRow.getCell(5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
  txSumRow.getCell(5).alignment = { vertical: "middle" };

  // Column widths
  transactions.getColumn(1).width = 28;
  transactions.getColumn(2).width = 14;
  transactions.getColumn(3).width = 12;
  transactions.getColumn(4).width = 18;
  transactions.getColumn(5).width = 16;
  transactions.getColumn(6).width = 16;
  transactions.getColumn(7).width = 25;

  // Auto filter
  transactions.autoFilter = { from: "A2", to: "G102" };

  // ============================================================
  // 3. KATEGORIEN Sheet
  // ============================================================
  const categories_ws = workbook.addWorksheet("Kategorien", {
    properties: { tabColor: { argb: colors.orange } },
  });

  categories_ws.mergeCells("A1:D1");
  categories_ws.getCell("A1").value = "KATEGORIEN";
  categories_ws.getCell("A1").font = { bold: true, size: 16, color: { argb: colors.primary }, name: "Calibri" };
  categories_ws.getRow(1).height = 35;

  const catHeaders = ["Kategorie", "Typ", "Monatsbudget", "Farbe"];
  const catHeaderRow = categories_ws.getRow(2);
  catHeaders.forEach((h, i) => {
    catHeaderRow.getCell(i + 1).value = h;
  });
  styleHeader(catHeaderRow, 4);

  const defaultCategories = [
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

  defaultCategories.forEach((cat, i) => {
    const row = categories_ws.getRow(3 + i);
    const isAlt = i % 2 === 1;

    for (let col = 1; col <= 4; col++) {
      const cell = row.getCell(col);
      cell.value = cat[col - 1];
      styleDataCell(cell, isAlt);
    }

    row.getCell(3).numFmt = currencyFormat;

    const typeCell = row.getCell(2);
    if (cat[1] === "Einnahme") {
      typeCell.font = { ...normalFont, bold: true, color: { argb: colors.green } };
    } else {
      typeCell.font = { ...normalFont, bold: true, color: { argb: colors.red } };
    }
  });

  // Empty rows for more categories
  for (let i = defaultCategories.length; i < 18; i++) {
    const row = categories_ws.getRow(3 + i);
    const isAlt = i % 2 === 1;
    for (let col = 1; col <= 4; col++) {
      const cell = row.getCell(col);
      styleDataCell(cell, isAlt);
      if (col === 3) cell.numFmt = currencyFormat;
    }
  }

  // Type validation
  for (let i = 3; i <= 20; i++) {
    categories_ws.getCell(`B${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ['"Einnahme,Ausgabe"'],
    };
  }

  categories_ws.getColumn(1).width = 22;
  categories_ws.getColumn(2).width = 14;
  categories_ws.getColumn(3).width = 16;
  categories_ws.getColumn(4).width = 12;

  // ============================================================
  // 4. BUDGETS Sheet
  // ============================================================
  const budgets = workbook.addWorksheet("Budgets", {
    properties: { tabColor: { argb: colors.purple } },
  });

  budgets.mergeCells("A1:F1");
  budgets.getCell("A1").value = "MONATLICHE BUDGETS";
  budgets.getCell("A1").font = { bold: true, size: 16, color: { argb: colors.primary }, name: "Calibri" };
  budgets.getRow(1).height = 35;

  const budgetHeaders2 = ["Kategorie", "Budget", "Ausgegeben", "Verbleibend", "Auslastung", "Status"];
  const budHRow = budgets.getRow(2);
  budgetHeaders2.forEach((h, i) => {
    budHRow.getCell(i + 1).value = h;
  });
  styleHeader(budHRow, 6);

  // Budget rows from expense categories
  const expenseCategories = defaultCategories.filter(c => c[1] === "Ausgabe");
  expenseCategories.forEach((cat, i) => {
    const row = budgets.getRow(3 + i);
    const isAlt = i % 2 === 1;
    const r = 3 + i;

    // Category
    const catCell = row.getCell(1);
    catCell.value = cat[0];
    styleDataCell(catCell, isAlt);
    catCell.font = { ...normalFont, bold: true };

    // Budget
    const budCell = row.getCell(2);
    budCell.value = cat[2];
    budCell.numFmt = currencyFormat;
    styleDataCell(budCell, isAlt);

    // Spent
    const spentCell = row.getCell(3);
    spentCell.value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$1000="Ausgabe")*(Transaktionen!D$3:D$1000=A${r})*(MONTH(Transaktionen!B$3:B$1000)=MONTH(TODAY()))*(YEAR(Transaktionen!B$3:B$1000)=YEAR(TODAY()))*Transaktionen!E$3:E$1000)` };
    spentCell.numFmt = currencyFormat;
    styleDataCell(spentCell, isAlt);
    spentCell.font = { ...normalFont, color: { argb: colors.red } };

    // Remaining
    const remCell = row.getCell(4);
    remCell.value = { formula: `B${r}-C${r}` };
    remCell.numFmt = currencyFormat;
    styleDataCell(remCell, isAlt);

    // Usage %
    const usageCell = row.getCell(5);
    usageCell.value = { formula: `IF(B${r}=0,0,C${r}/B${r})` };
    usageCell.numFmt = percentFormat;
    styleDataCell(usageCell, isAlt);

    // Status
    const statusCell = row.getCell(6);
    statusCell.value = { formula: `IF(B${r}=0,"—",IF(E${r}>1,"ÜBER BUDGET",IF(E${r}>0.8,"WARNUNG","OK")))` };
    styleDataCell(statusCell, isAlt);
    statusCell.alignment = { horizontal: "center", vertical: "middle" };
  });

  // Conditional formatting for status
  budgets.addConditionalFormatting({
    ref: "F3:F14",
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "ÜBER BUDGET",
        style: {
          font: { bold: true, color: { argb: colors.red } },
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: colors.redLight } },
        },
      },
      {
        type: "containsText",
        operator: "containsText",
        text: "WARNUNG",
        style: {
          font: { bold: true, color: { argb: colors.orange } },
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: colors.orangeLight } },
        },
      },
      {
        type: "containsText",
        operator: "containsText",
        text: "OK",
        style: {
          font: { bold: true, color: { argb: colors.green } },
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: colors.greenLight } },
        },
      },
    ],
  });

  // Conditional formatting for usage percentage (color scale)
  budgets.addConditionalFormatting({
    ref: "E3:E14",
    rules: [
      {
        type: "colorScale",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: 0.5 },
          { type: "num", value: 1 },
        ],
        color: [
          { argb: colors.greenLight },
          { argb: colors.yellowLight },
          { argb: colors.redLight },
        ],
      },
    ],
  });

  // Total row
  const budTotalRow = budgets.getRow(15);
  budTotalRow.height = 28;
  budTotalRow.getCell(1).value = "GESAMT";
  budTotalRow.getCell(1).font = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
  for (let c = 1; c <= 6; c++) {
    budTotalRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
    budTotalRow.getCell(c).border = thinBorder;
    budTotalRow.getCell(c).font = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
    budTotalRow.getCell(c).alignment = { vertical: "middle" };
  }
  budTotalRow.getCell(2).value = { formula: "SUM(B3:B14)" };
  budTotalRow.getCell(2).numFmt = currencyFormat;
  budTotalRow.getCell(3).value = { formula: "SUM(C3:C14)" };
  budTotalRow.getCell(3).numFmt = currencyFormat;
  budTotalRow.getCell(4).value = { formula: "B15-C15" };
  budTotalRow.getCell(4).numFmt = currencyFormat;
  budTotalRow.getCell(5).value = { formula: "IF(B15=0,0,C15/B15)" };
  budTotalRow.getCell(5).numFmt = percentFormat;

  budgets.getColumn(1).width = 20;
  budgets.getColumn(2).width = 16;
  budgets.getColumn(3).width = 16;
  budgets.getColumn(4).width = 16;
  budgets.getColumn(5).width = 14;
  budgets.getColumn(6).width = 16;

  // ============================================================
  // 5. KONTEN Sheet
  // ============================================================
  const konten = workbook.addWorksheet("Konten", {
    properties: { tabColor: { argb: "CA8A04" } },
  });

  konten.mergeCells("A1:E1");
  konten.getCell("A1").value = "KONTEN";
  konten.getCell("A1").font = { bold: true, size: 16, color: { argb: colors.primary }, name: "Calibri" };
  konten.getRow(1).height = 35;

  const accHeaders = ["Kontoname", "Typ", "Startguthaben", "Einnahmen", "Ausgaben", "Aktueller Stand"];
  const accHRow = konten.getRow(2);
  accHeaders.forEach((h, i) => {
    accHRow.getCell(i + 1).value = h;
  });
  styleHeader(accHRow, 6);

  const defaultAccounts = [
    ["Hauptkonto", "Bank", 2500],
    ["Bargeld", "Bargeld", 200],
    ["Kreditkarte", "Kreditkarte", 0],
    ["Sparkonto", "Sparkonto", 5000],
  ];

  defaultAccounts.forEach((acc, i) => {
    const row = konten.getRow(3 + i);
    const isAlt = i % 2 === 1;
    const r = 3 + i;

    row.getCell(1).value = acc[0];
    styleDataCell(row.getCell(1), isAlt);
    row.getCell(1).font = { ...normalFont, bold: true };

    row.getCell(2).value = acc[1];
    styleDataCell(row.getCell(2), isAlt);

    row.getCell(3).value = acc[2];
    row.getCell(3).numFmt = currencyFormat;
    styleDataCell(row.getCell(3), isAlt);

    // Income sum for this account
    const incCell = row.getCell(4);
    incCell.value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$1000="Einnahme")*(Transaktionen!F$3:F$1000=A${r})*Transaktionen!E$3:E$1000)` };
    incCell.numFmt = currencyFormat;
    styleDataCell(incCell, isAlt);
    incCell.font = { ...normalFont, color: { argb: colors.green } };

    // Expense sum for this account
    const expCell = row.getCell(5);
    expCell.value = { formula: `SUMPRODUCT((Transaktionen!C$3:C$1000="Ausgabe")*(Transaktionen!F$3:F$1000=A${r})*Transaktionen!E$3:E$1000)` };
    expCell.numFmt = currencyFormat;
    styleDataCell(expCell, isAlt);
    expCell.font = { ...normalFont, color: { argb: colors.red } };

    // Current balance
    const balCell = row.getCell(6);
    balCell.value = { formula: `C${r}+D${r}-E${r}` };
    balCell.numFmt = currencyFormat;
    styleDataCell(balCell, isAlt);
    balCell.font = { ...normalFont, bold: true, size: 12 };
  });

  // More empty account rows
  for (let i = defaultAccounts.length; i < 10; i++) {
    const row = konten.getRow(3 + i);
    const isAlt = i % 2 === 1;
    const r = 3 + i;
    for (let col = 1; col <= 6; col++) {
      const cell = row.getCell(col);
      styleDataCell(cell, isAlt);
      if (col >= 3) cell.numFmt = currencyFormat;
    }
    row.getCell(4).value = { formula: `IF(A${r}="","",SUMPRODUCT((Transaktionen!C$3:C$1000="Einnahme")*(Transaktionen!F$3:F$1000=A${r})*Transaktionen!E$3:E$1000))` };
    row.getCell(5).value = { formula: `IF(A${r}="","",SUMPRODUCT((Transaktionen!C$3:C$1000="Ausgabe")*(Transaktionen!F$3:F$1000=A${r})*Transaktionen!E$3:E$1000))` };
    row.getCell(6).value = { formula: `IF(A${r}="","",C${r}+D${r}-E${r})` };
  }

  // Account type validation
  for (let i = 3; i <= 12; i++) {
    konten.getCell(`B${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ['"Bank,Bargeld,Kreditkarte,Sparkonto"'],
    };
  }

  // Total row
  const accTotalRow = konten.getRow(15);
  accTotalRow.height = 28;
  accTotalRow.getCell(1).value = "GESAMT";
  for (let c = 1; c <= 6; c++) {
    accTotalRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerBg } };
    accTotalRow.getCell(c).border = thinBorder;
    accTotalRow.getCell(c).font = { bold: true, color: { argb: colors.headerText }, size: 11, name: "Calibri" };
    accTotalRow.getCell(c).alignment = { vertical: "middle" };
  }
  accTotalRow.getCell(3).value = { formula: "SUM(C3:C12)" };
  accTotalRow.getCell(3).numFmt = currencyFormat;
  accTotalRow.getCell(4).value = { formula: "SUM(D3:D12)" };
  accTotalRow.getCell(4).numFmt = currencyFormat;
  accTotalRow.getCell(5).value = { formula: "SUM(E3:E12)" };
  accTotalRow.getCell(5).numFmt = currencyFormat;
  accTotalRow.getCell(6).value = { formula: "SUM(F3:F12)" };
  accTotalRow.getCell(6).numFmt = currencyFormat;

  // Conditional formatting for balances
  konten.addConditionalFormatting({
    ref: "F3:F12",
    rules: [
      {
        type: "cellIs",
        operator: "lessThan",
        formulae: [0],
        style: {
          font: { bold: true, color: { argb: colors.red } },
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: colors.redLight } },
        },
      },
      {
        type: "cellIs",
        operator: "greaterThanOrEqual",
        formulae: [0],
        style: {
          font: { bold: true, color: { argb: colors.green } },
        },
      },
    ],
  });

  konten.getColumn(1).width = 20;
  konten.getColumn(2).width = 14;
  konten.getColumn(3).width = 18;
  konten.getColumn(4).width = 16;
  konten.getColumn(5).width = 16;
  konten.getColumn(6).width = 18;

  // ============================================================
  // 6. ANLEITUNG Sheet
  // ============================================================
  const help = workbook.addWorksheet("Anleitung", {
    properties: { tabColor: { argb: colors.gray } },
  });

  help.getColumn(1).width = 80;
  help.getRow(1).height = 40;
  help.mergeCells("A1:A1");
  help.getCell("A1").value = "BUDGET TRACKER - ANLEITUNG";
  help.getCell("A1").font = { bold: true, size: 18, color: { argb: colors.primary }, name: "Calibri" };
  help.getCell("A1").alignment = { vertical: "middle" };

  const instructions = [
    "",
    "SO BENUTZT DU DEN BUDGET TRACKER:",
    "",
    "1. TRANSAKTIONEN EINTRAGEN",
    "   → Gehe zum Tab 'Transaktionen'",
    "   → Trage jede Einnahme und Ausgabe ein",
    "   → Wähle Typ (Einnahme/Ausgabe), Kategorie und Konto aus den Dropdown-Listen",
    "   → Das Dashboard aktualisiert sich automatisch!",
    "",
    "2. KATEGORIEN ANPASSEN",
    "   → Im Tab 'Kategorien' kannst du eigene Kategorien hinzufügen",
    "   → Setze Monatsbudgets für jede Ausgabe-Kategorie",
    "",
    "3. BUDGETS VERWALTEN",
    "   → Der Tab 'Budgets' zeigt dir automatisch wie viel du pro Kategorie ausgegeben hast",
    "   → Status: OK (grün) = unter 80%, WARNUNG (orange) = 80-100%, ÜBER BUDGET (rot) = über 100%",
    "",
    "4. KONTEN",
    "   → Verwalte mehrere Konten (Bank, Bargeld, Kreditkarte, Sparkonto)",
    "   → Der aktuelle Stand wird automatisch berechnet (Startguthaben + Einnahmen - Ausgaben)",
    "",
    "5. DASHBOARD",
    "   → Das Dashboard zeigt dir alle wichtigen Kennzahlen auf einen Blick",
    "   → Gesamtguthaben, Monatseinnahmen/-ausgaben, Bilanz",
    "   → Monatsübersicht mit Sparquote",
    "   → Budget-Status für den aktuellen Monat",
    "",
    "TIPPS:",
    "   → Nutze die Filter im Tab 'Transaktionen' (Autofilter in der Kopfzeile)",
    "   → Trage Transaktionen regelmäßig ein für eine genaue Übersicht",
    "   → Passe die Budgets monatlich an deine Bedürfnisse an",
    "   → Du kannst die Beispieldaten löschen und mit deinen eigenen starten",
  ];

  instructions.forEach((text, i) => {
    const row = help.getRow(2 + i);
    const cell = row.getCell(1);
    cell.value = text;
    if (text.startsWith("SO BENUTZT") || text === "TIPPS:") {
      cell.font = { bold: true, size: 13, color: { argb: colors.dark }, name: "Calibri" };
    } else if (text.match(/^\d\./)) {
      cell.font = { bold: true, size: 12, color: { argb: colors.primary }, name: "Calibri" };
    } else if (text.startsWith("   →")) {
      cell.font = { size: 11, color: { argb: colors.gray }, name: "Calibri" };
    } else {
      cell.font = normalFont;
    }
  });

  // ============================================================
  // Protect sheets (allow editing data cells only)
  // ============================================================
  // Dashboard is read-only (formulas only)
  dashboard.protect("", {
    selectLockedCells: true,
    selectUnlockedCells: true,
  });

  // Set the active sheet to Dashboard
  workbook.views = [{ activeTab: 0 }];

  // Freeze panes
  transactions.views = [{ state: "frozen", ySplit: 2, activeCell: "A3" }];
  categories_ws.views = [{ state: "frozen", ySplit: 2, activeCell: "A3" }];
  budgets.views = [{ state: "frozen", ySplit: 2, activeCell: "A3" }];
  konten.views = [{ state: "frozen", ySplit: 2, activeCell: "A3" }];

  // ============================================================
  // Save
  // ============================================================
  const outputPath = path.join(__dirname, "Budget-Tracker.xlsx");
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel-Datei erstellt: ${outputPath}`);
}

generateBudgetTracker().catch(console.error);
