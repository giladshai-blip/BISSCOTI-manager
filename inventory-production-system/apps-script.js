/**
 * Google Apps Script for Inventory Production Planning System
 * מערכת תוכנית ייצור על סמך מלאי מינימום ומקסימום
 */

// Initialize sheets on open
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 תוכנית ייצור')
    .addItem('📥 יבא דוח מלאי', 'showImportDialog')
    .addItem('🔄 עדכן חישובים', 'calculateProduction')
    .addItem('🖨️ הדפס תוכנית', 'printProductionPlan')
    .addToUi();
}

// Create initial sheets structure
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create "דוח מלאי" sheet
  let inventorySheet = ss.getSheetByName('דוח מלאי');
  if (!inventorySheet) {
    inventorySheet = ss.insertSheet('דוח מלאי', 0);
    setupInventorySheet(inventorySheet);
  }

  // Create "תוכנית ייצור" sheet
  let productionSheet = ss.getSheetByName('תוכנית ייצור');
  if (!productionSheet) {
    productionSheet = ss.insertSheet('תוכנית ייצור', 1);
    setupProductionSheet(productionSheet);
  }
}

// Setup inventory sheet with headers
function setupInventorySheet(sheet) {
  const headers = ['קוד', 'שם מוצר', 'מלאי קיים', 'מינימום', 'מקסימום', 'צריך לייצר'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  // Set column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
}

// Setup production plan sheet
function setupProductionSheet(sheet) {
  // Add title
  sheet.getRange(1, 1).setValue('תוכנית ייצור שבועית');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');

  // Add headers
  const headers = ['מוצר', 'ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'סה״כ שבוע'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);

  // Format header
  const headerRange = sheet.getRange(3, 1, 1, headers.length);
  headerRange.setBackground('#34A853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  // Set column widths
  sheet.setColumnWidth(1, 150);
  for (let i = 2; i <= 8; i++) {
    sheet.setColumnWidth(i, 100);
  }
}

// Calculate production quantities
function calculateProduction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName('דוח מלאי');
  const productionSheet = ss.getSheetByName('תוכנית ייצור');

  if (!inventorySheet || !productionSheet) {
    SpreadsheetApp.getUi().alert('שגיאה: לא נמצאו גיליונות נדרשים');
    return;
  }

  // Get inventory data
  const data = inventorySheet.getDataRange().getValues();

  // Clear old production plan (keep header)
  if (productionSheet.getLastRow() > 3) {
    productionSheet.deleteRows(4, productionSheet.getLastRow() - 3);
  }

  // Add products to production plan
  let rowIndex = 4;
  for (let i = 1; i < data.length; i++) {
    const code = data[i][0];
    const name = data[i][1];
    const current = data[i][2];
    const min = data[i][3];
    const max = data[i][4];
    const needed = data[i][5];

    if (needed > 0) {
      productionSheet.getRange(rowIndex, 1).setValue(name + ' (' + needed + ')');

      // Add sum formula for the week
      const sumFormula = `=SUM(B${rowIndex}:G${rowIndex})`;
      productionSheet.getRange(rowIndex, 8).setFormula(sumFormula);

      rowIndex++;
    }
  }

  // Format the new rows
  if (rowIndex > 4) {
    const dataRange = productionSheet.getRange(4, 1, rowIndex - 4, 8);
    dataRange.setBorder(true, true, true, true, true, true);
    dataRange.setHorizontalAlignment('center');
    dataRange.setVerticalAlignment('middle');
  }

  SpreadsheetApp.getUi().alert('✅ תוכנית הייצור עודכנה בהצלחה!');
}

// Show import dialog
function showImportDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      textarea { width: 100%; height: 200px; font-size: 12px; }
      button {
        background-color: #4285F4;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        margin-top: 10px;
      }
      button:hover { background-color: #2c63c4; }
    </style>
    <h2>📥 יבא דוח מלאי</h2>
    <p>דביק את דוח המלאי שלך בעלון (CSV או טאב מופרד):</p>
    <textarea id="data"></textarea>
    <button onclick="importData()">יבא נתונים</button>

    <script>
      function importData() {
        const data = document.getElementById('data').value;
        google.script.run.importInventoryData(data);
        google.script.host.close();
      }
    </script>
  `);
  SpreadsheetApp.getUi().showModelessDialog(html, '📥 יבא דוח מלאי');
}

// Import inventory data
function importInventoryData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דוח מלאי');

  // Parse data
  const lines = data.trim().split('\n');
  const values = [];

  for (let line of lines) {
    const cells = line.split(/\t|,/);
    values.push(cells);
  }

  // Clear old data
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  // Insert new data
  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }

  // Add formulas for "צריך לייצר" column
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    const formula = `=IF(AND(C${i}<>"", D${i}<>""), MAX(0, E${i}-C${i}), "")`;
    sheet.getRange(i, 6).setFormula(formula);
  }

  SpreadsheetApp.getUi().alert('✅ נתונים יובאו בהצלחה!');
  calculateProduction();
}

// Print production plan
function printProductionPlan() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productionSheet = ss.getSheetByName('תוכנית ייצור');

  if (productionSheet) {
    // Set print options
    const printSettings = ss.getSpreadsheetTheme();
    SpreadsheetApp.getUi().alert('🖨️ עכשיו תוכל להדפיס את תוכנית הייצור (Ctrl+P או Cmd+P)');
  }
}
