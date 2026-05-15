/**
 * פונקציה להצגת דף האינטרנט
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('מערכת ניהול ייצור')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * פונקציה למציאת גיליון בצורה גמישה (מתעלמת מרווחים)
 */
function getSheetFlexible(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const trimmedName = name.trim();

  // ניסיון ראשון: שם מדויק
  let sheet = ss.getSheetByName(trimmedName);
  if (sheet) return sheet;

  // ניסיון שני: חיפוש ללא רווחים בהתחלה ובסוף
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === trimmedName) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * פונקציית תיקון עמוקה - מנקה חסימות ומזריקה נוסחאות
 */
function resetSheetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = getSheetFlexible('תוכנית עבודה');
  const printSheet = getSheetFlexible('הדפסה');

  if (!workSheet || !printSheet) {
    const allNames = ss.getSheets().map(s => '"' + s.getName() + '"').join(", ");
    return "שגיאה: לא נמצאו הגיליונות 'תוכנית עבודה' או 'הדפסה'. הגיליונות שקיימים כרגע הם: " + allNames;
  }

  try {
    // 1. ניקוי אגרסיבי של טורים B, D, E בגיליון תוכנית עבודה
    // זה קריטי כדי ש-ARRAYFORMULA תוכל להתרחב
    workSheet.getRange("B2:B150").clearContent();
    workSheet.getRange("D2:D150").clearContent();
    workSheet.getRange("E2:E150").clearContent();

    // 2. הזרקת נוסחאות ה-ARRAYFORMULA (עד שורה 150)
    workSheet.getRange("B2").setFormula('=ARRAYFORMULA(IF(A2:A150="", "", IFERROR(XLOOKUP(A2:A150, \'בסיס נתונים\'!$A$2:$A$150, \'בסיס נתונים\'!$B$2:$B$150), "מק""ט לא נמצא")))');
    workSheet.getRange("D2").setFormula('=ARRAYFORMULA(IF(A2:A150="", "", IFERROR(XLOOKUP(A2:A150, \'בסיס נתונים\'!$A$2:$A$150, \'בסיס נתונים\'!$D$2:$D$150), 0)))');
    workSheet.getRange("E2").setFormula('=ARRAYFORMULA(IF(A2:A150="", "", IF(D2:D150>0, C2:C150/D2:D150, 0)))');

    // 3. עדכון גיליון ההדפסה
    printSheet.getRange("A2:E150").clearContent();
    printSheet.getRange("A2").setFormula('=IFERROR(FILTER(\'תוכנית עבודה\'!A2:E150, \'תוכנית עבודה\'!C2:C150 > 0), "ממתין לנתונים...")');

    return "בוצע ניקוי עמוק והנוסחאות הוגדרו מחדש עד שורה 150! ✅";
  } catch (e) {
    return "שגיאה בתהליך האיפוס: " + e.message;
  }
}

/**
 * שליפת נתונים לממשק
 */
function getCategorizedData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = getSheetFlexible('בסיס נתונים');
    const printSheet = getSheetFlexible('הדפסה');

    if (!printSheet) throw new Error("גיליון הדפסה לא נמצא");

    const categoryMap = {};
    if (dbSheet) {
      const dbData = dbSheet.getRange("A2:C150").getValues();
      dbData.forEach(row => {
        if (row[0]) categoryMap[row[0].toString().trim()] = row[2] || "כללי";
      });
    }

    const data = printSheet.getRange("A2:E150").getValues();
    const groupedData = {};

    data.forEach(row => {
      const sku = row[0] ? row[0].toString().trim() : "";
      if (!sku || sku === "" || sku === "ממתין לנתונים...") return;

      const category = categoryMap[sku] || "כללי";
      if (!groupedData[category]) groupedData[category] = [];

      groupedData[category].push({
        sku: sku,
        desc: row[1] || "",
        qty: row[2] || 0,
        perCart: row[3] || 0,
        totalCarts: parseFloat(row[4] || 0)
      });
    });

    return groupedData;
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * שמירת שינויים חזרה ל'תוכנית עבודה'
 */
function saveEditsToSheet(flatData) {
  const workSheet = getSheetFlexible('תוכנית עבודה');
  if (!workSheet) return "שגיאה: גיליון תוכנית עבודה לא נמצא";

  // מנקים רק עמודות קלט A ו-C
  workSheet.getRange("A2:A150").clearContent();
  workSheet.getRange("C2:C150").clearContent();

  const skus = flatData.map(i => [i.sku]);
  const qtys = flatData.map(i => [i.qty]);

  if (skus.length > 0) {
    workSheet.getRange(2, 1, skus.length, 1).setValues(skus);
    workSheet.getRange(2, 3, qtys.length, 1).setValues(qtys);
  }
  return "השינויים נשמרו בשיט ✅";
}

/**
 * תיעוד בארכיון ודוח הרלס
 */
function archiveProduction(data, dateStr, harlessText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let archiveSheet = getSheetFlexible('ארכיון ייצור') || ss.insertSheet('ארכיון ייצור');

  data.forEach(row => {
    archiveSheet.appendRow([dateStr, row.sku, row.desc, row.qty, row.category, Math.round(row.totalCarts)]);
  });

  let harlessSheet = getSheetFlexible('תוכנית יומית להרלס') || ss.insertSheet('תוכנית יומית להרלס');
  harlessSheet.appendRow([dateStr, harlessText]);

  return "הנתונים תועדו בארכיון ✅";
}
