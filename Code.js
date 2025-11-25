/** ==========
 *  Code.gs
 *  ========== */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.isAuthorized = checkUserAuthorization();
  return tpl
    .evaluate()
    .setTitle('Planificación Logística y Teletrabajo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Permite incluir archivos parciales (styles, scripts, secciones)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Verifica si el usuario actual está autorizado
function checkUserAuthorization() {
  const userEmail = Session.getActiveUser().getEmail();
  const sheetId = '1cSkutacmPTEReg1RErr0dAzllF40CcwKGe-9Blfq0KA';
  const sheetName = 'Usuarios';
  
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) return false;
    
    const range = sheet.getRange("A:A");
    const values = range.getValues();
    const emails = values.flat().filter(String).map(e => e.toLowerCase());
    
    return emails.includes(userEmail.toLowerCase());
  } catch (e) {
    Logger.log("Error checking authorization: " + e.toString());
    return false;
  }
}