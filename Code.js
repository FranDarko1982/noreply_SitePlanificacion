/** ==========
 *  Code.gs
 *  ========== */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('index');
  return tpl
    .evaluate()
    .setTitle('Planificación Logística y Teletrabajo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Permite incluir archivos parciales (styles, scripts, secciones)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}