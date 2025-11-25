/**
 * @OnlyCurrentDoc
 */

/**
 * Devuelve el correo del usuario activo.
 * Si el script no puede identificar al usuario (por permisos o acceso anónimo),
 * devuelve 'Acceso Anónimo'.
 */
function getActiveUserEmail() {
  const activeUser = Session.getActiveUser();
  return activeUser ? activeUser.getEmail() : 'Acceso Anónimo';
}

function doGet(e) {
  // Registrar el acceso del usuario
  try {
    const userEmail = getActiveUserEmail();
    const timestamp = new Date();
    const spreadsheet = SpreadsheetApp.openById('1cSkutacmPTEReg1RErr0dAzllF40CcwKGe-9Blfq0KA');
    const sheetName = 'Insights';
    let sheet = spreadsheet.getSheetByName(sheetName);

    // Si la hoja 'Insights' no existe, la crea y añade cabeceras.
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.appendRow(['Correo Electrónico', 'Fecha y Hora']);
      sheet.getRange('A1:B1').setFontWeight('bold');
    }

    // Añade una nueva fila con la información del acceso.
    sheet.appendRow([userEmail, timestamp]);

  } catch (error) {
    // Si hay un error al registrar, lo anota en el log para depuración,
    // pero no impide que la app se cargue.
    console.error('Error al registrar el acceso: ' + error.toString());
  }

  // Sirve la página de inicio
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Planificación de Proyectos')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Permite incluir el contenido de otros ficheros HTML (como CSS o JS) dentro de una plantilla.
 * En tu HTML principal, puedes usar: <?!= include('styles'); ?> o <?!= include('scripts'); ?>
 *
 * @param {string} filename El nombre del fichero HTML a incluir (sin la extensión .html).
 * @return {string} El contenido del fichero.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}