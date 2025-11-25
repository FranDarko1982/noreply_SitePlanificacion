/**
 * @OnlyCurrentDoc
 */

// --- FUNCIONES DE AUTORIZACIÓN Y DATOS ---

/**
 * Comprueba si un correo electrónico está en la lista de autorizados de la hoja "usuarios".
 * @param {string} email El correo del usuario a comprobar.
 * @returns {boolean} Devuelve true si el usuario está autorizado, de lo contrario false.
 */
function isUserAuthorized(email) {
  if (!email || email === 'Acceso Anónimo') {
    return false;
  }
  try {
    const spreadsheet = SpreadsheetApp.openById('1cSkutacmPTEReg1RErr0dAzllF40CcwKGe-9Blfq0KA');
    const sheet = spreadsheet.getSheetByName('usuarios');
    if (!sheet) {
      console.error('La hoja "usuarios" no existe.');
      return false;
    }
    // Obtiene todos los correos de la columna A
    const authorizedEmails = sheet.getRange('A2:A').getValues()
      .flat() // Convierte el array 2D en 1D
      .filter(String) // Elimina celdas vacías
      .map(e => e.trim().toLowerCase()); // Limpia y normaliza los correos

    return authorizedEmails.includes(email.trim().toLowerCase());
  } catch (e) {
    console.error('Error al verificar la autorización: ' + e.toString());
    return false;
  }
}

/**
 * Devuelve el correo del usuario activo.
 */
function getActiveUserEmail() {
  const activeUser = Session.getActiveUser();
  return activeUser ? activeUser.getEmail() : 'Acceso Anónimo';
}

// --- FUNCIONES PRINCIPALES DE LA APLICACIÓN WEB ---

function doGet(e) {
  const userEmail = getActiveUserEmail();
  const isAuthorized = isUserAuthorized(userEmail);

  // Registrar el acceso del usuario (siempre se ejecuta)
  try {
    const timestamp = new Date();
    const spreadsheet = SpreadsheetApp.openById('1cSkutacmPTEReg1RErr0dAzllF40CcwKGe-9Blfq0KA');
    const sheet = spreadsheet.getSheetByName('Insights');
    if (sheet) {
      sheet.appendRow([userEmail, timestamp]);
    }
  } catch (error) {
    console.error('Error al registrar el acceso: ' + error.toString());
  }

  // Preparar la plantilla HTML y pasarle las variables
  const template = HtmlService.createTemplateFromFile('index');
  template.isAuthorized = isAuthorized; // Aquí pasamos la variable a la plantilla

  // Evaluar y devolver el HTML
  return template.evaluate()
      .setTitle('Planificación de Proyectos')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Permite incluir el contenido de otros ficheros HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
