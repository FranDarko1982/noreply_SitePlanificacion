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
    const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    console.error(JSON.stringify({
      message: 'Error en isUserAuthorized',
      user: email,
      error: e.message,
      stack: e.stack
    }, null, 2));
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
    const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Insights');
    if (sheet) {
      sheet.appendRow([userEmail, timestamp]);
    }
  } catch (error) {
    console.error(JSON.stringify({
      message: 'Error al registrar acceso en doGet',
      user: userEmail,
      error: error.message,
      stack: error.stack
    }, null, 2));
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

function getInsightsData() {
  Logger.log('---------- getInsightsData() started ----------');
  const userEmail = getActiveUserEmail();
  const authorized = isUserAuthorized(userEmail);
  Logger.log(`Authorization check for Insights - user: ${userEmail}, authorized: ${authorized}`);
  if (!authorized) {
    throw new Error('ACCESO_NO_AUTORIZADO');
  }
  try {
    const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    Logger.log(`SPREADSHEET_ID: ${SPREADSHEET_ID}`);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Insights');
    if (!sheet) {
      throw new Error('La hoja "Insights" no existe.');
    }
    Logger.log('Sheet "Insights" opened successfully.');

    const data = sheet.getDataRange().getValues();
    Logger.log(`Raw data from sheet (first 5 rows): ${JSON.stringify(data.slice(0, 5))}`);
    // Omitir la fila de la cabecera
    const records = data.slice(1).map(row => ({ email: row[0], timestamp: new Date(row[1]) }));
    Logger.log(`Processed records (first 5): ${JSON.stringify(records.slice(0, 5))}`);
    Logger.log(`Total records: ${records.length}`);


    // --- CÁLCULOS ---
    Logger.log('Starting calculations...');

    // 1. Accesos por día (últimos 7 días)
    const accessesPerDay = {
      labels: [],
      values: []
    };
    const days = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    Logger.log('Calculating accesses per day...');

    for (let i = 6; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(today.getDate() - i);
      const dayName = days[date.getDay()];
      accessesPerDay.labels.push(dayName);

      const count = records.filter(r => {
        const recordDate = new Date(r.timestamp);
        return recordDate >= date && recordDate < new Date(date.getTime() + 24 * 60 * 60 * 1000);
      }).length;
      accessesPerDay.values.push(count);
    }
    Logger.log(`Accesses per day: ${JSON.stringify(accessesPerDay)}`);


    // 2. Usuarios únicos por semana (últimas 4 semanas)
    const uniqueUsersPerPeriod = {
      labels: [],
      values: []
    };
    Logger.log('Calculating unique users per period...');
    for (let i = 3; i >= 0; i--) {
      const weekStartDate = new Date(today);
      weekStartDate.setDate(today.getDate() - today.getDay() - (i * 7)); // Inicio de la semana (domingo)
      const weekEndDate = new Date(weekStartDate);
      weekEndDate.setDate(weekStartDate.getDate() + 7);

      const uniqueUsers = new Set(
        records
          .filter(r => r.timestamp >= weekStartDate && r.timestamp < weekEndDate)
          .map(r => r.email)
      );
      uniqueUsersPerPeriod.labels.push(`Semana ${4 - i}`);
      uniqueUsersPerPeriod.values.push(uniqueUsers.size);
    }
    Logger.log(`Unique users per period: ${JSON.stringify(uniqueUsersPerPeriod)}`);


    // 3. Nº de accesos por usuario
    Logger.log('Calculating accesses per user...');
    const accessesByUser = records.reduce((acc, record) => {
      acc[record.email] = (acc[record.email] || 0) + 1;
      return acc;
    }, {});

    const accessesPerUser = {
      labels: ['1 acceso', '2-3 accesos', '4-7 accesos', '8+ accesos'],
      values: [0, 0, 0, 0]
    };
    Object.values(accessesByUser).forEach(count => {
      if (count === 1) accessesPerUser.values[0]++;
      else if (count >= 2 && count <= 3) accessesPerUser.values[1]++;
      else if (count >= 4 && count <= 7) accessesPerUser.values[2]++;
      else if (count >= 8) accessesPerUser.values[3]++;
    });
    Logger.log(`Accesses per user: ${JSON.stringify(accessesPerUser)}`);

    // 4. Heatmap (Día x Hora)
    Logger.log('Calculating heatmap data...');
    const heatmapData = days.map(day => ({ day, values: Array(24).fill(0) }));
    records.forEach(record => {
      const dayIndex = record.timestamp.getDay();
      const hour = record.timestamp.getHours();
      heatmapData[dayIndex].values[hour]++;
    });
     // Reordenar para que empiece en Lunes
    const reorderedHeatmapData = [];
    for (let i = 1; i < days.length; i++) {
      reorderedHeatmapData.push(heatmapData[i]);
    }
    reorderedHeatmapData.push(heatmapData[0]);
    Logger.log(`Heatmap data: ${JSON.stringify(reorderedHeatmapData)}`);


    // 5. Resumen
    Logger.log('Calculating summary data...');
    const allUsers = Object.keys(accessesByUser);
    const totalUsers = allUsers.length;

    const firstAccess = records.reduce((acc, record) => {
      if (!acc[record.email] || record.timestamp < acc[record.email]) {
        acc[record.email] = record.timestamp;
      }
      return acc;
    }, {});

    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(today.getDate() - 30);
    const newUsers = allUsers.filter(email => firstAccess[email] >= thirtyDaysAgo).length;
    const recurrentUsers = totalUsers - newUsers;

    const totalAccesses = records.length;
    const avgAccessesPerUser = totalUsers > 0 ? totalAccesses / totalUsers : 0;

    const accessesByDayOfWeek = Array(7).fill(0);
    records.forEach(r => accessesByDayOfWeek[r.timestamp.getDay()]++);
    const busiestDayIndex = accessesByDayOfWeek.indexOf(Math.max(...accessesByDayOfWeek));
    const busiestDay = days[busiestDayIndex];

    const summary = {
      totalUsers,
      newUsers,
      recurrentUsers,
      avgAccessesPerUser,
      busiestDay
    };
    Logger.log(`Summary data: ${JSON.stringify(summary)}`);

    Logger.log('---------- getInsightsData() finished successfully ----------');
    return {
      accessesPerDay,
      uniqueUsersPerPeriod,
      accessesPerUser,
      heatmapData: reorderedHeatmapData,
      summary
    };

  } catch (e) {
    Logger.log('---------- getInsightsData() failed ----------');
    Logger.error(JSON.stringify({
      message: 'Error en getInsightsData',
      error: e.message,
      stack: e.stack,
      fullError: e.toString()
    }, null, 2));
    // En caso de error, devolver una estructura vacía para no romper el frontend
    return {
      accessesPerDay: { labels: [], values: [] },
      uniqueUsersPerPeriod: { labels: [], values: [] },
      accessesPerUser: { labels: [], values: [] },
      heatmapData: [],
      summary: { totalUsers: 0, newUsers: 0, recurrentUsers: 0, avgAccessesPerUser: 0, busiestDay: '-' }
    };
  }
}
