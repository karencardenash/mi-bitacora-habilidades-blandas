function doPost(request) {
  var spreadsheet = SpreadsheetApp.openById('1c0VldiG34F_w8Y3Z33k7ASz-AiM8r1HD18wHkx15PwE');
  var sheet = spreadsheet.getSheetByName('Respuestas');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Respuestas');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Fecha y hora',
      'ID de sesion',
      'Evento',
      'Nombre',
      'Programa',
      'Semestre',
      'XP',
      'Niveles completados',
      'Habilidad',
      'Pregunta',
      'Respuesta',
      'Estado de animo',
      'Puntajes',
      'Diagnostico',
      'Dispositivo o navegador',
      'Pagina',
      'Detalles'
    ]);
  }

  var data = JSON.parse(request.postData.contents);
  var details = data.details || {};

  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.sessionId || '',
    data.event || '',
    data.name || '',
    data.program || '',
    data.semester || '',
    data.xp || 0,
    data.progress || 0,
    details.skill || '',
    details.prompt || '',
    details.response || JSON.stringify(details.answers || ''),
    details.mood || '',
    JSON.stringify(data.scores || {}),
    JSON.stringify(data.answers || []),
    data.userAgent || '',
    data.page || '',
    JSON.stringify(details)
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'Mi Bitacora' }))
    .setMimeType(ContentService.MimeType.JSON);
}
