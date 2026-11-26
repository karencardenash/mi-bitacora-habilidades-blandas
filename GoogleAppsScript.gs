function doPost(request) {
  var spreadsheet = SpreadsheetApp.openById('1c0VldiG34F_w8Y3Z33k7ASz-AiM8r1HD18wHkx15PwE');
  var sheet = spreadsheet.getSheetByName('Respuestas');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Respuestas');
  }

  var skillHeaders = [
    'Porcentaje: Comunicación de ideas',
    'Porcentaje: Gestión del tiempo',
    'Porcentaje: Adaptabilidad y presión',
    'Porcentaje: Trabajo en equipo',
    'Porcentaje: Pensamiento crítico y creativo',
    'Porcentaje: Recibir y dar feedback',
    'Porcentaje: Empatía y escucha',
    'Porcentaje: Toma de decisiones',
    'Porcentaje: Confianza y motivación',
    'Porcentaje: Networking y marca personal'
  ];
  var baseHeaders = [
    'Fecha y hora', 'ID de sesion', 'Evento', 'Nombre', 'Programa', 'Semestre',
    'XP', 'Niveles completados', 'Habilidad', 'Pregunta', 'Respuesta',
    'Estado de animo', 'Puntajes', 'Diagnostico', 'Dispositivo o navegador',
    'Pagina', 'Detalles'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(baseHeaders.concat(skillHeaders).concat(['Correo electronico']));
  } else {
    if (sheet.getRange(1, 18, 1, 10).getValues()[0].every(function(value) { return value === ''; })) {
      sheet.getRange(1, 18, 1, 10).setValues([skillHeaders]);
    }
    if (sheet.getRange(1, 28).getValue() === '') {
      sheet.getRange(1, 28).setValue('Correo electronico');
    }
  }

  var data = JSON.parse(request.postData.contents);
  var details = data.details || {};
  var percentages = data.skillPercentages || {};
  var skillIds = ['com', 'tiempo', 'adapt', 'equipo', 'critico', 'feedback', 'empatia', 'decis', 'conf', 'marca'];
  var skillValues = skillIds.map(function(id) {
    return percentages[id] === undefined ? '' : Number(percentages[id]) + '%';
  });

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
  ].concat(skillValues).concat([data.email || '']));

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'Mi Bitacora' }))
    .setMimeType(ContentService.MimeType.JSON);
}
