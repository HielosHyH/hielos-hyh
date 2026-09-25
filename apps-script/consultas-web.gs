/**
 * Consultas de la página web de Hielos HyH → Google Sheets.
 *
 * Se pega en la planilla (Extensiones → Apps Script) y se publica como
 * "Aplicación web". Cada mensaje del formulario queda como una fila nueva en la
 * hoja "Consultas" y además llega un aviso por correo.
 */

// Correo que recibe el aviso de cada consulta nueva
const CORREO_AVISO = 'hieloshyh@gmail.com';
const NOMBRE_HOJA = 'Consultas';
const COLUMNAS = ['Fecha', 'Nombre', 'Correo', 'Teléfono', 'Mensaje', 'Acepta novedades'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hoja = obtenerHoja_();
    hoja.appendRow([
      new Date(),
      limpiar_(p['Nombre']),
      limpiar_(p['Email']),
      limpiar_(p['Teléfono']),
      limpiar_(p['Mensaje']),
      p['Acepta novedades'] === 'Sí' ? 'Sí' : 'No',
    ]);
  } finally {
    lock.releaseLock();
  }

  MailApp.sendEmail({
    to: CORREO_AVISO,
    replyTo: p['Email'] || CORREO_AVISO,
    subject: 'Nueva consulta desde la web: ' + (p['Nombre'] || 'sin nombre'),
    body:
      'Nombre: ' + (p['Nombre'] || '') + '\n' +
      'Correo: ' + (p['Email'] || '') + '\n' +
      'Teléfono: ' + (p['Teléfono'] || '') + '\n' +
      'Acepta novedades: ' + (p['Acepta novedades'] === 'Sí' ? 'Sí' : 'No') + '\n\n' +
      'Mensaje:\n' + (p['Mensaje'] || ''),
  });

  return ContentService.createTextOutput('ok');
}

// Crea la hoja con sus títulos la primera vez
function obtenerHoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(COLUMNAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Evita que un texto que empiece con =, +, - o @ se interprete como fórmula
function limpiar_(valor) {
  const texto = String(valor || '').trim().slice(0, 5000);
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}

// Para probar desde el editor: Ejecutar → pruebaManual
function pruebaManual() {
  doPost({ parameter: {
    'Nombre': 'Prueba', 'Email': CORREO_AVISO, 'Teléfono': '',
    'Mensaje': 'Mensaje de prueba', 'Acepta novedades': 'Sí',
  } });
}
