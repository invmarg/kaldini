/**
 * Quiz Navideño Kaldini 2026 — receptor de participaciones
 * Guarda cada registro en Google Sheets y la foto de la factura en Google Drive.
 *
 * 1. Ejecuta  setup()  una sola vez (crea la hoja y la carpeta de fotos).
 * 2. Implementa como Aplicación web: Ejecutar como "Yo", Acceso "Cualquier usuario".
 * 3. Copia la URL que termina en /exec y pégala en index.html (CONFIG.ENDPOINT).
 */

const NOMBRE_HOJA = 'Participaciones';
const ENCABEZADOS = [
  'Código', 'Fecha', 'Nombre y apellido', 'Cédula', 'Teléfono / WhatsApp',
  'Supermercado', 'Productos probados', 'Favorito', 'Calificación (1-5)',
  'Número(s) de factura', 'Foto de factura', 'Factura repetida'
];

/* ---------- Configuración inicial (ejecutar una vez) ---------- */
function setup() {
  const props = PropertiesService.getScriptProperties();

  let ss;
  const sheetId = props.getProperty('SHEET_ID');
  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.create('Quiz Navideño Kaldini 2026 — Participaciones');
    props.setProperty('SHEET_ID', ss.getId());
  }

  let sh = ss.getSheetByName(NOMBRE_HOJA);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(NOMBRE_HOJA);
  }
  sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS])
    .setFontWeight('bold').setBackground('#0b4fd6').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, ENCABEZADOS.length, 170);

  if (!props.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('Quiz Kaldini 2026 — Fotos de facturas');
    props.setProperty('FOLDER_ID', folder.getId());
  }

  Logger.log('Hoja lista: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
}

/* ---------- Recibe cada participación ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);

    // Antispam: campo oculto que solo llenan los bots
    if (d.website) return responder({ ok: true, codigo: 'KAL-0000' });

    const req = ['nombre', 'cedula', 'telefono', 'favorito', 'factura', 'foto'];
    for (const k of req) {
      if (!d[k] || String(d[k]).trim() === '') return responder({ ok: false, error: 'Falta: ' + k });
    }
    const calif = Number(d.calificacion);
    if (!(calif >= 1 && calif <= 5)) return responder({ ok: false, error: 'Calificación inválida' });

    const props = PropertiesService.getScriptProperties();
    const sh = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName(NOMBRE_HOJA);
    const folder = DriveApp.getFolderById(props.getProperty('FOLDER_ID'));

    lock.waitLock(20000);

    const fila = sh.getLastRow();            // fila 1 = encabezados
    const codigo = 'KAL-' + String(fila).padStart(4, '0');

    // Foto
    const m = String(d.foto).match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
    if (!m) return responder({ ok: false, error: 'Foto inválida' });
    const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
    const archivo = folder.createFile(
      Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], codigo + ' - ' + limpiar(d.nombre) + '.' + ext)
    );

    // ¿La factura ya se registró antes?
    const facturas = String(d.factura).split(/[,;\s]+/).filter(String);
    let repetida = 'No';
    if (fila > 1) {
      const previas = sh.getRange(2, 10, fila - 1, 1).getValues().flat().join(',').split(/[,;\s]+/);
      if (facturas.some(f => previas.indexOf(f) !== -1)) repetida = 'Sí';
    }

    sh.appendRow([
      codigo,
      new Date(),
      limpiar(d.nombre),
      limpiar(d.cedula),
      limpiar(d.telefono),
      limpiar(d.supermercado),
      limpiar(d.probados),
      limpiar(d.favorito),
      calif,
      limpiar(d.factura),
      archivo.getUrl(),
      repetida
    ]);

    return responder({ ok: true, codigo: codigo });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* Para comprobar que la URL funciona desde el navegador */
function doGet() {
  return responder({ ok: true, mensaje: 'Quiz Kaldini activo' });
}

/* ---------- Sorteo: ejecuta y mira el registro (Ver > Registros) ---------- */
function sortearGanador() {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))
    .getSheetByName(NOMBRE_HOJA);
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, ENCABEZADOS.length).getValues()
    .filter(r => r[11] !== 'Sí');            // excluye facturas repetidas
  if (!filas.length) { Logger.log('No hay participaciones válidas'); return; }
  const g = filas[Math.floor(Math.random() * filas.length)];
  Logger.log('GANADOR: ' + g[0] + ' · ' + g[2] + ' · C.I. ' + g[3] + ' · ' + g[4] + ' · Factura ' + g[9]);
}

/* ---------- utilidades ---------- */
function limpiar(v) {
  let s = String(v == null ? '' : v).trim().slice(0, 300);
  if (/^[=+\-@0-9]/.test(s)) s = "'" + s;    // texto literal: evita fórmulas y conserva ceros (0414…)
  return s;
}
function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
