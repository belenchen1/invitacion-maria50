/**
 * RSVP Backend — Emilia 50
 * ══════════════════════════════════════════════════════════
 *
 * PASOS PARA CONFIGURAR:
 *
 * 1. Crear una hoja de cálculo nueva en Google Sheets.
 *    - Renombrá la primera hoja a "Invitados"
 *    - Pegá la lista de invitados en la columna A (sin encabezado, una persona por fila)
 *      o dejá que las confirmaciones se vayan registrando solas.
 *    - La columna B será "Confirmado" (se llena automáticamente con "Si")
 *    - La columna C será la fecha/hora de confirmación
 *
 *    Estructura sugerida de la hoja "Invitados":
 *    | A (Nombre)              | B (Confirmado) | C (Fecha)          |
 *    |-------------------------|----------------|--------------------|
 *    | Emilia Fernandez        | Si             | 30/12/2026 20:15   |
 *    | Belén Chen              |                |                    |
 *    | ...                     |                |                    |
 *
 * 2. En la hoja de cálculo: Extensiones > Apps Script
 *    - Borrar el código que viene por defecto
 *    - Pegar TODO el contenido de este archivo
 *    - Guardar (Ctrl+S)
 *
 * 3. Desplegar como aplicación web:
 *    - Botón "Implementar" > "Nueva implementación"
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta de Google)
 *    - Quién tiene acceso: Cualquier persona
 *    - Clic en "Implementar" → Autorizar → Copiar la URL
 *
 * 4. En index.html, reemplazar la línea:
 *    const APPS_SCRIPT_URL = 'REPLACE_WITH_YOUR_APPS_SCRIPT_URL';
 *    con la URL que copiaste en el paso anterior.
 *
 * 5. Cada vez que modifiques este archivo, creá una NUEVA implementación
 *    (no actualices la existente) para que los cambios tomen efecto.
 * ══════════════════════════════════════════════════════════
 */

// Nombre de la hoja dentro del Google Spreadsheet
var SHEET_NAME = 'Invitados';

/**
 * Manejador principal de peticiones GET.
 * Parámetros de query:
 *   action=confirm&name=Nombre Apellido  →  confirma la asistencia
 *   action=list                          →  devuelve la lista con estados
 */
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action || 'list';
  var result;

  try {
    if (action === 'confirm') {
      if (!params.name) throw new Error('Falta el parámetro "name"');
      result = confirmGuest(params.name);
    } else {
      result = listGuests();
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Devuelve todos los invitados con su estado de confirmación.
 */
function listGuests() {
  var ws = getSheet();
  var data = ws.getDataRange().getValues();
  var guests = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    guests.push({
      name: name,
      confirmed: String(data[i][1] || '').toLowerCase() === 'si'
    });
  }
  return { success: true, guests: guests };
}

/**
 * Marca a un invitado como confirmado.
 * Si el nombre no se encuentra, lo registra igual al final de la lista.
 */
function confirmGuest(name) {
  var ws = getSheet();
  var data = ws.getDataRange().getValues();
  var now = Utilities.formatDate(
    new Date(),
    'America/Argentina/Buenos_Aires',
    'dd/MM/yyyy HH:mm'
  );

  // Buscar nombre exacto primero
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === name.trim()) {
      ws.getRange(i + 1, 2).setValue('Si');
      ws.getRange(i + 1, 3).setValue(now);
      return { success: true, name: name, row: i + 1 };
    }
  }

  // Si no se encontró (caso improbable desde el buscador), igual se registra
  var lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1).setValue(name.trim());
  ws.getRange(lastRow, 2).setValue('Si');
  ws.getRange(lastRow, 3).setValue(now);
  return { success: true, name: name, row: lastRow, note: 'nuevo invitado agregado' };
}

/**
 * Helper: obtiene la hoja de invitados.
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(SHEET_NAME);
  if (!ws) {
    throw new Error('No se encontró la hoja "' + SHEET_NAME + '"');
  }
  return ws;
}
