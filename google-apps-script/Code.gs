/*************************************************
 * Miradesfest 2026 - Recepción de votos y comentarios
 *
 * - Votos: un solo voto por dispositivo y película.
 *   Si la pareja Device ID + Película ID ya existe, se
 *   ACTUALIZA esa fila (gana el último voto, con su timestamp).
 * - Comentarios: van a la pestaña "Comentarios" de la misma
 *   planilla, uno por dispositivo y película (el último
 *   comentario actualiza la fila). Nunca generan fila de voto.
 * - Sanitizado: los textos que empiezan con = + - @ se guardan
 *   como texto literal (prefijo ') para que no se ejecuten como
 *   fórmulas. Límites: nombre 32, comentario 512 caracteres.
 *************************************************/

var HOJA_VOTOS = 'Votos';
var HOJA_COMENTARIOS = 'Comentarios';

var MAX_NOMBRE = 32;
var MAX_COMENTARIO = 512;
var MAX_CAMPO = 200;

var VOTOS_VALIDOS = { 'like': true, 'dislike': true, 'later': true };

var ENCABEZADOS_VOTOS = ['Timestamp', 'Event ID', 'Device ID', 'Película ID', 'Película', 'Dirección', 'Voto', 'Comentario', 'Nombre', 'Idioma', 'Día', 'User agent'];
var ENCABEZADOS_COMENTARIOS = ['Timestamp', 'Event ID', 'Device ID', 'Película ID', 'Película', 'Comentario', 'Nombre', 'Idioma', 'Día', 'User agent'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    var deviceId = limpiar_(data.deviceId, MAX_CAMPO);
    var filmId = limpiar_(data.filmId, MAX_CAMPO);
    if (!deviceId || !filmId) {
      return responder_({ ok: false, error: 'missing deviceId or filmId' });
    }

    var d = {
      eventId: limpiar_(data.eventId, MAX_CAMPO),
      deviceId: deviceId,
      filmId: filmId,
      pelicula: limpiar_(data.pelicula, MAX_CAMPO),
      direccion: limpiar_(data.direccion, MAX_CAMPO),
      idioma: limpiar_(data.idioma, 10),
      dia: limpiar_(data.dia, MAX_CAMPO),
      userAgent: limpiar_(data.userAgent, 300)
    };

    var voto = (typeof data.voto === 'string' && VOTOS_VALIDOS[data.voto]) ? data.voto : '';
    var comentario = limpiar_(data.comentario, MAX_COMENTARIO);
    var nombre = limpiar_(data.nombre, MAX_NOMBRE) || 'Anónimo';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultado = { ok: true, voto: false, comentario: false };

    if (voto) {
      upsertVoto_(ss, d, voto, nombre);
      resultado.voto = true;
    }
    if (comentario) {
      upsertComentario_(ss, d, comentario, nombre);
      resultado.comentario = true;
    }
    return responder_(resultado);
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return responder_({ ok: true, servicio: 'Miradesfest 2026 votos' });
}

// Guarda el voto: una sola fila por Device ID + Película ID.
// Si ya existe, actualiza esa fila (último voto gana).
function upsertVoto_(ss, d, voto, nombre) {
  var hoja = obtenerHoja_(ss, HOJA_VOTOS, ENCABEZADOS_VOTOS);
  var datos = hoja.getDataRange().getValues();
  var fila = buscarFila_(datos, d.deviceId, d.filmId);
  var ahora = new Date();
  if (fila > 0) {
    var comentarioPrevio = datos[fila - 1][7]; // columna H: no tocar comentarios históricos
    hoja.getRange(fila, 1, 1, ENCABEZADOS_VOTOS.length).setValues([[
      ahora, d.eventId, d.deviceId, d.filmId, d.pelicula, d.direccion, voto,
      comentarioPrevio, nombre, d.idioma, d.dia, d.userAgent
    ]]);
  } else {
    hoja.appendRow([
      ahora, d.eventId, d.deviceId, d.filmId, d.pelicula, d.direccion, voto,
      '', nombre, d.idioma, d.dia, d.userAgent
    ]);
  }
}

// Guarda el comentario en la pestaña "Comentarios": una sola fila por
// Device ID + Película ID (el último comentario actualiza la fila).
function upsertComentario_(ss, d, comentario, nombre) {
  var hoja = obtenerHoja_(ss, HOJA_COMENTARIOS, ENCABEZADOS_COMENTARIOS);
  var datos = hoja.getDataRange().getValues();
  var fila = buscarFila_(datos, d.deviceId, d.filmId);
  var ahora = new Date();
  if (fila > 0) {
    hoja.getRange(fila, 1, 1, ENCABEZADOS_COMENTARIOS.length).setValues([[
      ahora, d.eventId, d.deviceId, d.filmId, d.pelicula, comentario,
      nombre, d.idioma, d.dia, d.userAgent
    ]]);
  } else {
    hoja.appendRow([
      ahora, d.eventId, d.deviceId, d.filmId, d.pelicula, comentario,
      nombre, d.idioma, d.dia, d.userAgent
    ]);
  }
}

// Devuelve el número de fila (1-based) de la última coincidencia de
// Device ID (col C) + Película ID (col D), o -1 si no existe.
// La fila 1 es el encabezado y nunca se considera.
function buscarFila_(datos, deviceId, filmId) {
  for (var i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][2]) === deviceId && String(datos[i][3]) === filmId) {
      return i + 1;
    }
  }
  return -1;
}

// Devuelve la pestaña; si no existe, la crea con sus encabezados.
function obtenerHoja_(ss, nombreHoja, encabezados) {
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Texto seguro para la planilla: recorta al máximo y neutraliza
// fórmulas (valores que empiezan con = + - @ quedan como texto).
function limpiar_(valor, max) {
  var s = (valor === null || valor === undefined) ? '' : String(valor);
  s = s.trim();
  if (s.length > max) s = s.substring(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function responder_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
