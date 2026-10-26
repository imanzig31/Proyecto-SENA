var AVI_LOG_SPREADSHEET_ID = '';
var AVI_LOG_SHEET_NAME = 'Hoja 1';

function doGet() {
  // 1. Obtenemos el correo del usuario autenticado
  var userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();

  // Fallback por si no devuelve nada (muy común fuera de dominio)
  if (!userEmail) {
    userEmail = 'anonimo@desconocido.com';
  }

  // 2. Creamos la plantilla HTML
  var template = HtmlService.createTemplateFromFile('Index');

  // 3. Pasamos el correo al HTML
  template.correo = userEmail;
  
  // Mantenemos un session-id corto para no salirnos de los limites del widget.
  template.sessionId = Utilities.getUuid();

  // 4. Devolvemos la página
  return template.evaluate()
    .setTitle('AVI CX - AVI 3.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Utilidad opcional
 */
function getEmail() {
  return Session.getActiveUser().getEmail();
}

function normalizarTextoBotAvi_(texto) {
  return String(texto || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/[¿?¡!.,;:()[\]{}"']/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function esRespuestaDeFalloAvi_(textoRespuesta) {
  var frases = [
    'lo siento estoy aprendiendo del tema',
    'ups no he entendido a que te refieres',
    'podrias repetirlo por favor',
    'disculpa',
    'decias',
    'como',
    'lo siento no tengo informacion sobre',
    'solo puedo responder preguntas basadas en la informacion'
  ];
  var textoNormalizado = normalizarTextoBotAvi_(textoRespuesta);
  if (!textoNormalizado) {
    return false;
  }

  return frases.some(function(frase) {
    return textoNormalizado === normalizarTextoBotAvi_(frase);
  });
}

function extraerCorreoAvi_(valor) {
  var texto = String(valor || '').trim();
  var match = texto.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0].toLowerCase() : '';
}

function dividirSesionAvi_(sessionPath) {
  var texto = String(sessionPath || '').trim();
  if (!texto) {
    return { idPregunta: '', sesion: '' };
  }

  var lastSlashIndex = texto.lastIndexOf('/');
  if (lastSlashIndex === -1) {
    return {
      idPregunta: '',
      sesion: texto
    };
  }

  return {
    idPregunta: texto.slice(0, lastSlashIndex),
    sesion: texto.slice(lastSlashIndex + 1)
  };
}

function existeRegistroRecienteAvi_(sheet, pregunta, respuesta, correo, sesion, idPregunta) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var totalFilas = Math.min(100, lastRow - 1);
  var inicio = lastRow - totalFilas + 1;
  var datos = sheet.getRange(inicio, 1, totalFilas, Math.max(sheet.getLastColumn(), 8)).getValues();

  for (var i = datos.length - 1; i >= 0; i--) {
    var fila = datos[i];
    var mismaPregunta = String(fila[1] || '').trim() === pregunta;
    var mismaRespuesta = String(fila[2] || '').trim() === respuesta;
    var mismoCorreo = extraerCorreoAvi_(fila[7]) === correo;
    var mismaSesion = String(fila[5] || '').trim() === sesion;
    var mismoIdPregunta = String(fila[4] || '').trim() === idPregunta;

    if (mismaPregunta && mismaRespuesta && mismoCorreo && mismaSesion && mismoIdPregunta) {
      return true;
    }
  }

  return false;
}

function registrarFallbackBotDesdeCliente(payload) {
  payload = payload || {};

  var pregunta = String(payload.pregunta || '').trim();
  var respuesta = String(payload.respuesta || '').trim();
  var correo = extraerCorreoAvi_(payload.correo) || 'N/A';
  var sesion = dividirSesionAvi_(payload.sessionPath);

  if (!pregunta || !respuesta || !esRespuestaDeFalloAvi_(respuesta)) {
    return { ok: false, reason: 'not_bot_failure' };
  }

  var ss = SpreadsheetApp.openById(AVI_LOG_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(AVI_LOG_SHEET_NAME);
  if (!sheet) {
    throw new Error('No se encontró la hoja de log.');
  }

  if (existeRegistroRecienteAvi_(sheet, pregunta, respuesta, correo, sesion.sesion, sesion.idPregunta)) {
    return { ok: true, reason: 'duplicate' };
  }

  sheet.appendRow([
    new Date(),
    pregunta,
    respuesta,
    'Sin feedback',
    sesion.idPregunta,
    sesion.sesion,
    '',
    correo
  ]);

  return { ok: true, reason: 'logged' };
}

/**
 * Función para traer los datos del Google Sheet de Contactos
 */
function getContacts() {
  try {
    var sheetId = '155mUvG69bWx9EQZnPlZoD7EqEPt2vohtNTjHfnJVnKc';
    var sheetName = 'Correos Corporativos';
    
    // Conectamos con el archivo
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error("No se encontró la pestaña llamada '" + sheetName + "'. Verifica mayúsculas y minúsculas.");
    }
    
    var data = sheet.getDataRange().getValues();
    var contacts = [];
    
    // Iteramos desde 1 para omitir la fila 0 (Los encabezados)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0] !== '') { 
        contacts.push({
          nombre: String(data[i][0]).trim(),
          correo: String(data[i][1]).trim(),
          cargo: String(data[i][2]).trim()
        });
      }
    }
    
    return contacts;
  } catch (e) {
    throw new Error("Error de permisos o lectura: " + e.message);
  }
}

/**
 * Traductor Inteligente de Google
 */
function translateText(text) {
  try {
    if (!text || text.trim() === '') return { text: '', label: 'Listo' };
    var translatedToEs = LanguageApp.translate(text, '', 'es');
    if (translatedToEs.toLowerCase().trim() === text.toLowerCase().trim()) {
      var translatedToEn = LanguageApp.translate(text, '', 'en');
      return { text: translatedToEn, label: 'Español ➔ Inglés' };
    } else {
      return { text: translatedToEs, label: 'Inglés ➔ Español' };
    }
  } catch(e) {
    return { text: "Error: " + e.message, label: "Error de sistema" };
  }
}

/**
 * Corrector Ortográfico (Solo Español)
 */
function correctSpelling(text) {
  try {
    if (!text || text.trim() === '') return { text: '', label: 'Listo' };
    var url = 'https://api.languagetool.org/v2/check';
    var payload = { 'text': text, 'language': 'es' };
    var options = { 'method': 'post', 'payload': payload };
    
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    
    var matches = data.matches;
    if (!matches || matches.length === 0) return { text: text, label: '¡Sin errores!' };
    
    var correctedText = text;
    matches.sort(function(a, b) { return b.offset - a.offset; });
    
    for (var i = 0; i < matches.length; i++) {
      var match = matches[i];
      if (match.replacements && match.replacements.length > 0) {
        var replacement = match.replacements[0].value; 
        correctedText = correctedText.substring(0, match.offset) + replacement + correctedText.substring(match.offset + match.length);
      }
    }
    return { text: correctedText, label: 'Corregido (' + matches.length + ' errores)' };
  } catch(e) {
    return { text: "Faltan permisos externos. Detalles: " + e.message, label: "Error de sistema" };
  }
}

function forzarPermisos() {
  UrlFetchApp.fetch('https://api.languagetool.org/v2/check', {
    method: 'post', payload: { text: 'hola', language: 'es' }
  });
  SpreadsheetApp.openById('155mUvG69bWx9EQZnPlZoD7EqEPt2vohtNTjHfnJVnKc');
}

