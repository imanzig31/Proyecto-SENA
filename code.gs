// ==========================================
// CONFIGURACIÓN GLOBAL
// ==========================================
var AVI_LOG_SPREADSHEET_ID = ''; 
var AVI_LOG_SHEET_NAME = 'Hoja 1';

/**
 * Función obligatoria de Google Apps Script que maneja las solicitudes HTTP GET.
 * Renderiza la interfaz de usuario (Index) y configura el contexto inicial del usuario.
 */
function doGet() { 
  // 1. Obtenemos el correo del usuario autenticado actualmente
  var userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  
  // Fallback por si no devuelve nada (muy común fuera de dominio) 
  if (!userEmail) { 
    userEmail = 'anonimo@desconocido.com'; 
  }
  
  // 2. Creamos la plantilla HTML basada en el archivo 'Index'
  var template = HtmlService.createTemplateFromFile('Index');
  
  // 3. Pasamos las variables de contexto al HTML
  template.correo = userEmail;
  template.sessionId = Utilities.getUuid(); // ID de sesión corto único
  
  // 4. Devolvemos la página renderizada con su título, permisos de marcos y metaetiquetas
  return template.evaluate() 
    .setTitle('AVI CX - AVI 3.0') 
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

/**
 * Utilidad opcional para obtener el correo electrónico del usuario activo.
 */
function getEmail() { 
  return Session.getActiveUser().getEmail(); 
}

/**
 * Normaliza una cadena de texto eliminando acentos, símbolos, etiquetas HTML y espacios sobrantes,
 * convirtiéndola a minúsculas para unificar comparaciones.
 */
function normalizarTextoBotAvi_(texto) { 
  return String(texto || '') 
    .toLowerCase() 
    .normalize('NFD') 
    .replace(/[\u0300-\u036f]/g, '') 
    .replace(/<[^>]*>/g, ' ') 
    .replace(/[¿?¡!.,;:()[]{}"']/g, ' ') 
    .replace(/\s+/g, ' ') 
    .trim(); 
}

/**
 * Comprueba si la respuesta dada por el bot coincide con alguna frase predefinida de fallo o desconocimiento.
 */
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

/**
 * Extrae y valida un correo electrónico válido de un texto dado mediante expresiones regulares.
 */
function extraerCorreoAvi_(valor) { 
  var texto = String(valor || '').trim(); 
  var match = texto.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i); 
  return match ? match[0].toLowerCase() : ''; 
}

/**
 * Divide la ruta de sesión (sessionPath) en un objeto separando el ID de la pregunta y la sesión.
 */
function dividirSesionAvi_(sessionPath) { 
  var texto = String(sessionPath || '').trim(); 
  if (!texto) { 
    return { idPregunta: '', sesion: '' }; 
  }
  var lastSlashIndex = texto.lastIndexOf('/'); 
  if (lastSlashIndex === -1) { 
    return { idPregunta: '', sesion: texto }; 
  }
  return { idPregunta: texto.slice(0, lastSlashIndex), sesion: texto.slice(lastSlashIndex + 1) }; 
}

/**
 * Verifica si existe un registro idéntico reciente en la hoja de cálculo de logs para evitar duplicados.
 */
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

/**
 * Registra los fallos o respuestas predeterminadas del bot enviadas desde el lado del cliente 
 * directamente en la hoja de cálculo configurada para logs.
 */
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
  
  sheet.appendRow([ new Date(), pregunta, respuesta, 'Sin feedback', sesion.idPregunta, sesion.sesion, '', correo ]);
  return { ok: true, reason: 'logged' }; 
}

/**
 * Función para traer los datos de contactos desde una hoja de cálculo externa de Google Sheets.
 */
function getContacts() { 
  try { 
    var sheetId = '155mUvG69bWx9EQZnPlZoD7EqEPt2vohtNTjHfnJVnKc'; 
    var sheetName = 'Correos Corporativos';
    
    // Conectamos con el archivo y la pestaña correspondiente
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
 * Traductor Inteligente utilizando LanguageApp. 
 * Detecta automáticamente si debe traducir de español a inglés o viceversa.
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
 * Corrector Ortográfico (Solo en Español) integrando la API externa de LanguageTool.
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

/**
 * Función auxiliar para forzar la concesión de permisos de red externos y acceso a hojas de cálculo.
 */
function forzarPermisos() { 
  UrlFetchApp.fetch('https://api.languagetool.org/v2/check', { 
    method: 'post', 
    payload: { text: 'hola', language: 'es' } 
  }); 
  SpreadsheetApp.openById('155mUvG69bWx9EQZnPlZoD7EqEPt2vohtNTjHfnJVnKc'); 
}