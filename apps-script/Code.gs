/**
 * Tablero personal de Irvin — API de la hoja para el tablero en Vercel
 * ---------------------------------------------------------------------
 * Pega este archivo en la hoja: Extensiones > Apps Script > Código.gs
 *
 * La página en Vercel NO habla directo con este script: lo hace a través de
 * /api/board (función de Vercel), que agrega la clave API_SECRET. Así la URL
 * y la clave nunca llegan al navegador.
 *
 *   GET  <url>/exec?api=1&secret=XXX                 -> tablero completo
 *   POST <url>/exec  body: {secret, action, ...}      -> save | move | delete
 *
 * Primer uso:
 *   1. Ejecuta crearHoja()       -> crea las pestañas con sus encabezados.
 *   2. Ejecuta generarSecreto()  -> copia la clave del registro a API_SECRET en Vercel.
 *
 * Tu WhatsApp NO va en este archivo (el repo puede ser público).
 * Va en Configuración del proyecto (⚙️) > Propiedades de la secuencia de comandos:
 *   MI_WHATSAPP       10 dígitos, ej. 55XXXXXXXX
 *   CALLMEBOT_APIKEY  (opcional) para el aviso automático diario, ver README
 */

const SHEET_TAREAS = 'Tareas';
const SHEET_EQUIPO = 'Equipo';
const SHEET_RECURSOS = 'Recursos';
const TZ = 'America/Mexico_City';

/* ================== CATEGORÍAS (edítalas aquí si cambian) ================== */
// Cada ámbito tiene su lista de categorías (materias / áreas de tu vida) y su lista de tipos.
// Una tarea siempre lleva: Ámbito + Categoría + Tipo (+ Prioridad).
const AMBITOS = {
  Escuela: {
    prefijo: 'ESC-',
    categorias: ['Algoritmos y Estructura de Datos', 'Programación para Ciencia de Datos',
                 'Introducción a la Ciencia de Datos', 'Ética y Legalidad'],
    tipos: ['Tarea', 'Recordatorio', 'Examen', 'Nota']
  },
  Personal: {
    prefijo: 'PER-',
    categorias: ['IMSS', 'Jandy', 'Familia', 'Pagos'],
    tipos: ['Recordatorio', 'Evento', 'Cita', 'Encargo']
  }
};
const AMBITO_NAMES = Object.keys(AMBITOS);

// Las columnas de "Tareas" se ubican por el TEXTO del encabezado (fila 1), no por su posición.
// Cada campo acepta varios nombres de encabezado (el primero es el que crea crearHoja()).
const HEADERS = {
  id: ['ID'],
  ambito: ['Ámbito', 'Ambito'],
  categoria: ['Categoría', 'Categoria', 'Materia'],
  tipo: ['Tipo'],
  titulo: ['Tarea', 'Título', 'Titulo'],
  descripcion: ['Descripción', 'Descripcion'],
  responsable: ['Responsable'],
  estado: ['Estado'],
  prioridad: ['Prioridad'],
  fecha: ['Fecha', 'Fecha límite'],
  hora: ['Hora'],
  notas: ['Notas'],
  enlace: ['Entregable', 'Enlace'],
  materiales: ['Recursos', 'Materiales'],
  creado: ['Creado'],
  actualizado: ['Actualizado'],
  actualizadoPor: ['Actualizado por']
};
const KEYS = Object.keys(HEADERS);
const REQUIRED = ['id', 'titulo', 'estado'];

const ESTADOS = ['To Do', 'Doing', 'Done'];
const PRIORIDADES = ['Urgente', 'Importante', 'Normal'];
const RESPONSABLE_DEFAULT = 'Irvin';

/* ============================ CONFIGURACIÓN ============================ */

/** Ejecútala UNA vez desde el editor (botón ▶). Crea la clave y la muestra en el registro. */
function generarSecreto() {
  const secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_SECRET', secret);
  Logger.log('API_SECRET = ' + secret);
  Logger.log('Cópiala en Vercel > Settings > Environment Variables > API_SECRET');
}

/** Crea las pestañas que falten, con encabezados y menús desplegables. No borra nada. */
function crearHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tareas = ensureSheet_(ss, SHEET_TAREAS, KEYS.map(k => HEADERS[k][0]));
  ensureSheet_(ss, SHEET_EQUIPO, ['Nombre', 'Color'], [[RESPONSABLE_DEFAULT, '#E5E5E5']]);
  ensureSheet_(ss, SHEET_RECURSOS, ['Etiqueta', 'Enlace', 'Notas', 'Ámbito', 'Categoría'], [
    ['Repo de prácticas PCD', 'https://github.com/', 'Cambia este link por el de tu repo', 'Escuela', 'Programación para Ciencia de Datos']
  ]);

  // Menús desplegables en la hoja (por si editas directo ahí)
  const cols = colMap_(tareas);
  const list = (values) => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
  const allCats = [].concat.apply([], AMBITO_NAMES.map(a => AMBITOS[a].categorias));
  const allTipos = AMBITO_NAMES.map(a => AMBITOS[a].tipos).reduce((acc, t) => acc.concat(t.filter(x => acc.indexOf(x) === -1)), []);
  const rows = Math.max(tareas.getMaxRows() - 1, 1);
  if (cols.ambito) tareas.getRange(2, cols.ambito, rows).setDataValidation(list(AMBITO_NAMES));
  if (cols.categoria) tareas.getRange(2, cols.categoria, rows).setDataValidation(list(allCats));
  if (cols.tipo) tareas.getRange(2, cols.tipo, rows).setDataValidation(list(allTipos));
  if (cols.estado) tareas.getRange(2, cols.estado, rows).setDataValidation(list(ESTADOS));
  if (cols.prioridad) tareas.getRange(2, cols.prioridad, rows).setDataValidation(list(PRIORIDADES));
  tareas.setFrozenRows(1);
  Logger.log('Listo. Pestañas: ' + [SHEET_TAREAS, SHEET_EQUIPO, SHEET_RECURSOS].join(', '));
}

function ensureSheet_(ss, name, headers, seed) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (seed && seed.length) sh.getRange(2, 1, seed.length, seed[0].length).setValues(seed);
  }
  return sh;
}

function checkSecret_(given) {
  const real = PropertiesService.getScriptProperties().getProperty('API_SECRET');
  if (!real) throw new Error('Falta API_SECRET: ejecuta generarSecreto() en el editor de Apps Script.');
  if (!given || String(given) !== real) throw new Error('No autorizado');
}

/* ============================ ENTRADAS WEB ============================ */

function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') {
    return json_(safe_(() => { checkSecret_(e.parameter.secret); return getBoard_(); }));
  }
  return ContentService.createTextOutput('API del tablero personal. Abre el tablero desde Vercel.');
}

function doPost(e) {
  return json_(safe_(() => {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    checkSecret_(body.secret);
    switch (body.action) {
      case 'save':   return saveTask_(body.task, body.user);
      case 'move':   return moveTask_(body.id, body.estado, body.user);
      case 'delete': return deleteTask_(body.id, body.user);
      default: throw new Error('Acción no válida: ' + body.action);
    }
  }));
}

/* ============================== OPERACIONES ============================== */

function getBoard_() {
  return { ok: true, tasks: readTasks_(), team: getTeam_(), recursos: getRecursos_(), ambitos: AMBITOS,
           estados: ESTADOS, prioridades: PRIORIDADES, miWhatsapp: myPhone_(), serverTime: now_() };
}

function readTasks_() {
  const sh = sheet_(SHEET_TAREAS);
  const cols = colMap_(sh);
  const last = sh.getLastRow();
  const tasks = [];
  if (last < 2) return tasks;
  sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().forEach(r => {
    if (!r[cols.id - 1]) return;
    const t = {};
    KEYS.forEach(k => t[k] = cols[k] ? fmt_(k, r[cols[k] - 1]) : '');
    if (ESTADOS.indexOf(t.estado) === -1) t.estado = 'To Do';
    if (!AMBITOS[t.ambito]) t.ambito = /^PER-/i.test(t.id) ? 'Personal' : 'Escuela';
    tasks.push(t);
  });
  return tasks;
}

function saveTask_(task, user) {
  if (!task || !String(task.titulo || '').trim()) throw new Error('La tarea necesita un título.');
  return withLock_(() => {
    const sh = sheet_(SHEET_TAREAS);
    const ambito = AMBITOS[task.ambito] ? task.ambito : 'Escuela';
    const conf = AMBITOS[ambito];
    const values = {
      ambito: ambito,
      categoria: conf.categorias.indexOf(task.categoria) > -1 ? task.categoria : clean_(task.categoria).slice(0, 80),
      tipo: conf.tipos.indexOf(task.tipo) > -1 ? task.tipo : conf.tipos[0],
      titulo: clean_(task.titulo).slice(0, 200),
      descripcion: clean_(task.descripcion).slice(0, 2000),
      responsable: clean_(task.responsable) || RESPONSABLE_DEFAULT,
      estado: ESTADOS.indexOf(task.estado) > -1 ? task.estado : 'To Do',
      prioridad: PRIORIDADES.indexOf(task.prioridad) > -1 ? task.prioridad : 'Normal',
      fecha: parseDate_(task.fecha),
      hora: /^\d{1,2}:\d{2}$/.test(String(task.hora || '')) ? "'" + task.hora : '',
      notas: clean_(task.notas).slice(0, 5000),
      enlace: clean_(task.enlace).slice(0, 1000),
      materiales: clean_(cleanMats_(task.materiales)),
      actualizado: now_(),
      actualizadoPor: clean_(user).slice(0, 60) || RESPONSABLE_DEFAULT
    };

    const cols = colMap_(sh);
    const row = task.id ? findRow_(sh, cols, task.id) : -1;
    if (row === -1) {
      values.id = nextId_(sh, cols, conf.prefijo);
      values.creado = now_();
      const full = new Array(sh.getLastColumn()).fill('');
      Object.keys(values).forEach(k => { if (cols[k]) full[cols[k] - 1] = values[k]; });
      sh.appendRow(full);
    } else {
      // Al editar, solo se tocan los campos que vienen en la petición.
      Object.keys(values).forEach(k => {
        if (k !== 'actualizado' && k !== 'actualizadoPor' && task[k] === undefined) delete values[k];
      });
      writeRow_(sh, cols, row, values);
    }
    return getBoard_();
  });
}

function moveTask_(id, estado, user) {
  if (ESTADOS.indexOf(estado) === -1) throw new Error('Estado no válido: ' + estado);
  return withLock_(() => {
    const sh = sheet_(SHEET_TAREAS);
    const cols = colMap_(sh);
    const row = findRow_(sh, cols, id);
    if (row === -1) throw new Error('No encontré la tarea ' + id + '.');
    writeRow_(sh, cols, row, { estado: estado, actualizado: now_(), actualizadoPor: clean_(user).slice(0, 60) || RESPONSABLE_DEFAULT });
    return getBoard_();
  });
}

function deleteTask_(id) {
  return withLock_(() => {
    const sh = sheet_(SHEET_TAREAS);
    const row = findRow_(sh, colMap_(sh), id);
    if (row !== -1) sh.deleteRow(row);
    return getBoard_();
  });
}

/* ===================== AVISO AUTOMÁTICO A TU WHATSAPP ===================== */
// Opcional. Usa CallMeBot (gratis) para mandarte un WhatsApp cada mañana con lo
// vencido, lo de hoy y lo de mañana. Necesita MI_WHATSAPP y CALLMEBOT_APIKEY en
// las propiedades del script (ver README). Luego ejecuta instalarAvisoDiario() una vez.

/** Crea (o reemplaza) el disparador que corre avisoDiario() todos los días a las 7 a. m. */
function instalarAvisoDiario() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'avisoDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('avisoDiario').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  Logger.log('Aviso diario instalado: todos los días a las 7:00 (hora CDMX).');
}

/** Quita el aviso diario. */
function quitarAvisoDiario() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'avisoDiario') ScriptApp.deleteTrigger(t); });
  Logger.log('Aviso diario desactivado.');
}

/** Manda el resumen ahora mismo (sirve para probar). */
function probarAviso() { avisoDiario(true); }

function avisoDiario(forzar) {
  const tasks = readTasks_().filter(t => t.estado !== 'Done' && t.fecha);
  const hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const man = Utilities.formatDate(new Date(Date.now() + 86400000), TZ, 'yyyy-MM-dd');
  const linea = t => '• ' + t.titulo + ' — ' + t.tipo + (t.categoria ? ' · ' + t.categoria : '') +
    (t.hora ? ' · ' + t.hora : '') + (t.prioridad === 'Urgente' ? ' ⚠️' : '');
  const vencidas = tasks.filter(t => t.fecha < hoy), deHoy = tasks.filter(t => t.fecha === hoy), deMan = tasks.filter(t => t.fecha === man);
  if (!vencidas.length && !deHoy.length && !deMan.length && forzar !== true) return;

  const out = ['*Tu día · ' + Utilities.formatDate(new Date(), TZ, 'dd/MM') + '*'];
  if (vencidas.length) out.push('', '*Vencidas*', ...vencidas.map(linea));
  if (deHoy.length) out.push('', '*Hoy*', ...deHoy.map(linea));
  if (deMan.length) out.push('', '*Mañana*', ...deMan.map(linea));
  if (out.length === 1) out.push('', 'Nada pendiente para hoy ni mañana.');
  sendWhatsApp_(out.join('\n'));
}

function sendWhatsApp_(text) {
  const props = PropertiesService.getScriptProperties();
  const phone = myPhone_(), key = props.getProperty('CALLMEBOT_APIKEY');
  if (!phone || !key) throw new Error('Faltan MI_WHATSAPP o CALLMEBOT_APIKEY en las propiedades del script.');
  const url = 'https://api.callmebot.com/whatsapp.php?phone=%2B' + phone +
    '&text=' + encodeURIComponent(text) + '&apikey=' + encodeURIComponent(key);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('CallMeBot respondió ' + res.getResponseCode());
}

/** Tu número con lada de país (52 + 10 dígitos), leído de las propiedades del script. */
function myPhone_() {
  let d = String(PropertiesService.getScriptProperties().getProperty('MI_WHATSAPP') || '').replace(/\D/g, '');
  if (d.length === 10) d = '52' + d;
  return d.length >= 11 ? d : '';
}

/* ================================ HELPERS ================================ */

function getTeam_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EQUIPO);
  const fallback = [{ nombre: RESPONSABLE_DEFAULT, color: '#E5E5E5' }];
  if (!sh || sh.getLastRow() < 2) return fallback;
  const team = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ nombre: String(r[0]).trim(), color: String(r[1] || '#E5E5E5').trim() }));
  return team.length ? team : fallback;
}

/** Recursos (pestaña "Recursos": Etiqueta · Enlace · Notas · Ámbito · Categoría). Ámbito/Categoría vacíos = aplica a todo. */
function getRecursos_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECURSOS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 5)).getValues()
    .filter(r => String(r[0]).trim() && /^https?:\/\//i.test(String(r[1]).trim()))
    .map(r => ({ etiqueta: String(r[0]).trim(), enlace: String(r[1]).trim(), notas: String(r[2] || '').trim(),
                 ambito: String(r[3] || '').trim(), categoria: String(r[4] || '').trim() }));
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('No existe la pestaña "' + name + '". Ejecuta crearHoja() en Apps Script.');
  return sh;
}

/** { clave: númeroDeColumna } leyendo los encabezados de la fila 1. */
function colMap_(sh) {
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
    .map(h => String(h).trim().toLowerCase());
  const cols = {};
  KEYS.forEach(k => {
    for (let n = 0; n < HEADERS[k].length; n++) {
      const i = head.indexOf(HEADERS[k][n].toLowerCase());
      if (i > -1) { cols[k] = i + 1; break; }
    }
  });
  const missing = REQUIRED.filter(k => !cols[k]);
  if (missing.length) {
    throw new Error('En la pestaña "Tareas" falta el encabezado: ' + missing.map(k => '"' + HEADERS[k][0] + '"').join(', '));
  }
  return cols;
}

function writeRow_(sh, cols, row, values) {
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const current = range.getValues()[0];
  Object.keys(values).forEach(k => { if (cols[k]) current[cols[k] - 1] = values[k]; });
  range.setValues([current]);
}

function findRow_(sh, cols, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sh.getRange(2, cols.id, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** Siguiente ID con el prefijo del ámbito: ESC-001, PER-001… */
function nextId_(sh, cols, prefix) {
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, cols.id, last - 1, 1).getValues().forEach(r => {
      const s = String(r[0]);
      if (s.indexOf(prefix) !== 0) return;
      const m = s.match(/(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function fmt_(key, v) {
  if (v instanceof Date) {
    if (key === 'fecha') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    if (key === 'hora') return Utilities.formatDate(v, TZ, 'HH:mm');
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  }
  return v === null || v === undefined ? '' : String(v);
}

function parseDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}

/** Deja solo líneas que traen un link http(s); máximo 30 líneas. */
function cleanMats_(s) {
  return String(s || '').split(/\r?\n/)
    .map(l => l.trim()).filter(l => /https?:\/\/\S+/i.test(l))
    .slice(0, 30).map(l => l.slice(0, 500)).join('\n');
}

/** Evita que un texto se interprete como fórmula en la hoja. */
function clean_(s) {
  s = s === null || s === undefined ? '' : String(s).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function safe_(fn) {
  try { return fn(); } catch (err) { return { ok: false, error: String(err.message || err) }; }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
