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
 * Tu WhatsApp y tus llaves NO van en este archivo (el repo es público).
 * Van en Configuración del proyecto (⚙️) > Propiedades de la secuencia de comandos:
 *   MI_WHATSAPP         10 dígitos, ej. 55XXXXXXXX (botón «Mandármelo por WhatsApp»)
 *   FORMSPREE_ENDPOINT  https://formspree.io/f/xxxx  (recordatorios por el plugin de WhatsApp de Formspree)
 *   FORMSPREE_LIMITE    (opcional) envíos al mes antes de pausar Formspree; por defecto 50 (plan gratis)
 *   CALLMEBOT_APIKEY    (opcional) canal gratis de WhatsApp; se usa si no hay Formspree o si se acabó el cupo
 *   AVISOS_WHATSAPP     (opcional) 'resumen' o 'todo'; con Google Calendar sincronizado es 'resumen' por defecto
 * Recordatorios a lo largo del día: ejecuta instalarRecordatorios() una vez (ver README).
 * Google Calendar: ejecuta instalarSincronizacion() una vez (crea el calendario «Dame la Vida», ver README).
 */

const SHEET_TAREAS = 'Tareas';
const SHEET_EQUIPO = 'Equipo';
const SHEET_RECURSOS = 'Recursos';
const SHEET_HORARIO = 'Horario';
const SHEET_PLAN = 'Plan';
const SHEET_CAL_IPN = 'Calendario IPN';
const SHEET_FIJAS = 'Fijas';
const SHEET_HECHOS = 'Hechos';
const SHEET_CORTES = 'Cortes';
const SHEET_CAMINOS = 'Caminos';
const SHEET_TEMARIO = 'Temario';
const SHEET_PAGOS_JANDY = 'Pagos Jandy';
const TZ = 'America/Mexico_City';

/* ================== CATEGORÍAS (edítalas aquí si cambian) ================== */
// Cada ámbito tiene su lista de categorías (materias / áreas de tu vida), su lista de tipos
// y, en Desarrollo profesional, la lista de escuelas / plataformas (cada asignación lleva hasta dos).
// Una tarea siempre lleva: Ámbito + Categoría + Tipo (+ Urgencia).
const AMBITOS = {
  'Desarrollo profesional': {
    prefijo: 'DP-',
    categorias: ['Algoritmos y Estructura de Datos', 'Programación para Ciencia de Datos',
                 'Introducción a la Ciencia de Datos', 'Ética y Legalidad', 'Matemáticas', 'Inglés',
                 'Nube', 'Framework', 'UX/UI', 'Python', 'C', 'FrontEnd', 'BackEnd'],
    tipos: ['Tarea', 'Práctica', 'Estudio', 'Curso', 'Workshop', 'Certificación',
            'Evento', 'Recordatorio', 'Examen', 'Nota'],
    instituciones: ['ESCOM', 'EBAC', 'Google', 'AWS', 'Oracle', 'FinOps', 'Udemy',
                    'Domestika', 'Código Facilito', 'Platzi', 'DataCamp']
  },
  Personal: {
    prefijo: 'PER-',
    categorias: ['IMSS', 'Salud', 'Jandy', 'Familia', 'Pagos'],
    tipos: ['Recordatorio', 'Evento', 'Cita', 'Encargo'],
    instituciones: []
  }
};
const MAX_INSTITUCIONES = 2;
const AMBITO_NAMES = Object.keys(AMBITOS);
const AMBITO_DP = 'Desarrollo profesional';
/** Nombres viejos que siguen funcionando si aparecen en la hoja. */
const AMBITO_ALIAS = { Escuela: AMBITO_DP, DP: AMBITO_DP };

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
  inicio: ['Inicio', 'Desde'],
  instituciones: ['Instituciones', 'Escuelas', 'Plataformas'],
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
    ['Repo de prácticas PCD', 'https://github.com/', 'Cambia este link por el de tu repo', AMBITO_DP, 'Programación para Ciencia de Datos']
  ]);
  ensureSheet_(ss, SHEET_PLAN, ['ID', 'Actividad', 'Días', 'Inicio', 'Fin', 'Desde', 'Hasta', 'Notas', 'Excepto']);
  ensureSheet_(ss, SHEET_FIJAS, ['Actividad', 'Tipo', 'Días', 'Día del mes', 'Inicio', 'Fin', 'Desde', 'Hasta', 'Excepto', 'Lugar', 'Notas', 'Aviso (min antes)']);
  ensureSheet_(ss, SHEET_CAL_IPN, ['Evento', 'Inicio', 'Fin', 'Tipo', 'Notas']);
  ensureSheet_(ss, SHEET_HECHOS, ['Clave', 'Fecha', 'Marcado en']);
  ensureSheet_(ss, SHEET_CORTES, CORTES_HEADERS);
  ensureSheet_(ss, SHEET_CAMINOS, ['ID', 'Título', 'Tipo', 'Actividad (Fijas)', 'Desde', 'Hasta', 'Meta', 'Meta fecha', 'Color', 'Notas']);
  ensureSheet_(ss, SHEET_TEMARIO, ['Camino', 'Bloque', 'Módulo', 'Lecciones', 'Fecha meta', 'Hecho', 'Notas']);

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
  Logger.log('Listo. Pestañas: ' + [SHEET_TAREAS, SHEET_EQUIPO, SHEET_RECURSOS, SHEET_PLAN, SHEET_CAL_IPN, SHEET_FIJAS].join(', '));
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
      case 'hecho':  return toggleHecho_(body.clave, body.fecha, body.done);
      case 'corte':  return saveCorte_(body.corte);
      case 'corteDel': return deleteCorte_(body.id);
      case 'modulo': return toggleModulo_(body.camino, body.modulo, body.done);
      case 'pagoJandy': return togglePagoJandy_(body.no, body.done);
      default: throw new Error('Acción no válida: ' + body.action);
    }
  }));
}

/* ============================== OPERACIONES ============================== */

function getBoard_() {
  const tasks = readTasks_();
  const plan = optional_(getPlan_);
  const hechos = optional_(getHechos_);
  progresoTareas_(tasks, plan, hechos);
  return { ok: true, tasks: tasks, team: getTeam_(), recursos: getRecursos_(), ambitos: AMBITOS,
           horario: optional_(getHorario_), plan: plan, calendario: optional_(getCalendarioIPN_),
           fijas: optional_(getFijas_), hechos: hechos,
           cortes: optional_(getCortes_), caminos: optional_(getCaminos_), temario: optional_(getTemario_),
           pagosJandy: optional_(getPagosJandy_),
           estados: ESTADOS, prioridades: PRIORIDADES, miWhatsapp: myPhone_(), serverTime: now_() };
}

/** Clave de ocurrencia marcable, igual en el servidor y en el navegador. */
function claveFija_(f) { return 'fija|' + f.actividad + '|' + f.dias + '|' + (f.diaMes || '') + '|' + (f.inicio || ''); }
function clavePlan_(b) { return 'plan|' + (b.id || '') + '|' + b.actividad + '|' + b.dias + '|' + b.inicio; }
function claveClase_(h) { return 'clase|' + h.dia + '|' + h.inicio + '|' + h.fin + '|' + infoClase_(h.texto).titulo; }

/** Marca o desmarca una ocurrencia (clave, fecha). Devuelve el tablero completo. */
function toggleHecho_(clave, fecha, done) {
  const c = String(clave || '').trim().slice(0, 300), fd = parseDate_(fecha), f = fd ? fmtDay_(fd) : '';
  if (!c || !f) throw new Error('Falta clave o fecha.');
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ensureSheet_(ss, SHEET_HECHOS, ['Clave', 'Fecha', 'Marcado en']);
    const last = sh.getLastRow();
    let fila = -1;
    if (last > 1) {
      const vals = sh.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === c && fmtDay_(vals[i][1]) === f) { fila = i + 2; break; }
    }
    if (done) {
      if (fila === -1) sh.appendRow([clean_(c), fd, now_()]);
      else sh.getRange(fila, 3).setValue(now_());
    } else if (fila !== -1) sh.deleteRow(fila);
    avisarCambio_();
    return getBoard_();
  });
}

/* ---- Cortes (evidencia), Caminos (hábitos y pagos) y Temario (módulos de cursos) ---- */
const CORTES_HEADERS = ['ID', 'Camino', 'Fecha', 'Nota', 'Link', 'Imagen', 'Creado'];
const EVIDENCIAS_CARPETA = 'Dame la Vida · Evidencias';

/** Cortes: puntos de revisión con nota, link y/o captura. Camino = ID de asignación (DP-002) o de la pestaña Caminos. */
function getCortes_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CORTES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({ id: String(r[0]).trim(), camino: String(r[1]).trim(), fecha: fmtDay_(r[2]), nota: String(r[3] || ''),
                 link: String(r[4] || ''), imagen: String(r[5] || '').trim(), creado: fmt_('creado', r[6]) }));
}

function saveCorte_(c) {
  c = c || {};
  const camino = clean_(c.camino).slice(0, 60);
  if (!camino) throw new Error('El corte necesita un camino.');
  const fd = parseDate_(c.fecha) || parseDate_(hoy_());
  const nota = clean_(c.nota).slice(0, 1000);
  const link = /^https?:\/\//i.test(String(c.link || '').trim()) ? String(c.link).trim().slice(0, 1000) : '';
  let imagen = '';
  if (c.imagen) imagen = guardarImagen_(String(c.imagen));
  if (!nota && !link && !imagen) throw new Error('Agrega una nota, un link o una captura.');
  return withLock_(() => {
    const sh = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_CORTES, CORTES_HEADERS);
    const n = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().reduce((m, r) => Math.max(m, parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0), 0) : 0;
    sh.appendRow(['COR-' + String(n + 1).padStart(3, '0'), camino, fd, nota, link, imagen, now_()]);
    return getBoard_();
  });
}

function deleteCorte_(id) {
  return withLock_(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CORTES);
    if (sh && sh.getLastRow() > 1) {
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
      for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(id)) {
        const img = String(vals[i][5] || '').trim();
        if (img) { try { DriveApp.getFileById(img).setTrashed(true); } catch (err) { Logger.log(err); } }
        sh.deleteRow(i + 2); break;
      }
    }
    return getBoard_();
  });
}

/** Guarda una captura (data URL jpeg/png/webp, ya reducida por el navegador) en Drive y devuelve su ID. */
function guardarImagen_(dataUrl) {
  const m = dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) throw new Error('La imagen debe ser JPG, PNG o WebP.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('La imagen pesa más de 3 MB.');
  const props = PropertiesService.getScriptProperties();
  let folder = null, fid = props.getProperty('EVIDENCIAS_FOLDER');
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (err) { folder = null; } }
  if (!folder) {
    const it = DriveApp.getFoldersByName(EVIDENCIAS_CARPETA);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(EVIDENCIAS_CARPETA);
    props.setProperty('EVIDENCIAS_FOLDER', folder.getId());
  }
  const ext = m[1] === 'image/png' ? 'png' : m[1] === 'image/webp' ? 'webp' : 'jpg';
  const file = folder.createFile(Utilities.newBlob(bytes, m[1], 'corte-' + now_().replace(/\D/g, '') + '.' + ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // para que el tablero pueda mostrarla
  return file.getId();
}

/** Caminos de hábitos y pagos (pestaña Caminos): ID · Título · Tipo · Actividad (Fijas) · Desde · Hasta · Meta · Meta fecha · Color · Notas. */
function getCaminos_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CAMINOS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({ id: String(r[0]).trim(), titulo: String(r[1]).trim(), tipo: norm_(r[2]) === 'pago' ? 'pago' : 'hábito', actividad: String(r[3] || '').trim(),
                 desde: fmtDay_(r[4]), hasta: fmtDay_(r[5]), meta: String(r[6] || '').trim(), metaFecha: fmtDay_(r[7]),
                 color: /^#[0-9a-f]{6}$/i.test(String(r[8] || '').trim()) ? String(r[8]).trim() : '', notas: String(r[9] || '') }));
}

/** Módulos de un curso (pestaña Temario): Camino (ID de asignación) · Bloque · Módulo · Lecciones · Fecha meta · Hecho · Notas. */
function getTemario_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEMARIO);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim())
    .map(r => ({ camino: String(r[0]).trim(), bloque: String(r[1] || '').trim(), modulo: String(r[2]).trim(), lecciones: Number(r[3]) || 0,
                 fecha: fmtDay_(r[4]), hecho: r[5] === true || /^true$/i.test(String(r[5])), notas: String(r[6] || '') }));
}

function toggleModulo_(camino, modulo, done) {
  return withLock_(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEMARIO);
    if (!sh) throw new Error('No existe la pestaña Temario.');
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
    for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === String(camino) && String(vals[i][2]).trim() === String(modulo)) {
      sh.getRange(i + 2, 6).setValue(!!done); break;
    }
    return getBoard_();
  });
}

/** Pagos a Jandy: una fila por quincena (No, Fecha, Monto, Estado, Saldo, Notas). */
function getPagosJandy_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAGOS_JANDY);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[0]).trim() && fmtDay_(r[1]))
    .map(r => ({ no: Number(r[0]) || 0, fecha: fmtDay_(r[1]), monto: Number(r[2]) || 0,
                 estado: String(r[3] || '').trim(), saldo: Number(r[4]) || 0, notas: String(r[5] || '') }));
}

function togglePagoJandy_(no, done) {
  const n = Number(no); if (!n) throw new Error('Falta número de pago.');
  return withLock_(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAGOS_JANDY);
    if (!sh) throw new Error('No existe la pestaña Pagos Jandy.');
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
    for (let i = 0; i < vals.length; i++) if ((Number(vals[i][0]) || 0) === n) {
      sh.getRange(i + 2, 4).setValue(done ? 'Pagado' : 'Pendiente'); break;
    }
    return getBoard_();
  });
}

function norm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }

function getHechos_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HECHOS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(r => String(r[0]).trim() && fmtDay_(r[1]))
    .map(r => ({ clave: String(r[0]).trim(), fecha: fmtDay_(r[1]), marcado: String(r[2] || '') }));
}

/** Para cada tarea en Doing con bloques en Plan, añade t.progreso = {total, elapsed, done}. */
function progresoTareas_(tasks, plan, hechos) {
  const hoy = hoy_();
  const hechosPorClave = {};
  hechos.forEach(h => { (hechosPorClave[h.clave] = hechosPorClave[h.clave] || {})[h.fecha] = 1; });
  tasks.forEach(t => {
    if (t.estado !== 'Doing') return;
    const blocks = plan.filter(b => b.id === t.id);
    if (!blocks.length) return;
    blocks.forEach(b => { if (!b.k) b.k = clavePlan_(b); });
    let desde = blocks.map(b => b.desde).filter(Boolean).sort()[0] || t.inicio || hoy;
    const hastas = blocks.map(b => b.hasta).filter(Boolean).sort();
    let hasta = hastas.length ? hastas[hastas.length - 1] : (t.fecha || '');
    if (!hasta || hasta < hoy) hasta = hoy;        // tarea sin fin: mide hasta hoy
    const tope = ymdMas_(hoy, 400);
    if (hasta > tope) hasta = tope;
    let total = 0, elapsed = 0, done = 0;
    for (let d = desde; d <= hasta; d = ymdMas_(d, 1)) {
      blocks.forEach(b => {
        if ((!b.desde || b.desde <= d) && (!b.hasta || d <= b.hasta) && !excluido_(b.excepto, d) && diasSet_(b.dias).indexOf(dowDe_(d)) > -1) {
          total++;
          if (d <= hoy) elapsed++;
          if (hechosPorClave[b.k] && hechosPorClave[b.k][d]) done++;
        }
      });
    }
    t.progreso = { total: total, elapsed: elapsed, done: done };
  });
}

/** Una pestaña opcional con un error (o que no existe) no debe tumbar el tablero. */
function optional_(fn) { try { return fn(); } catch (err) { Logger.log(err); return []; } }

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
    if (AMBITO_ALIAS[t.ambito]) t.ambito = AMBITO_ALIAS[t.ambito];
    if (!AMBITOS[t.ambito]) t.ambito = /^PER-/i.test(t.id) ? 'Personal' : AMBITO_DP;
    tasks.push(t);
  });
  return tasks;
}

function saveTask_(task, user) {
  if (!task || !String(task.titulo || '').trim()) throw new Error('La tarea necesita un título.');
  return withLock_(() => {
    const sh = sheet_(SHEET_TAREAS);
    const ambito = AMBITOS[task.ambito] ? task.ambito : (AMBITO_ALIAS[task.ambito] || AMBITO_DP);
    const conf = AMBITOS[ambito];
    const values = {
      ambito: ambito,
      categoria: conf.categorias.indexOf(task.categoria) > -1 ? task.categoria : clean_(task.categoria).slice(0, 80),  // vacío = sin categoría
      tipo: conf.tipos.indexOf(task.tipo) > -1 ? task.tipo : conf.tipos[0],
      titulo: clean_(task.titulo).slice(0, 200),
      descripcion: clean_(task.descripcion).slice(0, 2000),
      responsable: clean_(task.responsable) || RESPONSABLE_DEFAULT,
      estado: ESTADOS.indexOf(task.estado) > -1 ? task.estado : 'To Do',
      prioridad: PRIORIDADES.indexOf(task.prioridad) > -1 ? task.prioridad : 'Normal',
      fecha: parseDate_(task.fecha),
      hora: /^\d{1,2}:\d{2}$/.test(String(task.hora || '')) ? "'" + task.hora : '',
      inicio: parseDate_(task.inicio),
      instituciones: cleanInstituciones_(task.instituciones),
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
    avisarCambio_();
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
    avisarCambio_();
    return getBoard_();
  });
}

function deleteTask_(id) {
  return withLock_(() => {
    const sh = sheet_(SHEET_TAREAS);
    const row = findRow_(sh, colMap_(sh), id);
    if (row !== -1) sh.deleteRow(row);
    avisarCambio_();
    return getBoard_();
  });
}

/* ======================= RECORDATORIOS A TU WHATSAPP ======================= */
// Un disparador revisa cada 5 minutos qué empieza pronto (clases, actividades fijas,
// bloques del Plan y tareas con hora) y te manda un mensaje en el momento indicado.
// Canal: Formspree (plugin de WhatsApp del formulario) y, si no hay o se acabó el cupo,
// CallMeBot. Ver README para configurarlos.

const AVISO_CLASE_MIN = 10;   // clases: 10 min antes (para salir a tiempo)
const AVISO_PLAN_MIN = 0;     // bloques de estudio: al empezar
const AVISO_TAREA_MIN = 60;   // tareas con hora: 1 h antes
const RECORDATORIO_CADA_MIN = 5;

/** Instala el disparador cada 5 min (y quita el aviso diario viejo si existía). Ejecútala una vez. */
function instalarRecordatorios() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['recordatorios', 'avisoDiario'].indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('recordatorios').timeBased().everyMinutes(RECORDATORIO_CADA_MIN).create();
  PropertiesService.getScriptProperties().setProperty('RECORDATORIOS_ULTIMO', String(Date.now()));
  Logger.log('Recordatorios activos: revisa cada ' + RECORDATORIO_CADA_MIN + ' min.');
  const n = Object.keys(eventosConAviso_(hoy_()).reduce((o, e) => (o[e.aviso] = 1, o), {})).length;
  Logger.log('Hoy te llegarían ' + n + ' mensajes (unos ' + n * 30 + ' al mes). El plan gratis de Formspree trae 50 al mes.');
}

/** Apaga los recordatorios. */
function quitarRecordatorios() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['recordatorios', 'avisoDiario'].indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  Logger.log('Recordatorios apagados.');
}

/** Lista en el registro los avisos de hoy, sin mandar nada. */
function verAvisosDeHoy() {
  eventosConAviso_(hoy_()).forEach(e => Logger.log(e.aviso + '  ' + e.texto.split('\n')[0]));
}

/** Manda un mensaje de prueba por el canal configurado. */
function probarRecordatorio() { enviarAviso_('Prueba del tablero «Dame la Vida». Si lees esto, los recordatorios funcionan.'); }

/** Manda ahora el resumen del día (lo mismo que llega al despertar). */
function probarAviso() { enviarAviso_(resumenDelDia_(hoy_())); }

/** Lo llama el disparador. Manda lo que toque entre la última revisión y ahora. */
function recordatorios() {
  const props = PropertiesService.getScriptProperties();
  const ahora = Date.now();
  let desde = Number(props.getProperty('RECORDATORIOS_ULTIMO') || 0);
  if (!desde || ahora - desde > 20 * 60000) desde = ahora - RECORDATORIO_CADA_MIN * 60000; // si estuvo pausado, no inundar
  props.setProperty('RECORDATORIOS_ULTIMO', String(ahora));
  const hoy = hoy_(), porHora = {};
  eventosConAviso_(hoy).forEach(e => {
    const t = msDe_(hoy, e.aviso);
    if (t > desde && t <= ahora) (porHora[e.aviso] = porHora[e.aviso] || []).push(e.texto);
  });
  // Lo que cae a la misma hora va en un solo mensaje (cuida el cupo de Formspree).
  Object.keys(porHora).sort().forEach(h => { try { enviarAviso_(porHora[h].join('\n\n')); } catch (err) { Logger.log(err); } });
}

/** Todos los avisos de un día: [{aviso:'HH:mm', texto}] ordenados. */
function eventosConAviso_(ymd) {
  const out = [], resumenEn = { hecho: false };
  const add = (hhmm, minAntes, texto, esencial) => {
    const m = minutos_(hhmm); if (m === null) return;
    out.push({ aviso: hhmmDe_(Math.max(0, m - minAntes)), texto: texto, esencial: !!esencial });
  };
  const ev = eventosDelDia_(ymd);
  ev.clases.forEach(c => add(c.inicio, AVISO_CLASE_MIN, 'En ' + AVISO_CLASE_MIN + ' min: ' + c.titulo + (c.lugar ? ' · ' + c.lugar : '') + ' (' + c.inicio + '–' + c.fin + ')'));
  ev.fijas.forEach(f => {
    if (f.aviso === '' || !f.inicio) return;
    const antes = Number(f.aviso) || 0;
    let txt = (antes ? 'En ' + antes + ' min: ' : '') + f.actividad + (f.fin ? ' (' + f.inicio + '–' + f.fin + ')' : ' · ' + f.inicio) +
      (f.lugar ? '\n' + f.lugar : '') + (f.notas ? '\n' + f.notas : '');
    if (!resumenEn.hecho && /despert/i.test(f.actividad)) { txt = '¡Buenos días! ' + txt + '\n\n' + resumenDelDia_(ymd); resumenEn.hecho = true; }
    add(f.inicio, antes, txt, resumenEn.hecho && /despert/i.test(f.actividad) && txt.indexOf('Buenos días') > -1);
  });
  ev.plan.forEach(b => add(b.inicio, AVISO_PLAN_MIN, 'Ahora: ' + b.actividad + ' hasta las ' + b.fin + (b.notas ? '\n' + b.notas : '')));
  ev.tareas.filter(t => t.hora).forEach(t => add(t.hora, AVISO_TAREA_MIN, 'En 1 h: ' + t.tipo + ' · ' + t.titulo + ' (' + t.hora + ')' + (t.descripcion ? '\n' + t.descripcion : ''), true));
  if (!resumenEn.hecho) out.push({ aviso: '07:00', texto: resumenDelDia_(ymd), esencial: true });
  // Con la sincronización de Google Calendar activa, el calendario avisa de lo rutinario y a WhatsApp
  // solo llega el resumen de la mañana y las citas con hora (así alcanza el plan gratis de Formspree).
  const lista = modoAvisos_() === 'resumen' ? out.filter(e => e.esencial) : out;
  return lista.sort((a, b) => a.aviso < b.aviso ? -1 : 1);
}

/** 'resumen' = solo resumen de la mañana y citas con hora; 'todo' = todos los avisos de la hoja. */
function modoAvisos_() {
  const p = PropertiesService.getScriptProperties();
  return p.getProperty('AVISOS_WHATSAPP') || (p.getProperty('CAL_SYNC_ACTIVA') ? 'resumen' : 'todo');
}

/** Resumen corto: pagos y entregas de hoy, vencidas, lo de mañana y el plan de estudio. */
function resumenDelDia_(ymd) {
  const ev = eventosDelDia_(ymd), man = ymdMas_(ymd, 1);
  const tasks = readTasks_().filter(t => t.estado !== 'Done');
  const linea = t => '• ' + t.titulo + ' — ' + t.tipo + (t.hora ? ' · ' + t.hora : '') + (t.prioridad === 'Urgente' ? ' ⚠️' : '');
  const out = ['*Tu día · ' + ymd.slice(8, 10) + '/' + ymd.slice(5, 7) + '*'];
  const pagos = ev.fijas.filter(f => !f.inicio);
  if (pagos.length) out.push('', '*Hoy toca*', ...pagos.map(f => '• ' + f.actividad + (f.notas ? ' — ' + f.notas : '')));
  const venc = tasks.filter(t => t.fecha && t.fecha < ymd);
  if (venc.length) out.push('', '*Vencidas*', ...venc.map(linea));
  if (ev.tareas.length) out.push('', '*Fecha límite hoy*', ...ev.tareas.map(linea));
  const deMan = tasks.filter(t => t.fecha === man);
  if (deMan.length) out.push('', '*Mañana*', ...deMan.map(linea));
  if (ev.clases.length) out.push('', '*Clases*', ...ev.clases.map(c => '• ' + c.inicio + ' ' + c.titulo));
  if (ev.plan.length) out.push('', '*Plan de estudio*', ...ev.plan.map(b => '• ' + b.inicio + '–' + b.fin + ' ' + b.actividad));
  if (dowDe_(ymd) === 0) out.push('', '*Domingo de corte*', 'Sube tu evidencia de la semana en Progreso: un link o una captura de cada avance.');
  return out.join('\n');
}

/** Clases, fijas, bloques del plan y tareas con fecha de un día. */
function eventosDelDia_(ymd) {
  const hayClases = hayClasesEn_(optional_(getCalendarioIPN_), ymd);
  const dow = dowDe_(ymd);
  const clases = !hayClases ? [] : optional_(getHorario_).filter(h => h.dia === dow).map(h => {
    const ci = infoClase_(h.texto); return { titulo: ci.titulo, lugar: ci.lugar, inicio: h.inicio, fin: h.fin };
  });
  const fijas = fijasDelDia_(optional_(getFijas_), ymd).filter(f => hayClases || !/^clase$/i.test(f.tipo));
  const tareasPorId = {}; const tareas = readTasks_();
  tareas.forEach(t => tareasPorId[t.id] = t);
  const plan = planForDate_(optional_(getPlan_), ymd).filter(b => !(tareasPorId[b.id] && tareasPorId[b.id].estado === 'Done'));
  return { clases: clases, fijas: fijas, plan: plan, tareas: tareas.filter(t => t.estado !== 'Done' && t.fecha === ymd) };
}

/** ¿Hay clases ese día? Sí dentro del periodo «Clases» del calendario IPN y fuera de descansos y vacaciones. */
function hayClasesEn_(cal, ymd) {
  const enRango = e => e.inicio <= ymd && ymd <= (e.fin || e.inicio);
  const libre = cal.some(e => enRango(e) && /descanso|vacaci/i.test(e.tipo));
  const periodos = cal.filter(e => /^clases$/i.test(e.tipo));
  return !libre && (!periodos.length || periodos.some(enRango));
}

function fijasDelDia_(fijas, ymd) {
  const dow = dowDe_(ymd), dia = Number(ymd.slice(8, 10));
  const ultimo = new Date(Number(ymd.slice(0, 4)), Number(ymd.slice(5, 7)), 0).getDate();
  return fijas.filter(f => {
    if ((f.desde && ymd < f.desde) || (f.hasta && ymd > f.hasta) || excluido_(f.excepto, ymd)) return false;
    if (f.diaMes) return String(f.diaMes).split(/[,\s]+/).some(x => /últ|ult/i.test(x) ? dia === ultimo : Number(x) === dia || (Number(x) > ultimo && dia === ultimo));
    return diasSet_(f.dias).indexOf(dow) > -1;
  });
}

function excluido_(excepto, ymd) { return String(excepto || '').split(/[,\s]+/).indexOf(ymd) > -1; }

/** "AED · Edif 1 · Salón 111" a partir del texto de la celda del horario. */
function infoClase_(txt) {
  const n = String(txt || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const SHORT = { 'Algoritmos y Estructura de Datos': 'AED', 'Programación para Ciencia de Datos': 'PCD', 'Introducción a la Ciencia de Datos': 'ICD', 'Ética y Legalidad': 'Ética' };
  let titulo = '';
  AMBITOS[AMBITO_DP].categorias.forEach(c => {
    if (titulo) return;
    const toks = c.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').split(/\s+/).filter(w => w.length > 3);
    if (toks.length && toks.every(w => n.indexOf(w.slice(0, 6)) > -1)) titulo = SHORT[c] || c;
  });
  if (!titulo) titulo = String(txt).split('\n').map(l => l.trim()).filter(l => l && !/^\d[A-Z]{2}\d$/.test(l))[0] || 'Clase';
  titulo = titulo.split(/\s{2,}/).filter(Boolean).join(' · ');
  const loc = String(txt).match(/Edif:?\s*(\w+)\s*-\s*Sal[oó]n:?\s*(\w+)/i);
  return { titulo: titulo, lugar: loc ? 'Edif ' + loc[1] + ' · Salón ' + loc[2] : '' };
}

/** Manda por Formspree (si está configurado y hay cupo) o por CallMeBot. */
function enviarAviso_(text) {
  const props = PropertiesService.getScriptProperties();
  const endpoint = props.getProperty('FORMSPREE_ENDPOINT');
  if (endpoint) {
    const mes = Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), limite = Number(props.getProperty('FORMSPREE_LIMITE') || 50);
    const key = 'FORMSPREE_USO_' + mes, uso = Number(props.getProperty(key) || 0);
    if (uso < limite) {
      const res = UrlFetchApp.fetch(endpoint, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Accept: 'application/json' },
        payload: JSON.stringify({ _subject: 'Dame la Vida · ' + text.split('\n')[0].slice(0, 60), message: text }) });
      Logger.log('Formspree respondió ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
      if (res.getResponseCode() < 300) { props.setProperty(key, String(uso + 1)); return; }
    } else Logger.log('Formspree: se alcanzó el límite de ' + limite + ' envíos este mes.');
  }
  if (props.getProperty('CALLMEBOT_APIKEY')) return sendWhatsApp_(text);
  if (!endpoint) throw new Error('Configura FORMSPREE_ENDPOINT o CALLMEBOT_APIKEY en las propiedades del script.');
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

/* --- fechas en hora de CDMX --- */
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ymdMas_(ymd, n) { const p = ymd.split('-').map(Number); return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2] + n, 12), TZ, 'yyyy-MM-dd'); }
function dowDe_(ymd) { const p = ymd.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12).getDay(); }
function minutos_(hhmm) { const m = String(hhmm || '').match(/^(\d{1,2}):(\d{2})/); return m ? Number(m[1]) * 60 + Number(m[2]) : null; }
function hhmmDe_(m) { return String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0'); }
/** Milisegundos de "ymd HH:mm" en CDMX (UTC-6 todo el año desde 2022). */
function msDe_(ymd, hhmm) { const p = ymd.split('-').map(Number), m = minutos_(hhmm); return Date.UTC(p[0], p[1] - 1, p[2], 6, 0) + m * 60000; }

/* Compatibilidad con la versión anterior: si aún tienes el disparador viejo, manda el resumen. */
function avisoDiario() { enviarAviso_(resumenDelDia_(hoy_())); }
function instalarAvisoDiario() { instalarRecordatorios(); }
function quitarAvisoDiario() { quitarRecordatorios(); }

/** Bloques de la pestaña Plan que tocan en una fecha (yyyy-MM-dd), ordenados por hora. */
function planForDate_(plan, ymd) {
  const p = ymd.split('-').map(Number);
  const dow = new Date(p[0], p[1] - 1, p[2]).getDay();
  return plan.filter(b => (!b.desde || b.desde <= ymd) && (!b.hasta || ymd <= b.hasta) && !excluido_(b.excepto, ymd) && diasSet_(b.dias).indexOf(dow) > -1)
    .sort((a, b) => a.inicio < b.inicio ? -1 : 1);
}

/** "Diario", "L-V", "L,M,X" → [0..6] (0 = domingo). */
function diasSet_(txt) {
  const s = String(txt || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  if (!s.trim() || /DIARIO|TODOS/.test(s)) return [0, 1, 2, 3, 4, 5, 6];
  const map = { L: 1, M: 2, X: 3, J: 4, V: 5, S: 6, D: 0 }, order = ['L', 'M', 'X', 'J', 'V', 'S', 'D'], out = [];
  const words = { LUN: 'L', MAR: 'M', MIE: 'X', JUE: 'J', VIE: 'V', SAB: 'S', DOM: 'D' };
  s.split(/[,\s/]+/).map(t => words[t.slice(0, 3)] && t.length > 1 && t.indexOf('-') === -1 ? words[t.slice(0, 3)] : t).forEach(tok => {
    const r = tok.match(/^([LMXJVSD])-([LMXJVSD])$/);
    if (r) { for (let i = order.indexOf(r[1]); i <= order.indexOf(r[2]); i++) out.push(map[order[i]]); }
    else if (map[tok] !== undefined) out.push(map[tok]);
  });
  return out;
}

/** Tu número con lada de país (52 + 10 dígitos), leído de las propiedades del script. */
function myPhone_() {
  let d = String(PropertiesService.getScriptProperties().getProperty('MI_WHATSAPP') || '').replace(/\D/g, '');
  if (d.length === 10) d = '52' + d;
  return d.length >= 11 ? d : '';
}

/* ================== SINCRONIZACIÓN CON GOOGLE CALENDAR ================== */
// Mantiene un calendario aparte, «Dame la Vida», igual a tu hoja: clases, rutina fija, bloques del Plan,
// fechas límite de las asignaciones y calendario del IPN, con recordatorios en el celular.
// La hoja manda: lo que cambies aquí se corrige solo en el calendario (cada día a las 5 a.m., y unos
// minutos después de guardar una asignación desde el tablero). Los eventos se guardan en la pestaña
// «Sync Calendario» para saber cuáles son tuyos; tu calendario personal nunca se toca.
//   instalarSincronizacion()   activa todo y hace la primera carga
//   simularSincronizacion()    solo cuenta qué crearía/cambiaría/borraría (no toca el calendario)
//   sincronizarCalendario()    sincroniza ahora
//   quitarSincronizacion()     apaga la sincronización (el calendario se queda como está)
//   borrarCalendario()         borra el calendario «Dame la Vida» completo

const CAL_NOMBRE = 'Dame la Vida';
const SHEET_SYNC = 'Sync Calendario';
const CAL_MESES_ADELANTE = 3;        // hasta qué mes se crean las repeticiones que tienen fin y los pagos
const CAL_HORA_SYNC = 5;             // sincronización diaria (a.m.)
const CAL_TIEMPO_MAX_MS = 270000;    // Apps Script corta a los 6 min: si falta, sigue en otra corrida
const CAL_PAUSA_MS = 80;
// Minutos antes del inicio. 900 = un día antes a las 9:00 (para eventos de todo el día).
const CAL_AVISOS = {
  clase: [15], horario: [15], plan: [15], deporte: [15], traslado: [15], hogar: [15],
  medicamento: [15, 0], comida: [0], rutina: [0], pago: [900],
  tarea: [900],            // fecha límite sin hora
  tareaHora: [1440, 60],   // con hora (citas, exámenes): un día antes y 1 h antes
  ipn: [900]
};
const CAL_COLORES = {
  clase: 'BLUE', horario: 'BLUE', deporte: 'BLUE', comida: 'PALE_GREEN', medicamento: 'MAUVE', rutina: 'GRAY',
  traslado: 'GRAY', hogar: 'CYAN', pago: 'YELLOW', ipn: 'GRAY', Urgente: 'RED', Importante: 'ORANGE', Normal: 'GREEN'
};
const CAL_DOW = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];

/** Activa la sincronización diaria y hace la primera carga. Ejecútala una vez. */
function instalarSincronizacion() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sincronizarCalendario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizarCalendario').timeBased().everyDays(1).atHour(CAL_HORA_SYNC).create();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CAL_SYNC_ACTIVA', '1');
  props.deleteProperty('CAL_SYNC_PENDIENTE');
  Logger.log('Sincronización diaria activa (≈' + CAL_HORA_SYNC + ' a.m.). Los avisos de WhatsApp pasan a modo «resumen» (propiedad AVISOS_WHATSAPP = todo para volver a todos).');
  sincronizarCalendario();
}

function quitarSincronizacion() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sincronizarCalendario') ScriptApp.deleteTrigger(t); });
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('CAL_SYNC_ACTIVA'); props.deleteProperty('CAL_SYNC_PENDIENTE');
  Logger.log('Sincronización apagada. Los eventos ya creados siguen en el calendario «' + CAL_NOMBRE + '».');
}

/** Borra el calendario «Dame la Vida» con todos sus eventos y limpia la pestaña de control. */
function borrarCalendario() {
  quitarSincronizacion();
  const cals = CalendarApp.getCalendarsByName(CAL_NOMBRE);
  cals.forEach(c => c.deleteCalendar());
  const id = PropertiesService.getScriptProperties().getProperty('CAL_ID');
  if (id) { const c = CalendarApp.getCalendarById(id); if (c) c.deleteCalendar(); }
  PropertiesService.getScriptProperties().deleteProperty('CAL_ID');
  PropertiesService.getScriptProperties().deleteProperty('CAL_SYNC_INICIO');
  syncGuardar_({});
  Logger.log('Calendario «' + CAL_NOMBRE + '» borrado.');
}

function sincronizarCalendario() { return sincronizar_(false); }
function simularSincronizacion() { return sincronizar_(true); }

function sincronizar_(simular) {
  const props = PropertiesService.getScriptProperties(), t0 = Date.now();
  const enCurso = Number(props.getProperty('CAL_SYNC_EN_CURSO') || 0);
  if (!simular && enCurso && t0 - enCurso < 7 * 60000) { Logger.log('Ya hay una sincronización en curso.'); return null; }
  if (!simular) props.setProperty('CAL_SYNC_EN_CURSO', String(t0));
  try {
    const pend = props.getProperty('CAL_SYNC_PENDIENTE');
    if (pend && !simular) {
      ScriptApp.getProjectTriggers().forEach(t => { if (t.getUniqueId() === pend) ScriptApp.deleteTrigger(t); });
      props.deleteProperty('CAL_SYNC_PENDIENTE');
    }
    let mapa = syncLeer_(), cal = null;
    if (!simular) {
      const c = calCalendario_();
      cal = c.cal;
      if (c.nuevo) mapa = {};   // el calendario se creó de cero: nada de lo anotado existe ya
    }
    const deseados = {};
    calSpecs_(!simular).forEach(s => { deseados[s.key] = s; });

    const borrar = [], crear = [];
    Object.keys(mapa).forEach(k => { if (!deseados[k]) borrar.push(k); });
    Object.keys(deseados).forEach(k => {
      if (!mapa[k]) crear.push(k);
      else if (mapa[k].hash !== deseados[k].hash) { borrar.push(k); crear.push(k); }
    });
    const cambian = crear.filter(k => mapa[k]).length;
    const resumen = 'Crear ' + (crear.length - cambian) + ' · cambiar ' + cambian + ' · borrar ' + (borrar.length - cambian) +
      ' · sin cambios ' + (Object.keys(deseados).length - crear.length);
    if (simular) {
      Logger.log('Simulación: ' + resumen);
      crear.slice(0, 40).forEach(k => Logger.log('  + ' + deseados[k].titulo + ' (' + deseados[k].kind + ')'));
      if (crear.length > 40) Logger.log('  … y ' + (crear.length - 40) + ' más');
      return { crear: crear.length, borrar: borrar.length, cambian: cambian };
    }

    let hechos = 0, parado = false;
    const sinTiempo = () => Date.now() - t0 > CAL_TIEMPO_MAX_MS;
    for (const k of borrar) {
      if (sinTiempo()) { parado = true; break; }
      calBorrar_(cal, mapa[k]); delete mapa[k]; hechos++; Utilities.sleep(CAL_PAUSA_MS);
    }
    if (!parado) for (const k of crear) {
      if (sinTiempo()) { parado = true; break; }
      const s = deseados[k];
      try {
        mapa[k] = { hash: s.hash, kind: s.kind, id: calCrear_(cal, s), titulo: s.titulo };
      } catch (err) { Logger.log('No pude crear «' + s.titulo + '»: ' + err); }
      hechos++; Utilities.sleep(CAL_PAUSA_MS);
      if (hechos % 25 === 0) syncGuardar_(mapa);
    }
    syncGuardar_(mapa);
    Logger.log('Sincronizado: ' + resumen + (parado ? ' (parcial: sigo en un minuto)' : '') + '.');
    if (parado) programarSync_(1);
    return { creados: crear.length, parcial: parado };
  } finally { if (!simular) props.deleteProperty('CAL_SYNC_EN_CURSO'); }
}

/** Calendario «Dame la Vida»: lo busca o lo crea. nuevo = true si hubo que crearlo. */
function calCalendario_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CAL_ID');
  let cal = id ? CalendarApp.getCalendarById(id) : null, nuevo = false;
  if (!cal) cal = CalendarApp.getCalendarsByName(CAL_NOMBRE)[0] || null;
  if (!cal) {
    cal = CalendarApp.createCalendar(CAL_NOMBRE, { summary: 'Clases, rutina, estudio y fechas del tablero «Dame la Vida». Se actualiza solo desde la hoja.', timeZone: TZ });
    nuevo = true;
  }
  try { cal.setTimeZone(TZ); cal.setSelected(true); cal.setHidden(false); } catch (err) { Logger.log(err); }
  if (nuevo) { try { cal.setColor('#4986E7'); } catch (err) { Logger.log(err); } }
  props.setProperty('CAL_ID', cal.getId());
  return { cal: cal, nuevo: nuevo };
}

function calCrear_(cal, s) {
  const opts = { description: s.desc };
  if (s.lugar) opts.location = s.lugar;
  const dia = ymd => new Date(msDe_(ymd, '12:00'));
  let ev;
  if (s.kind === 'serie') {
    const rec = CalendarApp.newRecurrence();
    const regla = rec.addWeeklyRule().onlyOnWeekdays(s.dias.map(d => CalendarApp.Weekday[CAL_DOW[d]]));
    if (s.veces) regla.times(s.veces);
    ev = cal.createEventSeries(s.titulo, new Date(s.ini), new Date(s.fin), rec, opts);
  } else if (s.kind === 'evento') {
    ev = cal.createEvent(s.titulo, new Date(s.ini), new Date(s.fin), opts);
  } else if (s.ymdFin && s.ymdFin > s.ymd) {
    ev = cal.createAllDayEvent(s.titulo, dia(s.ymd), dia(ymdMas_(s.ymdFin, 1)), opts);   // el fin es exclusivo
  } else {
    ev = cal.createAllDayEvent(s.titulo, dia(s.ymd), opts);
  }
  ev.removeAllReminders();
  s.avisos.forEach(m => ev.addPopupReminder(m));
  if (s.color && CalendarApp.EventColor[s.color]) ev.setColor(CalendarApp.EventColor[s.color]);
  return ev.getId();
}

function calBorrar_(cal, e) {
  try {
    if (e.kind === 'serie') { const s = cal.getEventSeriesById(e.id); if (s) s.deleteEventSeries(); }
    else { const v = cal.getEventById(e.id); if (v) v.deleteEvent(); }
  } catch (err) { Logger.log('No pude borrar «' + e.titulo + '»: ' + err); }
}

/** Tras guardar, mover o borrar una asignación: el calendario se pone al día en un par de minutos. */
function avisarCambio_() { try { programarSync_(2); } catch (err) { Logger.log(err); } }

/** Programa una sincronización en unos minutos (una sola a la vez). Solo si la sincronización está activa. */
function programarSync_(minutos) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CAL_SYNC_ACTIVA')) return;
  const id = props.getProperty('CAL_SYNC_PENDIENTE');
  if (id && ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === id)) return;
  const t = ScriptApp.newTrigger('sincronizarCalendario').timeBased().after(minutos * 60000).create();
  props.setProperty('CAL_SYNC_PENDIENTE', t.getUniqueId());
}

/* --- pestaña de control: Clave · Huella · Tipo · ID · Título --- */
function syncLeer_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SYNC);
  const mapa = {};
  if (!sh || sh.getLastRow() < 2) return mapa;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    if (r[0] && r[3]) mapa[String(r[0])] = { hash: String(r[1]), kind: String(r[2]), id: String(r[3]), titulo: String(r[4]) };
  });
  return mapa;
}

function syncGuardar_(mapa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureSheet_(ss, SHEET_SYNC, ['Clave', 'Huella', 'Tipo', 'ID', 'Título']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  const filas = Object.keys(mapa).sort().map(k => [k, mapa[k].hash, mapa[k].kind, mapa[k].id, clean_(mapa[k].titulo)]);
  if (filas.length) sh.getRange(2, 1, filas.length, 5).setValues(filas);
}

/** Primer día que se sincroniza: se fija en la primera corrida (no se crean eventos del pasado). */
function calInicio_(guardar) {
  const props = PropertiesService.getScriptProperties();
  let d = props.getProperty('CAL_SYNC_INICIO');
  if (!d) { d = hoy_(); if (guardar) props.setProperty('CAL_SYNC_INICIO', d); }
  return d;
}

function calHash_(obj) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(obj), Utilities.Charset.UTF_8);
  return 'h' + bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

/** Todo lo que debe existir en el calendario: [{key, kind: serie|evento|dia, titulo, desc, lugar, color, avisos, hash, …}]. */
function calSpecs_(guardarInicio) {
  const hoy = hoy_(), inicio = calInicio_(guardarInicio);
  const p = hoy.split('-').map(Number);
  const finMes = new Date(p[0], p[1] - 1 + CAL_MESES_ADELANTE + 1, 0, 12);      // último día del mes de adelante
  const fin = Utilities.formatDate(finMes, TZ, 'yyyy-MM-dd');
  const horario = optional_(getHorario_), plan = optional_(getPlan_), fijas = optional_(getFijas_), ipn = optional_(getCalendarioIPN_);
  const tareas = readTasks_(), porId = {};
  tareas.forEach(t => porId[t.id] = t);
  const hecha = id => !!(porId[id] && porId[id].estado === 'Done');
  const tope = ymdMas_(inicio, 400);                      // las repeticiones con "Hasta" se calculan completas
  let finMax = fin;
  fijas.concat(plan).forEach(x => { if (x.hasta && x.hasta > finMax) finMax = x.hasta < tope ? x.hasta : tope; });

  // 1. En qué fechas cae cada cosa (la misma lógica que el tablero y los avisos)
  const fechasDe = {}, push = (k, d) => (fechasDe[k] = fechasDe[k] || []).push(d);
  horario.forEach(h => { h.ci = infoClase_(h.texto); h.k = 'h|' + h.dia + '|' + h.inicio + '|' + h.fin + '|' + h.ci.titulo; });
  plan.forEach(b => { b.k = 'p|' + b.id + '|' + b.actividad + '|' + b.dias + '|' + b.inicio + '|' + b.fin + '|' + b.desde; });
  fijas.forEach(f => { f.k = 'f|' + f.actividad + '|' + f.dias + '|' + f.diaMes + '|' + f.inicio + '|' + f.desde; });
  for (let d = inicio; d <= finMax; d = ymdMas_(d, 1)) {
    const clases = hayClasesEn_(ipn, d), dow = dowDe_(d);
    horario.forEach(h => { if (clases && h.dia === dow) push(h.k, d); });
    fijasDelDia_(fijas, d).filter(f => clases || !/^clase$/i.test(f.tipo)).forEach(f => push(f.k, d));
    planForDate_(plan, d).forEach(b => { if (!hecha(b.id)) push(b.k, d); });
  }

  const specs = [], vistas = {};
  const add = s => {
    let k = s.key, n = 2;
    while (vistas[k]) k = s.key + '~' + n++;           // dos filas idénticas no se pisan
    vistas[k] = 1; s.key = k;
    const base = Object.assign({}, s); delete base.key;
    s.hash = calHash_(base); specs.push(s);
  };
  const pie = '\n\n— Se actualiza solo desde tu tablero «Dame la Vida»; cámbialo en la hoja, no aquí.';
  const norm = v => String(v || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  const hm = (ymd, hhmm) => msDe_(ymd, hhmm);

  // Una repetición con huecos se parte en tramos continuos; cada tramo es una serie de Calendar.
  const series = (origen, wset, abierta, mk) => {
    if (!wset.length) return;
    let fechas = (fechasDe[origen.k] || []).slice();
    if (!origen.hasta) fechas = fechas.filter(d => d <= fin);   // sin "Hasta": solo hasta el horizonte (la serie queda abierta)
    if (!fechas.length) return;
    const set = {}; fechas.forEach(d => set[d] = 1);
    const ultimaPosible = (() => { let d = fin; while (wset.indexOf(dowDe_(d)) === -1) d = ymdMas_(d, -1); return d; })();
    const tramos = []; let cur = null;
    for (let d = fechas[0]; d <= fechas[fechas.length - 1]; d = ymdMas_(d, 1)) {
      if (wset.indexOf(dowDe_(d)) === -1) continue;
      if (set[d]) { if (!cur) { cur = []; tramos.push(cur); } cur.push(d); } else cur = null;
    }
    tramos.forEach(t => {
      const abiertaT = abierta && !origen.hasta && t[t.length - 1] === ultimaPosible;
      mk(t, abiertaT);
    });
  };

  // 2. Clases del horario
  horario.forEach(h => {
    const lineas = String(h.texto).split('\n').map(l => l.trim()).filter(Boolean);
    series(h, [h.dia], true, (t, abierta) => emitirTramo(add, { key: h.k, tramo: t, abierta: abierta, titulo: 'Clase · ' + h.ci.titulo,
      desc: lineas.join('\n') + pie, lugar: h.ci.lugar, inicio: h.inicio, fin: h.fin, tipo: 'clase', avisos: CAL_AVISOS.clase, color: CAL_COLORES.clase, dias: [h.dia] }));
  });

  // 3. Rutina fija (pestaña Fijas)
  fijas.forEach(f => {
    const tipo = norm(f.tipo);
    const sinAviso = tipo !== 'pago' && (f.aviso === '' || isNaN(Number(f.aviso)));   // columna «Aviso» vacía = sin recordatorio
    const avisos = sinAviso ? [] : CAL_AVISOS[tipo] || [Number(f.aviso)];
    const color = CAL_COLORES[tipo] || 'GRAY';
    const desc = [f.notas, f.lugar && f.lugar !== '' ? 'Lugar: ' + f.lugar : ''].filter(Boolean).join('\n') + pie;
    const dMin = f.inicio ? (f.fin ? ((minutos_(f.fin) - minutos_(f.inicio)) + 1440) % 1440 || 15 : 15) : 0;
    if (f.diaMes || !f.inicio) {
      // pagos y cosas de todo el día: un evento por fecha (el calendario se va llenando mes a mes)
      (fechasDe[f.k] || []).filter(d => d <= fin).forEach(d =>
        add(f.inicio
          ? { key: f.k + '#' + d, kind: 'evento', titulo: f.actividad, desc: desc, lugar: f.lugar, color: color, avisos: avisos, ini: hm(d, f.inicio), fin: hm(d, f.inicio) + dMin * 60000 }
          : { key: f.k + '#' + d, kind: 'dia', titulo: f.actividad, desc: desc, lugar: f.lugar, color: color, avisos: avisos, ymd: d }));
      return;
    }
    series(f, diasSet_(f.dias), true, (t, abierta) => emitirTramo(add, { key: f.k, tramo: t, abierta: abierta, titulo: f.actividad, desc: desc,
      lugar: f.lugar, inicio: f.inicio, fin: null, dur: dMin, tipo: tipo, avisos: avisos, color: color, dias: diasSet_(f.dias) }));
  });

  // 4. Bloques de estudio (pestaña Plan): color por la urgencia de la asignación
  plan.forEach(b => {
    const t = porId[b.id];
    const desc = [b.notas, t ? t.tipo + ' · ' + (t.categoria || '') + (t.prioridad ? ' · ' + t.prioridad : '') : '', t && t.materiales ? 'Recursos:\n' + t.materiales : '', b.id ? 'Asignación ' + b.id : ''].filter(Boolean).join('\n') + pie;
    series(b, diasSet_(b.dias), true, (tr, abierta) => emitirTramo(add, { key: b.k, tramo: tr, abierta: abierta, titulo: b.actividad, desc: desc, lugar: '',
      inicio: b.inicio, fin: b.fin, tipo: 'plan', avisos: CAL_AVISOS.plan, color: CAL_COLORES[t ? t.prioridad : 'Normal'] || 'GREEN', dias: diasSet_(b.dias) }));
  });

  // 5. Asignaciones con fecha límite
  tareas.filter(t => t.estado !== 'Done' && t.fecha && t.fecha >= inicio).forEach(t => {
    const desc = [t.descripcion, t.enlace ? 'Entregable: ' + t.enlace : '', t.materiales ? 'Recursos:\n' + t.materiales : '', t.notas, t.id].filter(Boolean).join('\n\n') + pie;
    const titulo = t.tipo + ': ' + t.titulo, color = CAL_COLORES[t.prioridad] || 'GREEN';
    if (t.hora) add({ key: 't|' + t.id, kind: 'evento', titulo: titulo, desc: desc, lugar: '', color: color, avisos: CAL_AVISOS.tareaHora, ini: hm(t.fecha, t.hora), fin: hm(t.fecha, t.hora) + 3600000 });
    else add({ key: 't|' + t.id, kind: 'dia', titulo: titulo, desc: desc, lugar: '', color: color, avisos: CAL_AVISOS.tarea, ymd: t.fecha });
  });

  // 6. Calendario del IPN (el rango de «Clases» solo sirve para saber cuándo hay clases)
  ipn.filter(e => !/^clases$/i.test(e.tipo) && (e.fin || e.inicio) >= inicio).forEach(e => {
    add({ key: 'i|' + e.evento + '|' + e.inicio, kind: 'dia', titulo: e.evento, desc: [e.tipo, e.notas].filter(Boolean).join(' · ') + pie, lugar: '',
      color: CAL_COLORES.ipn, avisos: CAL_AVISOS.ipn, ymd: e.inicio, ymdFin: e.fin });
  });
  return specs;
}

/** Convierte un tramo de fechas en una serie (o en un evento suelto si es una sola fecha). */
function emitirTramo(add, o) {
  const t = o.tramo, d0 = t[0];
  const dur = o.dur || (((minutos_(o.fin) - minutos_(o.inicio)) + 1440) % 1440 || 15);
  const ini = msDe_(d0, o.inicio), fin = ini + dur * 60000;
  const comun = { titulo: o.titulo, desc: o.desc, lugar: o.lugar, color: o.color, avisos: o.avisos };
  if (t.length === 1 && !o.abierta) add(Object.assign({ key: o.key + '#' + d0, kind: 'evento', ini: ini, fin: fin }, comun));
  else add(Object.assign({ key: o.key + '#' + d0, kind: 'serie', ini: ini, fin: fin, dias: o.dias.slice().sort(),
    veces: o.abierta ? 0 : t.length }, comun));
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

/**
 * Horario de clases (pestaña "Horario", tal como la armaste): una columna de horas "07:00 - 08:30"
 * y una columna por día (Lunes … Domingo). Devuelve [{dia, inicio, fin, texto}] con dia 0 = domingo.
 */
function getHorario_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HORARIO);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const norm = s => String(s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const DIAS = { domingo: 0, lunes: 1, martes: 2, miercoles: 3, jueves: 4, viernes: 5, sabado: 6 };
  const hr = data.findIndex(r => r.some(c => DIAS[norm(c)] !== undefined));
  if (hr === -1) return [];
  const head = data[hr].map(norm);
  const horaCol = head.findIndex(h => /^hora/.test(h));
  const out = [];
  for (let i = hr + 1; i < data.length; i++) {
    const m = String(data[i][horaCol] || '').match(/(\d{1,2}):(\d{2})\s*[-–a]+\s*(\d{1,2}):(\d{2})/);
    if (!m) continue;
    const inicio = m[1].padStart(2, '0') + ':' + m[2], fin = m[3].padStart(2, '0') + ':' + m[4];
    head.forEach((h, c) => {
      const txt = String(data[i][c] || '').trim();
      if (DIAS[h] !== undefined && txt) out.push({ dia: DIAS[h], inicio: inicio, fin: fin, texto: txt });
    });
  }
  return out;
}

/** Bloques de estudio (pestaña "Plan": ID · Actividad · Días · Inicio · Fin · Desde · Hasta · Notas). */
function getPlan_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAN);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .filter(r => String(r[1]).trim() && fmtTime_(r[3]) && fmtTime_(r[4]))
    .map(r => ({ id: String(r[0] || '').trim(), actividad: String(r[1]).trim(), dias: String(r[2] || '').trim(),
                 inicio: fmtTime_(r[3]), fin: fmtTime_(r[4]), desde: fmtDay_(r[5]), hasta: fmtDay_(r[6]),
                 notas: String(r[7] || '').trim(), excepto: fmtDays_(r[8]) }));
}

/**
 * Actividades fijas (pestaña "Fijas"): Actividad · Tipo · Días · Día del mes · Inicio · Fin · Desde · Hasta ·
 * Excepto · Lugar · Notas · Aviso (min antes). Sin Inicio = todo el día (pagos). Sin Fin = marca (comidas).
 * Tipo "Clase" se oculta en días sin clases, igual que el horario.
 */
function getFijas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FIJAS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues()
    .filter(r => String(r[0]).trim())
    .map((r, i) => ({ fila: i + 2, actividad: String(r[0]).trim(), tipo: String(r[1] || '').trim(), dias: String(r[2] || '').trim(),
                      diaMes: String(r[3] || '').trim(), inicio: fmtTime_(r[4]), fin: fmtTime_(r[5]),
                      desde: fmtDay_(r[6]), hasta: fmtDay_(r[7]), excepto: fmtDays_(r[8]), lugar: String(r[9] || '').trim(),
                      notas: String(r[10] || '').trim(), aviso: r[11] === '' || r[11] === null ? '' : String(r[11]).trim() }));
}

/** Celda con una o varias fechas (Date o "2026-11-14, 2026-11-21") → "2026-11-14, 2026-11-21". */
function fmtDays_(v) {
  if (v instanceof Date) return fmtDay_(v);
  if (typeof v === 'number' && v > 30000) return fmtDay_(new Date(Math.round((v - 25569) * 86400000) + 12 * 3600000)); // fecha guardada como número de serie
  return (String(v || '').match(/\d{4}-\d{2}-\d{2}/g) || []).join(', ');
}

/** Calendario académico (pestaña "Calendario IPN": Evento · Inicio · Fin · Tipo · Notas). */
function getCalendarioIPN_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CAL_IPN);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0]).trim() && fmtDay_(r[1]))
    .map(r => ({ evento: String(r[0]).trim(), inicio: fmtDay_(r[1]), fin: fmtDay_(r[2]) || fmtDay_(r[1]),
                 tipo: String(r[3] || '').trim(), notas: String(r[4] || '').trim() }));
}

function fmtDay_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}
function fmtTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = String(v || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, '0') + ':' + m[2] : '';
}

/** "AWS, Udemy" → máximo dos escuelas/plataformas, sin repetir. */
function cleanInstituciones_(v) {
  const list = (Array.isArray(v) ? v : String(v || '').split(','))
    .map(s => String(s).replace(/[=+@]/g, '').trim()).filter(Boolean);
  return list.filter((s, i) => list.indexOf(s) === i).slice(0, MAX_INSTITUCIONES).map(s => s.slice(0, 40)).join(', ');
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

/** Siguiente ID con el prefijo del ámbito: DP-001, PER-001… */
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
    if (key === 'fecha' || key === 'inicio') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
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
