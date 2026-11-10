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
 * Recordatorios a lo largo del día: ejecuta instalarRecordatorios() una vez (ver README).
 */

const SHEET_TAREAS = 'Tareas';
const SHEET_EQUIPO = 'Equipo';
const SHEET_RECURSOS = 'Recursos';
const SHEET_HORARIO = 'Horario';
const SHEET_PLAN = 'Plan';
const SHEET_CAL_IPN = 'Calendario IPN';
const SHEET_FIJAS = 'Fijas';
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
      default: throw new Error('Acción no válida: ' + body.action);
    }
  }));
}

/* ============================== OPERACIONES ============================== */

function getBoard_() {
  return { ok: true, tasks: readTasks_(), team: getTeam_(), recursos: getRecursos_(), ambitos: AMBITOS,
           horario: optional_(getHorario_), plan: optional_(getPlan_), calendario: optional_(getCalendarioIPN_),
           fijas: optional_(getFijas_),
           estados: ESTADOS, prioridades: PRIORIDADES, miWhatsapp: myPhone_(), serverTime: now_() };
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
  const add = (hhmm, minAntes, texto) => {
    const m = minutos_(hhmm); if (m === null) return;
    out.push({ aviso: hhmmDe_(Math.max(0, m - minAntes)), texto: texto });
  };
  const ev = eventosDelDia_(ymd);
  ev.clases.forEach(c => add(c.inicio, AVISO_CLASE_MIN, 'En ' + AVISO_CLASE_MIN + ' min: ' + c.titulo + (c.lugar ? ' · ' + c.lugar : '') + ' (' + c.inicio + '–' + c.fin + ')'));
  ev.fijas.forEach(f => {
    if (f.aviso === '' || !f.inicio) return;
    const antes = Number(f.aviso) || 0;
    let txt = (antes ? 'En ' + antes + ' min: ' : '') + f.actividad + (f.fin ? ' (' + f.inicio + '–' + f.fin + ')' : ' · ' + f.inicio) +
      (f.lugar ? '\n' + f.lugar : '') + (f.notas ? '\n' + f.notas : '');
    if (!resumenEn.hecho && /despert/i.test(f.actividad)) { txt = '¡Buenos días! ' + txt + '\n\n' + resumenDelDia_(ymd); resumenEn.hecho = true; }
    add(f.inicio, antes, txt);
  });
  ev.plan.forEach(b => add(b.inicio, AVISO_PLAN_MIN, 'Ahora: ' + b.actividad + ' hasta las ' + b.fin + (b.notas ? '\n' + b.notas : '')));
  ev.tareas.filter(t => t.hora).forEach(t => add(t.hora, AVISO_TAREA_MIN, 'En 1 h: ' + t.tipo + ' · ' + t.titulo + ' (' + t.hora + ')' + (t.descripcion ? '\n' + t.descripcion : '')));
  if (!resumenEn.hecho) out.push({ aviso: '07:00', texto: resumenDelDia_(ymd) });
  return out.sort((a, b) => a.aviso < b.aviso ? -1 : 1);
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
  return out.join('\n');
}

/** Clases, fijas, bloques del plan y tareas con fecha de un día. */
function eventosDelDia_(ymd) {
  const cal = optional_(getCalendarioIPN_);
  const enRango = e => e.inicio <= ymd && ymd <= (e.fin || e.inicio);
  const libre = cal.some(e => enRango(e) && /descanso|vacaci/i.test(e.tipo));
  const periodos = cal.filter(e => /^clases$/i.test(e.tipo));
  const hayClases = !libre && (!periodos.length || periodos.some(enRango));
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
