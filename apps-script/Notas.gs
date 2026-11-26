/* ============================ NOTAS (estilo Keep) ============================
 * Vista "Notas" del tablero: tarjetas de texto o checklist, con color, que viven
 * en el Sheet (pestañas "Notas" y "NotaItems"). No tocan el calendario, el tablero
 * ni el pomodoro: son solo notas.
 *
 *   instalarNotas()   crea las pestañas "Notas" y "NotaItems" (ejecútala UNA vez)
 *   sembrarLibros()   carga tus libros leídos / en curso / por leer (ejecútala UNA vez)
 *
 * Una nota checklist puede tener "Mover a" = id de otra nota: al marcar un ítem como
 * 'hecho', ese ítem se pasa solo a la nota destino y sale de la lista (libros, cursos…).
 * =========================================================================== */

const SHEET_NOTAS = 'Notas';
const SHEET_NOTA_ITEMS = 'NotaItems';
const NOTAS_HEADERS = ['ID', 'Título', 'Tipo', 'Color', 'Cuerpo', 'Orden', 'Mover a', 'Archivada', 'Creado', 'Actualizado'];
const NOTA_ITEMS_HEADERS = ['ID', 'Nota', 'Texto', 'Estado', 'Nota ítem', 'Orden', 'Creado', 'Actualizado'];
const NOTA_TIPOS = ['texto', 'checklist'];
const NOTA_COLORES = ['default', 'rojo', 'naranja', 'amarillo', 'verde', 'azul', 'morado', 'rosa', 'gris'];
const NOTA_ITEM_ESTADOS = ['pendiente', 'en_curso', 'hecho'];

/** Crea las pestañas de notas si no existen. Ejecútala una vez. */
function instalarNotas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_NOTAS, NOTAS_HEADERS);
  ensureSheet_(ss, SHEET_NOTA_ITEMS, NOTA_ITEMS_HEADERS);
  Logger.log('Listo: pestañas "' + SHEET_NOTAS + '" y "' + SHEET_NOTA_ITEMS + '" creadas. Ahora puedes ejecutar sembrarLibros().');
}

/** Lee todas las notas (no archivadas) con sus ítems, ordenadas. */
function getNotas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shN = ss.getSheetByName(SHEET_NOTAS);
  if (!shN || shN.getLastRow() < 2) return [];
  const notas = shN.getRange(2, 1, shN.getLastRow() - 1, NOTAS_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      id: String(r[0]).trim(), titulo: String(r[1] || ''),
      tipo: NOTA_TIPOS.indexOf(String(r[2] || '').trim().toLowerCase()) > -1 ? String(r[2]).trim().toLowerCase() : 'texto',
      color: String(r[3] || 'default').trim().toLowerCase(), cuerpo: String(r[4] || ''),
      orden: Number(r[5]) || 0, moverA: String(r[6] || '').trim(), archivada: toBool_(r[7]), items: []
    }))
    .filter(n => !n.archivada);
  const byId = {}; notas.forEach(n => (byId[n.id] = n));
  const shI = ss.getSheetByName(SHEET_NOTA_ITEMS);
  if (shI && shI.getLastRow() >= 2) {
    shI.getRange(2, 1, shI.getLastRow() - 1, NOTA_ITEMS_HEADERS.length).getValues()
      .filter(r => String(r[0]).trim() && byId[String(r[1]).trim()])
      .forEach(r => byId[String(r[1]).trim()].items.push({
        id: String(r[0]).trim(), texto: String(r[2] || ''),
        estado: NOTA_ITEM_ESTADOS.indexOf(String(r[3] || '').trim().toLowerCase()) > -1 ? String(r[3]).trim().toLowerCase() : 'pendiente',
        nota: String(r[4] || ''), orden: Number(r[5]) || 0
      }));
  }
  notas.forEach(n => n.items.sort((a, b) => a.orden - b.orden));
  notas.sort((a, b) => a.orden - b.orden);
  return notas;
}

/** Dispatcher de las acciones de notas (lo llama doPost en el caso default). */
function notasAccion_(body) {
  switch (body.action) {
    case 'notaSave':       return notaSave_(body.nota);
    case 'notaDelete':     return notaDelete_(body.id);
    case 'notaOrden':      return notaOrden_(body.ids);
    case 'notaItemSave':   return notaItemSave_(body.item);
    case 'notaItemEstado': return notaItemEstado_(body.id, body.estado);
    case 'notaItemDelete': return notaItemDelete_(body.id);
    case 'notaItemOrden':  return notaItemOrden_(body.ids);
    default: throw new Error('Acción no válida: ' + body.action);
  }
}

/** Número de fila (1-based) del registro con ese ID en la columna 1, o 0. */
function filaPorId_(sh, id) {
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === id) return i + 2;
  return 0;
}

function maxOrden_(sh, col) {
  if (sh.getLastRow() < 2) return 0;
  return sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0);
}

function notaSave_(n) {
  if (!n) throw new Error('Falta la nota.');
  return withLock_(() => {
    const sh = sheet_(SHEET_NOTAS), now = now_();
    const tipo = NOTA_TIPOS.indexOf(String(n.tipo || '').toLowerCase()) > -1 ? String(n.tipo).toLowerCase() : 'texto';
    const color = NOTA_COLORES.indexOf(String(n.color || '').toLowerCase()) > -1 ? String(n.color).toLowerCase() : 'default';
    let id = String(n.id || '').trim(), fila = id ? filaPorId_(sh, id) : 0;
    if (!id || !fila) {
      id = 'nota-' + Utilities.getUuid().slice(0, 8);
      sh.appendRow([id, String(n.titulo || ''), tipo, color, String(n.cuerpo || ''), maxOrden_(sh, 6) + 1, String(n.moverA || ''), '', now, now]);
    } else {
      const cur = sh.getRange(fila, 1, 1, NOTAS_HEADERS.length).getValues()[0];
      sh.getRange(fila, 1, 1, NOTAS_HEADERS.length).setValues([[id,
        n.titulo != null ? String(n.titulo) : cur[1], tipo, color,
        n.cuerpo != null ? String(n.cuerpo) : cur[4], cur[5],
        n.moverA != null ? String(n.moverA) : cur[6], cur[7], cur[8], now]]);
    }
    return getBoard_();
  });
}

function notaDelete_(id) {
  id = String(id || '').trim(); if (!id) throw new Error('Falta id de nota.');
  return withLock_(() => {
    const sh = sheet_(SHEET_NOTAS), fila = filaPorId_(sh, id);
    if (fila) sh.deleteRow(fila);
    const shI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTA_ITEMS);
    if (shI && shI.getLastRow() > 1) {
      const vals = shI.getRange(2, 1, shI.getLastRow() - 1, 2).getValues();
      for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][1]).trim() === id) shI.deleteRow(i + 2);
    }
    return getBoard_();
  });
}

function notaOrden_(ids) {
  if (!Array.isArray(ids)) throw new Error('Falta el orden.');
  return withLock_(() => {
    const sh = sheet_(SHEET_NOTAS);
    ids.forEach((id, i) => { const f = filaPorId_(sh, String(id).trim()); if (f) sh.getRange(f, 6).setValue(i + 1); });
    return getBoard_();
  });
}

function notaItemSave_(it) {
  if (!it || !String(it.nota || '').trim()) throw new Error('Falta la nota del ítem.');
  return withLock_(() => {
    const sh = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NOTA_ITEMS, NOTA_ITEMS_HEADERS), now = now_();
    const estado = NOTA_ITEM_ESTADOS.indexOf(String(it.estado || '').toLowerCase()) > -1 ? String(it.estado).toLowerCase() : 'pendiente';
    let id = String(it.id || '').trim(), fila = id ? filaPorId_(sh, id) : 0;
    if (!id || !fila) {
      id = 'it-' + Utilities.getUuid().slice(0, 8);
      sh.appendRow([id, String(it.nota).trim(), String(it.texto || ''), estado, String(it.notaItem || ''), maxOrden_(sh, 6) + 1, now, now]);
    } else {
      const cur = sh.getRange(fila, 1, 1, NOTA_ITEMS_HEADERS.length).getValues()[0];
      sh.getRange(fila, 1, 1, NOTA_ITEMS_HEADERS.length).setValues([[id, cur[1],
        it.texto != null ? String(it.texto) : cur[2], estado,
        it.notaItem != null ? String(it.notaItem) : cur[4], cur[5], cur[6], now]]);
    }
    return getBoard_();
  });
}

/** Cambia el estado de un ítem. Si pasa a 'hecho' y su nota tiene "Mover a", lo migra. */
function notaItemEstado_(id, estado) {
  id = String(id || '').trim(); if (!id) throw new Error('Falta id de ítem.');
  estado = NOTA_ITEM_ESTADOS.indexOf(String(estado || '').toLowerCase()) > -1 ? String(estado).toLowerCase() : 'pendiente';
  return withLock_(() => {
    const shI = sheet_(SHEET_NOTA_ITEMS), fila = filaPorId_(shI, id);
    if (!fila) return getBoard_();
    const notaId = String(shI.getRange(fila, 2).getValue()).trim();
    const shN = sheet_(SHEET_NOTAS), filaN = filaPorId_(shN, notaId);
    const moverA = filaN ? String(shN.getRange(filaN, 7).getValue() || '').trim() : '';
    if (estado === 'hecho' && moverA && filaPorId_(shN, moverA)) {
      shI.getRange(fila, 2).setValue(moverA);           // cambia de nota
      shI.getRange(fila, 4).setValue('hecho');
      shI.getRange(fila, 6).setValue(maxOrden_(shI, 6) + 1);
      shI.getRange(fila, 8).setValue(now_());
    } else {
      shI.getRange(fila, 4).setValue(estado);
      shI.getRange(fila, 8).setValue(now_());
    }
    return getBoard_();
  });
}

function notaItemDelete_(id) {
  id = String(id || '').trim(); if (!id) throw new Error('Falta id de ítem.');
  return withLock_(() => {
    const sh = sheet_(SHEET_NOTA_ITEMS), fila = filaPorId_(sh, id);
    if (fila) sh.deleteRow(fila);
    return getBoard_();
  });
}

function notaItemOrden_(ids) {
  if (!Array.isArray(ids)) throw new Error('Falta el orden.');
  return withLock_(() => {
    const sh = sheet_(SHEET_NOTA_ITEMS);
    ids.forEach((id, i) => { const f = filaPorId_(sh, String(id).trim()); if (f) sh.getRange(f, 6).setValue(i + 1); });
    return getBoard_();
  });
}

function toBool_(v) { const s = String(v).trim().toLowerCase(); return s === 'true' || s === '1' || s === 'sí' || s === 'si' || s === 'x' || s === 'verdadero'; }

/* ------------------------- SEMBRAR LIBROS (una vez) -------------------------
 * Crea dos notas: "✅ Libros leídos" y "📚 Por leer" (con los que estás leyendo
 * marcados en curso). Al terminar uno en "Por leer" y marcarlo leído, se pasa solo
 * a "Libros leídos". Revisa/ajusta los títulos después en la vista Notas.
 * Regla de carga: Keep tiene prioridad sobre el Word y los que ya están leídos se
 * quitan de "por leer" (p. ej. La sociedad del cansancio y La consagración de la
 * autenticidad, que estaban en ambas listas). */
function sembrarLibros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shN = ensureSheet_(ss, SHEET_NOTAS, NOTAS_HEADERS);
  const shI = ensureSheet_(ss, SHEET_NOTA_ITEMS, NOTA_ITEMS_HEADERS);

  // No duplicar si ya se sembró
  const yaN = shN.getLastRow() > 1 ? shN.getRange(2, 2, shN.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  if (yaN.some(t => /libros le[ií]dos|por leer/i.test(t))) {
    Logger.log('Parece que los libros ya se sembraron. No hago nada.');
    return;
  }

  const leidosId = 'nota-' + Utilities.getUuid().slice(0, 8);
  const porLeerId = 'nota-' + Utilities.getUuid().slice(0, 8);
  const now = now_();

  shN.appendRow([leidosId, '✅ Libros leídos', 'checklist', 'verde', '', 1, '', '', now, now]);
  // "Por leer" migra a "Leídos" al marcar un ítem como leído
  shN.appendRow([porLeerId, '📚 Por leer', 'checklist', 'azul', '', 2, leidosId, '', now, now]);

  const leidos = [
    'Y si quedamos como amigos', 'El diario de Ana Frank (2ª vez)', 'Compitiendo en la era de la IA',
    'Zonas erróneas', 'No soy un robot', 'Haz que te pasen cosas buenas', 'Encuentra tu persona vitamina',
    'Recupera tu mente, reconquista tu vida', 'Sapiens', 'Ikigai',
    'La sociedad del cansancio — Byung-Chul Han', 'Desde los zulos', 'IA, neurociencia y psicología',
    'IA: Secretos del cerebro'
  ];
  // {t: título, e: estado, n: nota del ítem}
  const porLeer = [
    { t: 'Las ciencias cognitivas: tendencias y perspectivas', e: 'en_curso' },
    { t: 'De cerebros, mentes y máquinas', e: 'en_curso' },
    { t: 'Interfaces humano-máquina (Secretos del cerebro)' },
    { t: 'La conciencia (Secretos del cerebro)' },
    { t: 'La emulación del cerebro (Secretos del cerebro)' },
    { t: 'Ciencia de datos — John D. Kelleher' },
    { t: 'Qué es la conciencia (Psicología y Neurociencia)' },
    { t: 'Homo Deus', n: 'me quedé al 20%' },
    { t: 'La vida secreta de la mente' },
    { t: 'Felicidad real en el trabajo' },
    { t: 'Maestría' },
    { t: 'La consagración de la autenticidad — Gilles Lipovetsky' },
    { t: 'Fisiología de la conducta' },
    { t: 'Matemáticas para Iñaki' },
    { t: '21 lecciones para el siglo XXI' },
    { t: 'Hábitos atómicos' },
    { t: 'Estoicismo' },
    { t: '¿Por qué a mí?' },
    { t: 'La sociedad del consumo — Baudrillard' },
    { t: 'Focus — Daniel Goleman' },
    { t: 'El yo saturado — Kenneth J. Gergen', n: 'buscar' },
    { t: 'Argumentando se entiende la gente' },
    { t: 'La inteligencia artificial (Los secretos del cerebro)' },
    { t: 'Tokio Blues' },
    { t: 'El conectoma (Los secretos del cerebro)' },
    { t: 'El viaje de los colibríes' },
    { t: 'Nexus' },
    { t: 'Confiar e inspirar — Stephen Covey', n: 'buscar' },
    { t: 'La estructura de las revoluciones científicas' },
    { t: 'Cómo educar a los niños en libertad — Colette Hovasse', n: 'buscar' },
    { t: 'Ser, saber, hacer — Mario Bunge' },
    { t: 'La ciencia, su método y su filosofía' },
    { t: 'Historia de las ideas científicas' },
    { t: 'Tecnología: un enfoque filosófico' },
    { t: 'Filosofía de la tecnología — Mario Bunge' }
  ];

  const filas = [];
  let orden = 1;
  leidos.forEach(t => filas.push(['it-' + Utilities.getUuid().slice(0, 8), leidosId, t, 'hecho', '', orden++, now, now]));
  orden = 1;
  porLeer.forEach(b => filas.push(['it-' + Utilities.getUuid().slice(0, 8), porLeerId, b.t, b.e || 'pendiente', b.n || '', orden++, now, now]));
  shI.getRange(shI.getLastRow() + 1, 1, filas.length, NOTA_ITEMS_HEADERS.length).setValues(filas);

  Logger.log('Sembrados ' + leidos.length + ' leídos y ' + porLeer.length + ' por leer (2 en curso).');
}
