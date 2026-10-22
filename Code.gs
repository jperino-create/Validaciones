/**
 * VALIDACIONES TT — Backend (Google Apps Script)
 * Vive dentro de la planilla "Plan de Mejora Continua".
 * Crea y mantiene las hojas VAL_* y guarda archivos/fotos en una carpeta de Drive.
 * "Hoja 1" (la planilla original) nunca se modifica.
 */

var APP_TITLE = 'Validaciones TT';
var FOLDER_NAME = 'Validaciones - Archivos';

var SCHEMA = {
  VAL_Proyectos: ['id', 'numero', 'nombre', 'fecha', 'solicita', 'impacto', 'areas', 'estado', 'tipo', 'prioridad',
    'piezas', 'supervisor', 'fechaObjetivo', 'cantidadAnual', 'tc', 'observacion', 'linkOriginal', 'creadoPor',
    'creado', 'actualizado'],
  VAL_Tareas: ['id', 'proyectoId', 'orden', 'item', 'actor', 'tarea', 'indicador', 'responsable', 'costoPlanta',
    'costoTercero', 'inicio', 'fin', 'fechaPiloto', 'revision', 'finPrueba', 'avance', 'estado', 'notas', 'actualizado'],
  VAL_Bitacora: ['id', 'proyectoId', 'tareaId', 'fecha', 'autor', 'tipo', 'texto'],
  VAL_Archivos: ['id', 'proyectoId', 'tareaId', 'nombre', 'mime', 'driveId', 'url', 'tamano', 'descripcion', 'autor', 'fecha'],
  VAL_Config: ['clave', 'valor']
};

var HEADERS_ES = {
  VAL_Proyectos: ['ID', 'N°', 'Necesidad / Proyecto', 'Fecha', 'Solicita', 'Impacto', 'Áreas', 'Estado', 'Tipo de mejora',
    'Prioridad', 'Piezas / códigos', 'Supervisor', 'Fecha objetivo', 'Cantidad anual', 'Tipo de cambio', 'Observación',
    'Planilla original', 'Creado por', 'Creado', 'Actualizado'],
  VAL_Tareas: ['ID', 'Proyecto ID', 'Orden', 'Ítem', 'Actor / Área', 'Tarea', 'Indicador de proceso', 'Responsable',
    'Costo planta (U$S)', 'Costo tercero (U$S)', 'Inicio', 'Fin', 'Fecha prueba piloto', 'Revisión', 'Fin de prueba',
    'Avance %', 'Estado', 'Notas', 'Actualizado'],
  VAL_Bitacora: ['ID', 'Proyecto ID', 'Tarea ID', 'Fecha', 'Autor', 'Tipo', 'Texto'],
  VAL_Archivos: ['ID', 'Proyecto ID', 'Tarea ID', 'Nombre', 'Tipo', 'Drive ID', 'URL', 'Tamaño (bytes)', 'Descripción', 'Autor', 'Fecha'],
  VAL_Config: ['Clave', 'Valor (JSON)']
};

var DATE_FIELDS = { fecha: 1, fechaObjetivo: 1, inicio: 1, fin: 1, fechaPiloto: 1, revision: 1, finPrueba: 1 };
var NUM_FIELDS = { numero: 1, orden: 1, costoPlanta: 1, costoTercero: 1, avance: 1, cantidadAnual: 1, tc: 1, tamano: 1 };

var DEFAULT_CONFIG = {
  estados: ['PROPUESTO', 'RECOLECCIÓN DE DATOS', 'EN ANÁLISIS', 'PRUEBA PILOTO', 'IMPLEMENTADO', 'NO IMPLEMENTADO', 'CANCELADO'],
  estadosTarea: ['Pendiente', 'En curso', 'Completada', 'Bloqueada'],
  areas: ['SUMINISTROS', 'PRODUCCIÓN', 'INGENIERÍA', 'CALIDAD', 'COMPRAS', 'LOGÍSTICA', 'POSVENTA', 'GERENCIA'],
  tipos: ['Rendimiento', 'Precio / Costo', 'Calidad', 'Disponibilidad', 'Seguridad', 'Otro'],
  prioridades: ['Alta', 'Media', 'Baja'],
  personas: ['MMORAD', 'SBELLO', 'JAUDISIO'],
  tc: 880,
  labelA: 'Costo planta',
  labelB: 'Costo tercero'
};

/* ------------------------------------------------------------------ */
/* Web app + menú                                                      */
/* ------------------------------------------------------------------ */

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---- API JSON para la versión publicada en GitHub Pages ---- */
function apiFns_() {
  return { getData: getData, saveConfig: saveConfig, saveProyecto: saveProyecto, deleteProyecto: deleteProyecto,
    saveTarea: saveTarea, saveTareasOrden: saveTareasOrden, deleteTarea: deleteTarea, addBitacora: addBitacora,
    deleteBitacora: deleteBitacora, uploadArchivo: uploadArchivo, updateArchivo: updateArchivo,
    deleteArchivo: deleteArchivo, getImagen: getImagen };
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (!req.key || req.key !== getApiKey_()) throw new Error('Clave de acceso incorrecta');
    var fns = apiFns_();
    if (!fns.hasOwnProperty(req.fn)) throw new Error('Función no permitida: ' + req.fn);
    out = { ok: true, result: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getApiKey_() {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty('API_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase(); props.setProperty('API_KEY', k); }
  return k;
}

/** Ejecutar desde el editor: muestra la clave de acceso de la app en el registro. */
function verClave() {
  var k = getApiKey_();
  Logger.log('CLAVE DE ACCESO: ' + k);
  return k;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Validaciones')
    .addItem('1) Crear / reparar hojas VAL_', 'setup')
    .addItem('2) Importar proyectos desde Hoja 1', 'importarDesdeHoja1')
    .addSeparator()
    .addItem('Abrir app', 'abrirApp')
    .addItem('Ver clave de acceso', 'mostrarClave')
    .addToUi();
}

function abrirApp() {
  var url = ScriptApp.getService().getUrl();
  var html = url
    ? '<p>Abrí la app acá:</p><p><a href="' + url + '" target="_blank">' + url + '</a></p>'
    : '<p>Todavía no está implementada como aplicación web.<br>En el editor de Apps Script: <b>Implementar → Nueva implementación → Aplicación web</b>.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(140), APP_TITLE);
}

function mostrarClave() {
  SpreadsheetApp.getUi().alert('Clave de acceso de Validaciones: ' + getApiKey_());
}

function setup() {
  ensureSheets_();
  Logger.log('Hojas VAL_ listas. CLAVE DE ACCESO: ' + getApiKey_());
  try { SpreadsheetApp.getActive().toast('Hojas VAL_ listas', APP_TITLE, 4); } catch (e) {}
}

/* ------------------------------------------------------------------ */
/* Hojas                                                               */
/* ------------------------------------------------------------------ */

function ss_() { return SpreadsheetApp.getActive(); }

function ensureSheets_() {
  var ss = ss_();
  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([HEADERS_ES[name]])
        .setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      if (name === 'VAL_Config') {
        var rows = Object.keys(DEFAULT_CONFIG).map(function (k) { return [k, JSON.stringify(DEFAULT_CONFIG[k])]; });
        sh.getRange(2, 1, rows.length, 2).setValues(rows);
      }
    }
  });
  ensureResumen_();
}

function ensureResumen_() {
  var ss = ss_();
  if (ss.getSheetByName('VAL_Resumen')) return;
  var sh = ss.insertSheet('VAL_Resumen');
  sh.getRange('A1').setValue('RESUMEN VALIDACIONES (se calcula solo)').setFontWeight('bold').setFontSize(14);
  sh.getRange('A3:B3').setValues([['Estado', 'Cantidad']]).setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
  var est = DEFAULT_CONFIG.estados;
  for (var i = 0; i < est.length; i++) {
    sh.getRange(4 + i, 1, 1, 2).setValues([[est[i], '=COUNTIF(VAL_Proyectos!H:H,A' + (4 + i) + ')']]);
  }
  var r = 5 + est.length;
  sh.getRange(r, 1, 4, 2).setValues([
    ['Total proyectos', '=COUNTA(VAL_Proyectos!A2:A)'],
    ['Tareas totales', '=COUNTA(VAL_Tareas!A2:A)'],
    ['Tareas completadas', '=COUNTIF(VAL_Tareas!Q:Q,"Completada")'],
    ['Archivos cargados', '=COUNTA(VAL_Archivos!A2:A)']
  ]);
  sh.setColumnWidth(1, 220);
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { ensureSheets_(); sh = ss_().getSheetByName(name); }
  return sh;
}

function tz_() { return ss_().getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires'; }

function toOut_(field, v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) {
    var hm = Utilities.formatDate(v, tz_(), 'HH:mm');
    return Utilities.formatDate(v, tz_(), hm !== '00:00' ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return v;
}

function readAll_(name) {
  var sh = sheet_(name), cols = SCHEMA[name];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === '' || vals[i][0] === null) continue;
    var o = {};
    for (var c = 0; c < cols.length; c++) o[cols[c]] = toOut_(cols[c], vals[i][c]);
    out.push(o);
  }
  return out;
}

function toRow_(name, obj) {
  return SCHEMA[name].map(function (f) {
    var v = obj[f];
    if (v === undefined || v === null) return '';
    if (NUM_FIELDS[f]) return v === '' ? '' : (isNaN(Number(v)) ? '' : Number(v));
    if (DATE_FIELDS[f] && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
      var p = v.split('-');
      return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    }
    if (typeof v === 'string' && /^[=+\-@]/.test(v) && !NUM_FIELDS[f]) return "'" + v; // evita fórmulas
    return v;
  });
}

function findRow_(name, id) {
  var sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function upsert_(name, obj) {
  var sh = sheet_(name);
  var row = toRow_(name, obj);
  var r = findRow_(name, obj.id);
  if (r < 0) sh.appendRow(row);
  else sh.getRange(r, 1, 1, row.length).setValues([row]);
}

function deleteWhere_(name, pred) {
  var sh = sheet_(name), cols = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues(), removed = [];
  for (var i = vals.length - 1; i >= 0; i--) {
    var o = {};
    for (var c = 0; c < cols.length; c++) o[cols[c]] = vals[i][c];
    if (pred(o)) { sh.deleteRow(i + 2); removed.push(o); }
  }
  return removed;
}

function uid_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function currentUser_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

/* ------------------------------------------------------------------ */
/* API pública (google.script.run)                                     */
/* ------------------------------------------------------------------ */

function getData() {
  ensureSheets_();
  var cfg = {};
  readAll_('VAL_Config').forEach(function (r) {
    try { cfg[r.clave] = JSON.parse(r.valor); } catch (e) { cfg[r.clave] = r.valor; }
  });
  Object.keys(DEFAULT_CONFIG).forEach(function (k) { if (cfg[k] === undefined) cfg[k] = DEFAULT_CONFIG[k]; });
  return JSON.stringify({
    proyectos: readAll_('VAL_Proyectos'),
    tareas: readAll_('VAL_Tareas'),
    bitacora: readAll_('VAL_Bitacora'),
    archivos: readAll_('VAL_Archivos'),
    config: cfg,
    user: currentUser_(),
    sheetUrl: ss_().getUrl()
  });
}

function saveConfig(cfgJson) {
  var cfg = JSON.parse(cfgJson);
  return withLock_(function () {
    Object.keys(cfg).forEach(function (k) { upsertConfig_(k, cfg[k]); });
    return true;
  });
}

function upsertConfig_(k, v) {
  var sh = sheet_('VAL_Config'), r = findRow_('VAL_Config', k);
  var row = [k, JSON.stringify(v)];
  if (r < 0) sh.appendRow(row); else sh.getRange(r, 1, 1, 2).setValues([row]);
}

function saveProyecto(pJson, autor) {
  var p = JSON.parse(pJson);
  return withLock_(function () {
    var list = readAll_('VAL_Proyectos');
    var prev = null;
    for (var i = 0; i < list.length; i++) if (list[i].id === p.id) prev = list[i];
    if (!p.id) p.id = uid_('P');
    if (!p.numero) {
      var max = 0;
      list.forEach(function (x) { if (Number(x.numero) > max) max = Number(x.numero); });
      p.numero = max + 1;
    }
    if (!prev) { p.creado = now_(); p.creadoPor = autor || currentUser_(); }
    else { p.creado = prev.creado; p.creadoPor = prev.creadoPor; }
    p.actualizado = now_();
    upsert_('VAL_Proyectos', p);
    if (!prev) addLog_(p.id, '', autor, 'Alta', 'Proyecto creado: ' + (p.nombre || ''));
    else if (prev.estado !== p.estado) addLog_(p.id, '', autor, 'Cambio de estado', (prev.estado || '—') + ' → ' + (p.estado || '—'));
    return JSON.stringify(p);
  });
}

function deleteProyecto(id) {
  return withLock_(function () {
    deleteWhere_('VAL_Proyectos', function (o) { return o.id === id; });
    deleteWhere_('VAL_Tareas', function (o) { return o.proyectoId === id; });
    deleteWhere_('VAL_Bitacora', function (o) { return o.proyectoId === id; });
    var arch = deleteWhere_('VAL_Archivos', function (o) { return o.proyectoId === id; });
    arch.forEach(function (a) { try { DriveApp.getFileById(a.driveId).setTrashed(true); } catch (e) {} });
    return true;
  });
}

function saveTarea(tJson, autor) {
  var t = JSON.parse(tJson);
  return withLock_(function () {
    var isNew = !t.id;
    var prev = null;
    if (!isNew) readAll_('VAL_Tareas').forEach(function (x) { if (x.id === t.id) prev = x; });
    if (isNew) t.id = uid_('T');
    t.actualizado = now_();
    upsert_('VAL_Tareas', t);
    if (prev && prev.estado !== t.estado && t.estado) {
      addLog_(t.proyectoId, t.id, autor, 'Tarea', '"' + (t.tarea || t.indicador || 'Tarea') + '": ' + (prev.estado || '—') + ' → ' + t.estado);
    }
    touchProyecto_(t.proyectoId);
    return JSON.stringify(t);
  });
}

function saveTareasOrden(listJson) {
  var list = JSON.parse(listJson); // [{id, orden}]
  return withLock_(function () {
    var all = readAll_('VAL_Tareas'), map = {};
    list.forEach(function (x) { map[x.id] = x.orden; });
    all.forEach(function (t) { if (map[t.id] !== undefined && map[t.id] !== t.orden) { t.orden = map[t.id]; upsert_('VAL_Tareas', t); } });
    return true;
  });
}

function deleteTarea(id) {
  return withLock_(function () {
    deleteWhere_('VAL_Tareas', function (o) { return o.id === id; });
    return true;
  });
}

function touchProyecto_(pid) {
  var r = findRow_('VAL_Proyectos', pid);
  if (r > 0) sheet_('VAL_Proyectos').getRange(r, SCHEMA.VAL_Proyectos.indexOf('actualizado') + 1).setValue(now_());
}

function addLog_(pid, tid, autor, tipo, texto) {
  var e = { id: uid_('B'), proyectoId: pid, tareaId: tid || '', fecha: now_(), autor: autor || currentUser_(), tipo: tipo, texto: texto };
  sheet_('VAL_Bitacora').appendRow(toRow_('VAL_Bitacora', e));
  return e;
}

function addBitacora(eJson) {
  var e = JSON.parse(eJson);
  return withLock_(function () {
    return JSON.stringify(addLog_(e.proyectoId, e.tareaId, e.autor, e.tipo || 'Nota', e.texto));
  });
}

function deleteBitacora(id) {
  return withLock_(function () { deleteWhere_('VAL_Bitacora', function (o) { return o.id === id; }); return true; });
}

/* ------------------------------------------------------------------ */
/* Archivos / fotos en Drive                                           */
/* ------------------------------------------------------------------ */

function rootFolder_() {
  var parent;
  try {
    var it = DriveApp.getFileById(ss_().getId()).getParents();
    parent = it.hasNext() ? it.next() : DriveApp.getRootFolder();
  } catch (e) { parent = DriveApp.getRootFolder(); }
  var f = parent.getFoldersByName(FOLDER_NAME);
  if (f.hasNext()) return f.next();
  var folder = parent.createFolder(FOLDER_NAME);
  try {
    var eds = ss_().getEditors().map(function (u) { return u.getEmail(); }).filter(String);
    if (eds.length) folder.addEditors(eds);
  } catch (e) {}
  return folder;
}

function projectFolder_(pid) {
  var root = rootFolder_();
  var p = null;
  readAll_('VAL_Proyectos').forEach(function (x) { if (x.id === pid) p = x; });
  var name = p ? ('P' + p.numero + ' - ' + String(p.nombre || '').slice(0, 60)) : pid;
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function uploadArchivo(metaJson, base64) {
  var m = JSON.parse(metaJson);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), m.mime || 'application/octet-stream', m.nombre || 'archivo');
  var file = projectFolder_(m.proyectoId).createFile(blob);
  if (m.descripcion) file.setDescription(m.descripcion);
  var a = {
    id: uid_('A'), proyectoId: m.proyectoId, tareaId: m.tareaId || '', nombre: m.nombre, mime: m.mime,
    driveId: file.getId(), url: file.getUrl(), tamano: file.getSize(), descripcion: m.descripcion || '',
    autor: m.autor || currentUser_(), fecha: now_()
  };
  withLock_(function () {
    sheet_('VAL_Archivos').appendRow(toRow_('VAL_Archivos', a));
    addLog_(m.proyectoId, m.tareaId, m.autor, 'Archivo', 'Subió ' + a.nombre);
  });
  return JSON.stringify(a);
}

function updateArchivo(aJson) {
  var a = JSON.parse(aJson);
  return withLock_(function () {
    var cur = null;
    readAll_('VAL_Archivos').forEach(function (x) { if (x.id === a.id) cur = x; });
    if (!cur) return false;
    cur.descripcion = a.descripcion || '';
    cur.tareaId = a.tareaId || '';
    upsert_('VAL_Archivos', cur);
    return true;
  });
}

function deleteArchivo(id) {
  return withLock_(function () {
    var rem = deleteWhere_('VAL_Archivos', function (o) { return o.id === id; });
    rem.forEach(function (a) { try { DriveApp.getFileById(a.driveId).setTrashed(true); } catch (e) {} }); // va a la papelera (recuperable)
    return true;
  });
}

/** Miniatura (o imagen completa si full=true y pesa < 4 MB) como data URL. */
function getImagen(driveId, full) {
  try {
    var f = DriveApp.getFileById(driveId);
    var b = (full && f.getSize() < 4 * 1024 * 1024) ? f.getBlob() : f.getThumbnail();
    if (!b) return '';
    return 'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

/* ------------------------------------------------------------------ */
/* Importación desde Hoja 1 + planillas de cada proyecto               */
/* ------------------------------------------------------------------ */

function importarDesdeHoja1() {
  ensureSheets_();
  var ss = ss_();
  var src = ss.getSheetByName('Hoja 1') || ss.getSheets()[0];
  var last = src.getLastRow();
  if (last < 3) return 'Hoja 1 vacía';
  var rng = src.getRange(3, 1, last - 2, 8);
  var vals = rng.getDisplayValues();
  var rich = [];
  try { rich = src.getRange(3, 1, last - 2, 1).getRichTextValues(); } catch (e) {}
  var forms = [];
  try { forms = src.getRange(3, 1, last - 2, 1).getFormulas(); } catch (e) {}

  var existentes = {};
  readAll_('VAL_Proyectos').forEach(function (p) { existentes[String(p.numero)] = p; });

  var creados = 0, tareas = 0, errores = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v[3] && !v[1]) continue; // sin necesidad ni fecha → fila vacía
    var num = Number(String(v[0]).replace(/\D/g, '')) || (i + 1);
    if (existentes[String(num)]) continue; // ya importado
    var link = '';
    try { link = rich[i] && rich[i][0] && rich[i][0].getLinkUrl() || ''; } catch (e) {}
    if (!link && forms[i] && /HYPERLINK\("([^"]+)"/i.test(forms[i][0])) link = RegExp.$1;

    var p = {
      id: uid_('P'), numero: num, nombre: v[3], fecha: parseFechaAR_(v[1]), solicita: v[2], impacto: v[4],
      areas: v[5], estado: String(v[6] || 'PROPUESTO').toUpperCase(), observacion: v[7], linkOriginal: link,
      tipo: '', prioridad: 'Media', tc: DEFAULT_CONFIG.tc, creado: now_(), creadoPor: 'Importación', actualizado: now_()
    };
    if (link) {
      try {
        var det = leerPlantilla_(link);
        if (det.supervisor) p.supervisor = det.supervisor;
        if (det.tc) p.tc = det.tc;
        det.tareas.forEach(function (t, k) {
          if (p.estado === 'IMPLEMENTADO' && (!t.estado || t.estado === 'Pendiente')) { t.estado = 'Completada'; t.avance = 100; }
          t.id = uid_('T'); t.proyectoId = p.id; t.orden = k + 1; t.actualizado = now_();
          upsert_('VAL_Tareas', t); tareas++;
        });
      } catch (e) { errores.push('P' + num + ': ' + e.message); }
    }
    upsert_('VAL_Proyectos', p);
    addLog_(p.id, '', 'Importación', 'Alta', 'Importado desde Hoja 1');
    creados++;
  }
  var msg = 'Importados ' + creados + ' proyectos y ' + tareas + ' filas de plan de acción.' + (errores.length ? ' Avisos: ' + errores.join(' | ') : '');
  try { ss.toast(msg, APP_TITLE, 8); } catch (e) {}
  return msg;
}

function parseFechaAR_(s) {
  var m = String(s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return '';
  return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

function parseMonto_(s) {
  s = String(s || '').replace(/U\$S|\$|\s/g, '');
  if (!s || s === '-') return '';
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); // 1.215,44
  else s = s.replace(/,/g, '');                                            // 1,215.44
  var n = Number(s);
  return isNaN(n) ? '' : n;
}

/** Lee una planilla con el formato "Plantilla Proyecto" (Plan de Acción). */
function leerPlantilla_(url) {
  var sh = SpreadsheetApp.openByUrl(url).getSheets()[0];
  var v = sh.getDataRange().getDisplayValues();
  var out = { tareas: [], supervisor: '', tc: '' };
  var head = -1;
  for (var r = 0; r < v.length; r++) {
    var row = v[r];
    var iSup = row.indexOf('Supervisor');
    if (iSup >= 0 && row[iSup + 1]) out.supervisor = row[iSup + 1];
    if (String(row[0]).trim().toLowerCase().indexOf('necesidad') === 0) {
      for (var c = row.length - 1; c > 0; c--) { var n = parseMonto_(row[c]); if (n !== '' && n > 1) { out.tc = n; break; } }
    }
    if (String(row[0]).trim().toLowerCase() === 'ítem' || String(row[0]).trim().toLowerCase() === 'item') { head = r; break; }
  }
  if (head < 0) return out;
  var item = '', actor = '', tarea = '';
  for (var r2 = head + 1; r2 < v.length; r2++) {
    var x = v[r2];
    var a0 = String(x[0]).trim();
    if (/^total/i.test(a0) || /^total/i.test(String(x[1]).trim())) break;
    if (!x.slice(0, 10).join('').trim()) continue;
    if (a0) item = a0;
    if (String(x[1]).trim()) actor = String(x[1]).trim();
    if (String(x[2]).trim()) tarea = String(x[2]).trim();
    out.tareas.push({
      item: item, actor: actor, tarea: tarea, indicador: String(x[3]).trim(),
      costoPlanta: parseMonto_(x[4]), costoTercero: parseMonto_(x[5]),
      fechaPiloto: parseFechaAR_(x[6]), revision: parseFechaAR_(x[7]), finPrueba: parseFechaAR_(x[8]),
      estado: x[9] ? String(x[9]).trim() : 'Pendiente', avance: 0, responsable: '', inicio: '', fin: '', notas: ''
    });
  }
  return out;
}
