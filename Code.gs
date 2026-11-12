/**
 * WASTE CONSOLIDADO v2 - Historias y un Café
 *
 * Proyecto de Apps Script NUEVO (no es el de Kitchen Ops).
 * Lee el Waste de las dos apps directamente de sus Sheets y lo entrega
 * a la página de GitHub Pages. Solo lectura: no escribe nada en los Sheets.
 *
 *  - Cocina:     Sheet Kitchen Ops, pestañas "Cierres" y "Archivo"
 *                (filas con Sección = Waste e ID que empieza con W)
 *  - Coffee bar: Sheet del Coffee Bar, pestaña "WASTE HISTORIAL"
 *
 * Carga por fases (igual que el Dashboard de Reviews):
 *  - getRecentData(force): mes actual y mes anterior. Es lo que se ve al abrir.
 *  - getOlderData():       meses anteriores, en segundo plano.
 *
 * Sin PIN: cualquiera que tenga el enlace puede ver el Waste.
 */

const KITCHEN_ID = '1BfgJnnmD6RoqrP2EEkK8qU0zfluW3K9Xr8QrkkhMiB8';
const COFFEE_ID  = '1tsuZNjLBo5KNCBkiWC-L_dQyDi4hobmwMyYpjtoUArw';
const TZ = 'America/Puerto_Rico';

const DIAS_ATRAS = 400;      // cuánto historial se envía en total
const CACHE_RECIENTE = 60;   // segundos
const CACHE_ANTERIOR = 1800; // 30 minutos

// ── ENTRADA WEB ───────────────────────────────────────────────────────────────
function doGet() {
  return ContentService
    .createTextOutput('Waste consolidado: API activa')
    .setMimeType(ContentService.MimeType.TEXT);
}

// La página manda {fn, args}. Solo se permiten las funciones de esta lista.
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const permitidas = { getRecentData: getRecentData, getOlderData: getOlderData };
    if (!permitidas[req.fn]) throw new Error('Función no permitida');
    out = { ok: true, data: permitidas[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── FASES ─────────────────────────────────────────────────────────────────────
// Contrato: { scope: 'recent'|'older', hoy: 'yyyy-MM-dd', from: 'yyyy-MM',
//             rows: [[fecha, hora, 'C'|'B', producto, cantidad, unidad,
//                     razon, responsable, registro, nota], ...] }
// 'C' = Cocina, 'B' = Coffee bar. 'from' es el primer mes que cubre "recent".
function getRecentData(force) {
  const hoy = hoyClave_();
  const inicio = inicioMesAnterior_(hoy);
  const cache = CacheService.getScriptCache();
  const key = 'wc_recent_' + hoy;
  if (!force) {
    const c = cache.get(key);
    if (c) return JSON.parse(c);
  }
  const rows = leerTodo_(inicio, null);
  const out = { scope: 'recent', hoy: hoy, from: inicio.substring(0, 7), rows: rows };
  guardarCache_(cache, key, out, CACHE_RECIENTE);
  return out;
}

function getOlderData() {
  const hoy = hoyClave_();
  const inicio = inicioMesAnterior_(hoy);
  const cache = CacheService.getScriptCache();
  const key = 'wc_older_' + inicio;
  const c = cache.get(key);
  if (c) return JSON.parse(c);
  const rows = leerTodo_(restarDias_(hoy, DIAS_ATRAS), inicio);   // fechas anteriores a "inicio"
  const out = { scope: 'older', hoy: hoy, from: inicio.substring(0, 7), rows: rows };
  guardarCache_(cache, key, out, CACHE_ANTERIOR);
  return out;
}

function guardarCache_(cache, key, obj, seg) {
  try { cache.put(key, JSON.stringify(obj), seg); } catch (e) { /* si no cabe, se omite */ }
}

function leerTodo_(desde, antesDe) {
  return leerCocina_(desde, antesDe).concat(leerCoffee_(desde, antesDe));
}

// ── LECTURA: COCINA ───────────────────────────────────────────────────────────
// Columnas de Cierres/Archivo: A Fecha, B Timestamp, C Tarea ID, D Tarea,
// E Empleado (quien registró), F Hora, G Sección, H Razón, I Responsable,
// J Cantidad, K Unidad
function leerCocina_(desde, antesDe) {
  const ss = SpreadsheetApp.openById(KITCHEN_ID);
  const tz = ss.getSpreadsheetTimeZone() || TZ;
  const out = [];
  ['Cierres', 'Archivo'].forEach(function (nombre) {
    const sh = ss.getSheetByName(nombre);
    if (!sh) return;
    const n = sh.getLastRow() - 1;
    if (n < 1) return;
    const vals = sh.getRange(2, 1, n, 11).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[6]) !== 'Waste') continue;
      if (!/^W\d+/.test(String(r[2]))) continue;   // descarta SUPERVISION y marcas del sistema
      const f = fechaTexto_(r[0], tz);
      if (!f || f < desde || (antesDe && f >= antesDe)) continue;
      out.push([
        f,
        horaTexto_(r[5], tz),
        'C',
        String(r[3] || '').trim(),
        numero_(r[9]),
        String(r[10] || '').trim(),
        String(r[7] || '').trim(),
        String(r[8] || '').trim(),
        String(r[4] || '').trim(),
        ''
      ]);
    }
  });
  return out;
}

// ── LECTURA: COFFEE BAR ───────────────────────────────────────────────────────
// Columnas de WASTE HISTORIAL: A Fecha, B Hora, C Producto, D Cantidad,
// E Unidad, F Razón, G Empleado, H Notas
function leerCoffee_(desde, antesDe) {
  const ss = SpreadsheetApp.openById(COFFEE_ID);
  const tz = ss.getSpreadsheetTimeZone() || TZ;
  const sh = ss.getSheetByName('WASTE HISTORIAL');
  if (!sh) return [];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, 8).getValues();
  const horas = sh.getRange(2, 2, n, 1).getDisplayValues();   // lo que se ve en pantalla
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const f = fechaTexto_(r[0], tz);
    const p = String(r[2] || '').trim();
    if (!f || f < desde || (antesDe && f >= antesDe) || !p) continue;
    let h = String(horas[i][0] || '').trim();
    if (/\d{1,2}\/\d{1,2}\/\d{2,4}/.test(h)) h = '';   // salió como fecha 12/30/1899: no hay hora real
    out.push([
      f,
      h,
      'B',
      p,
      numero_(r[3]),
      String(r[4] || '').trim(),
      String(r[5] || '').trim(),
      String(r[6] || '').trim(),
      '',
      String(r[7] || '').trim()
    ]);
  }
  return out;
}

// ── UTILIDADES ────────────────────────────────────────────────────────────────
function hoyClave_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function inicioMesAnterior_(ymd) {
  const p = ymd.split('-');
  const d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 2, 1));
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-01';
}

function restarDias_(ymd, dias) {
  const p = ymd.split('-');
  const d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
  d.setUTCDate(d.getUTCDate() - dias);
  return d.getUTCFullYear() + '-' +
    String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(d.getUTCDate()).padStart(2, '0');
}

// Acepta fecha real, texto yyyy-MM-dd o texto m/d/yyyy. Devuelve yyyy-MM-dd o ''.
function fechaTexto_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
  return '';
}

function horaTexto_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'h:mm a');
  return String(v || '').trim();
}

function numero_(v) {
  if (typeof v === 'number') return v;
  const s = String(v === null || v === undefined ? '' : v).trim().replace(',', '.');
  if (s === '') return '';
  const n = parseFloat(s);
  return isNaN(n) ? '' : n;
}

// ── PRUEBA DESDE EL EDITOR ────────────────────────────────────────────────────
// Córrela una vez para autorizar los permisos y ver cuántos registros lee.
function probarLectura() {
  const r = getRecentData(true);
  const o = getOlderData();
  const cuenta = function (rows, t) { return rows.filter(function (x) { return x[2] === t; }).length; };
  Logger.log('Mes actual y anterior -> Cocina: ' + cuenta(r.rows, 'C') + ' | Coffee bar: ' + cuenta(r.rows, 'B'));
  Logger.log('Meses anteriores     -> Cocina: ' + cuenta(o.rows, 'C') + ' | Coffee bar: ' + cuenta(o.rows, 'B'));
  if (r.rows.length) Logger.log('Ejemplo: ' + JSON.stringify(r.rows[r.rows.length - 1]));
}
