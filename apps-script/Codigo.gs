// Conecta la página de Tesorería (GitHub Pages) con la planilla de Google.
// Se publica como app web: "Ejecutar como: yo" y "Quién tiene acceso: cualquier persona".
// La clave de acceso está en la pestaña Config, celda B2.

const PLANILLA_ID = '17t29hTJ08SV7CQw18ZVEpYAD2vyREloek-ZEdGNWwyE';
const HOJA = 'Jugadores';
const HOJA_CONFIG = 'Config';
const COLUMNAS = [
  ['id', 'ID'], ['actividad', 'Actividad'], ['nombre', 'Nombre y apellido'], ['grupo', 'Plantel / categoría'],
  ['dni', 'DNI'], ['nacimiento', 'Nacimiento'], ['activo', 'Activo'], ['reparto', 'Qué paga'],
  ['cuotaEspecial', 'Cuota especial'], ['telefono', 'Teléfono'], ['contacto', 'Contacto'], ['notas', 'Notas'],
  ['alta', 'Alta'], ['creado', 'Creado'], ['pagos', 'Pagos (no editar)'],
];
const ESTADOS = { pago: 'Pagó', debe: 'Debe', novino: 'No vino', bonif: 'Bonificado' };
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

function doGet() {
  return responder(() => ({ estado: 'funcionando' }));
}

function doPost(e) {
  return responder(() => atender(JSON.parse(e.postData.contents)));
}

function responder(fn) {
  let salida;
  try {
    salida = Object.assign({ ok: true }, fn());
  } catch (err) {
    salida = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

function atender(p) {
  const ss = SpreadsheetApp.openById(PLANILLA_ID);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cfg = hojaConfig(ss);
    const clave = String(cfg.getRange('B2').getValue() || '');
    if (!clave && p.accion === 'importar' && p.clave) cfg.getRange('B2').setValue(String(p.clave));
    else if (!clave || String(p.clave || '') !== clave) throw new Error('clave');

    switch (p.accion) {
      case 'version':
        return { version: version() };
      case 'cargar':
        return { config: leerConfig(cfg), socios: leerSocios(ss), version: version() };
      case 'guardarSocio':
        guardarSocio(ss, String(p.id), p.datos || {});
        return { version: subirVersion() };
      case 'borrarSocio':
        borrarSocio(ss, String(p.id));
        return { version: subirVersion() };
      case 'guardarConfig':
        cfg.getRange('B1').setValue(JSON.stringify(p.config));
        return { version: subirVersion() };
      case 'importar':
        importar(ss, cfg, p.config, p.socios || []);
        return { version: subirVersion() };
      default:
        throw new Error('accion');
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Versión: avisa a las páginas abiertas que hubo cambios ---------- */
function version() {
  return Number(PropertiesService.getScriptProperties().getProperty('version') || 0);
}
function subirVersion() {
  const v = version() + 1;
  PropertiesService.getScriptProperties().setProperty('version', String(v));
  return v;
}

/* ---------- Pestañas ---------- */
function hojaConfig(ss) {
  let h = ss.getSheetByName(HOJA_CONFIG);
  if (!h) {
    h = ss.insertSheet(HOJA_CONFIG);
    h.getRange('A1:B3').setNumberFormat('@');
    h.getRange('A1:A3').setValues([['Datos del sistema (no editar)'], ['Clave de acceso'], ['Para cambiar la clave, escribí la nueva en B2.']]);
    h.getRange('A1:A2').setFontWeight('bold');
    h.setColumnWidth(1, 260);
  }
  return h;
}

function hojaJugadores(ss) {
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    const primera = ss.getSheets()[0];
    if (primera && primera.getName() !== HOJA_CONFIG && primera.getLastRow() === 0) {
      h = primera.setName(HOJA);
    } else {
      h = ss.insertSheet(HOJA, 0);
    }
  }
  if (h.getLastRow() === 0) escribirEncabezado(h, []);
  return h;
}

function escribirEncabezado(h, meses) {
  const enc = COLUMNAS.map(c => c[1]).concat(meses.map(nombreMes));
  const r = h.getRange(1, 1, 1, enc.length);
  r.setNumberFormat('@').setValues([enc]).setFontWeight('bold').setBackground('#0E1116').setFontColor('#FFFFFF');
  h.setFrozenRows(1);
  h.setFrozenColumns(3);
}

/* ---------- Config ---------- */
function leerConfig(cfg) {
  const txt = String(cfg.getRange('B1').getValue() || '');
  return txt ? JSON.parse(txt) : null;
}

/* ---------- Jugadores ---------- */
function nombreMes(m) {
  return MESES[Number(m.slice(5)) - 1] + ' ' + m.slice(0, 4);
}
function claveMes(texto) {
  const partes = String(texto).trim().split(/\s+/);
  const i = MESES.findIndex(x => x.toLowerCase() === String(partes[0]).toLowerCase());
  return i >= 0 && /^\d{4}$/.test(partes[1] || '') ? partes[1] + '-' + String(i + 1).padStart(2, '0') : null;
}
function texto(v, formato) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), formato);
  return v === null || v === undefined ? '' : String(v);
}

function filaASocio(f) {
  const s = {};
  COLUMNAS.forEach(([k], i) => {
    const v = f[i];
    if (k === 'activo') s.activo = !(v === false || String(v).toLowerCase() === 'no');
    else if (k === 'cuotaEspecial') s.cuotaEspecial = v === '' || v === null ? null : Number(v);
    else if (k === 'pagos') s.pagos = v ? JSON.parse(v) : {};
    else if (k === 'nacimiento') s.nacimiento = texto(v, 'yyyy-MM-dd');
    else if (k === 'alta') s.alta = texto(v, 'yyyy-MM');
    else s[k] = texto(v, "yyyy-MM-dd'T'HH:mm:ss");
  });
  return s;
}

function celdaMes(p) {
  if (!p || !ESTADOS[p.e]) return '';
  const plata = n => '$' + Math.round(n || 0).toLocaleString('es-AR');
  if (p.e === 'pago') return 'Pagó ' + plata((p.a || 0) + (p.s || 0));
  if (p.e === 'debe') return 'Debe ' + plata(p.v);
  return ESTADOS[p.e];
}

function socioAFila(s, meses) {
  const base = COLUMNAS.map(([k]) => {
    if (k === 'activo') return s.activo === false ? 'No' : 'Sí';
    if (k === 'cuotaEspecial') return s.cuotaEspecial === null || s.cuotaEspecial === undefined || s.cuotaEspecial === '' ? '' : String(s.cuotaEspecial);
    if (k === 'pagos') return JSON.stringify(s.pagos || {});
    return s[k] === null || s[k] === undefined ? '' : String(s[k]);
  });
  return base.concat(meses.map(m => celdaMes((s.pagos || {})[m])));
}

function mesesDeEncabezado(h) {
  const n = h.getLastColumn();
  if (n <= COLUMNAS.length) return [];
  return h.getRange(1, COLUMNAS.length + 1, 1, n - COLUMNAS.length).getValues()[0].map(claveMes);
}

function leerSocios(ss) {
  const h = hojaJugadores(ss);
  if (h.getLastRow() < 2) return [];
  const datos = h.getRange(2, 1, h.getLastRow() - 1, COLUMNAS.length).getValues();
  return datos.filter(f => f[0] !== '').map(filaASocio);
}

function filaDe(h, id) {
  if (h.getLastRow() < 2) return -1;
  const ids = h.getRange(2, 1, h.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function guardarSocio(ss, id, datos) {
  const h = hojaJugadores(ss);
  const fila = filaDe(h, id);
  const actual = fila > 0 ? filaASocio(h.getRange(fila, 1, 1, COLUMNAS.length).getValues()[0]) : { id: id, pagos: {} };
  Object.keys(datos).forEach(k => {
    if (k === 'pagos') {
      Object.keys(datos.pagos || {}).forEach(m => {
        if (datos.pagos[m] === null) delete actual.pagos[m];
        else actual.pagos[m] = datos.pagos[m];
      });
    } else if (k !== 'id') {
      actual[k] = datos[k];
    }
  });
  actual.id = id;

  let meses = mesesDeEncabezado(h);
  const faltan = Object.keys(actual.pagos).filter(m => meses.indexOf(m) < 0).sort();
  if (faltan.length) {
    meses = meses.concat(faltan);
    const col = COLUMNAS.length + meses.length - faltan.length + 1;
    h.getRange(1, col, 1, faltan.length).setNumberFormat('@').setValues([faltan.map(nombreMes)])
      .setFontWeight('bold').setBackground('#0E1116').setFontColor('#FFFFFF');
  }
  const valores = socioAFila(actual, meses);
  const destino = fila > 0 ? fila : h.getLastRow() + 1;
  h.getRange(destino, 1, 1, valores.length).setNumberFormat('@').setValues([valores]);
}

function borrarSocio(ss, id) {
  const h = hojaJugadores(ss);
  const fila = filaDe(h, id);
  if (fila > 0) h.deleteRow(fila);
}

function importar(ss, cfg, config, socios) {
  cfg.getRange('B1').setValue(JSON.stringify(config));
  const h = hojaJugadores(ss);
  h.clear();
  const meses = Array.from(new Set(socios.reduce((t, s) => t.concat(Object.keys(s.pagos || {})), []))).sort();
  escribirEncabezado(h, meses);
  if (!socios.length) return;
  const orden = socios.slice().sort((a, b) => (a.actividad + a.grupo + a.nombre).localeCompare(b.actividad + b.grupo + b.nombre));
  const filas = orden.map(s => socioAFila(s, meses));
  h.getRange(2, 1, filas.length, filas[0].length).setNumberFormat('@').setValues(filas);
  h.setColumnWidth(1, 60);
  h.setColumnWidth(3, 200);
  h.setColumnWidth(COLUMNAS.length, 120);
}
