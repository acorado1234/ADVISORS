/**
 * Status de campañas · Advisors — servidor (Google Apps Script)
 * La hoja de cálculo de Andrea es la base de datos. Este script recibe los
 * cambios del tablero, revisa el link personal, escribe en la hoja y guarda el historial.
 *
 * Las advisors NO están fijas en el código: salen de la pestaña _Usuarios
 * (filas con Rol = advisor). Andrea crea perfiles nuevos desde el tablero (Equipo).
 */
const TABLERO_URL = 'https://acorado1234.github.io/ADVISORS/';
const ADVISORS_INICIALES = ['Mafer', 'Adriana', 'Cintia'];
const USUARIOS = '_Usuarios', HISTORIAL = '_Historial', IDC = '_id';
const CLAVES = ['Status', 'Agencia', 'Marca', 'Campaña'];
const LISTAS_STD = {
  'Puntaje': ['A | 1', 'B | 2', 'C | 3'], 'Errores operativos': ['1', '2', '3'],
  'Status': ['Propuesta','En revisión','Cambios','Activación','Contratos','Calendario','Producción','Reporte','Finanzas','NPS','Cerrada','Pérdida','Pausada'],
  'Propuesta': ['Enviada','Aprobada','En cambios','Pérdida','POR TALENTO'], 'Activación': ['Confirmada','Pérdida','En pausa'],
  'Status calendario': ['Enviado','Pendiente','En cambios','Aprobado','En espera de fechas'],
  'Status de producción': ['En producción','Pendiente','Enviado','En cambios','Aprobado','Publicado','ATRASADO','En pausa']
};
const ZONA = () => Session.getScriptTimeZone();

function doGet(e) { return salida(e.parameter || {}); }
function doPost(e) { let p = {}; try { p = JSON.parse(e.postData.contents || '{}'); } catch (x) {} return salida(p); }
function salida(p) {
  let out; try { out = Object.assign({ ok: true }, rutear(p)); } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function rutear(p) {
  if (p.action === 'entrar') return entrar(p.llave);
  const u = verificar(p.perfil, p.token);
  switch (p.action) {
    case 'equipo': return equipo(u);
    case 'crearPerfil': return conLock(() => crearPerfil(u, p));
    case 'nuevoLink': return conLock(() => nuevoLink(u, p.objetivo));
    case 'datos': return datos(u);
    case 'histCampana': return { historial: histCampana(u, p.id) };
    case 'cambiar': return conLock(() => cambiar(u, p.id, p.campo, p.valor));
    case 'agregar': return conLock(() => agregar(u, p.valores || {}));
    case 'transferir': return conLock(() => transferir(u, p.id, p.dest));
    default: throw new Error('Acción no reconocida.');
  }
}
function conLock(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- usuarios, link personal y sesión ----------
 * Nadie escribe PIN. Cada perfil tiene una llave al azar (columna Llave de _Usuarios)
 * y un link personal = TABLERO_URL + '#k=' + llave. Al abrir el link, la laptop queda
 * recordada DIAS_SESION días. Si un link se comparte por error, "Nuevo link" cambia la
 * llave y cierra la sesión en todas las laptops donde se usó el anterior.
 */
const COLS = ['Perfil', 'Hoja', 'Rol', 'Llave', 'Link'];
const ROLES = ['advisor', 'admin', 'presentacion'];
const DIAS_SESION = 180;
function secreto() { const pr = PropertiesService.getScriptProperties(); let s = pr.getProperty('SECRETO'); if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); pr.setProperty('SECRETO', s); } return s; }
function firmar(t) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(t, secreto())); }
function llaveNueva() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function linkDe(llave) { return TABLERO_URL + '#k=' + llave; }
function hojaUsuarios() { const sh = SpreadsheetApp.getActive().getSheetByName(USUARIOS); if (!sh) throw new Error('Falta la hoja _Usuarios. Ejecuta configurar().'); return sh; }
function usuarios() {
  return hojaUsuarios().getDataRange().getValues().slice(1).map((r, i) => ({ fila: i + 2, perfil: String(r[0] || '').trim(), hoja: String(r[1] || '').trim(), rol: String(r[2] || 'advisor').trim().toLowerCase(), llave: String(r[3] || '').trim() })).filter(u => u.perfil);
}
/** Pestañas de advisors, en el orden de _Usuarios. */
function hojas() { const out = []; usuarios().forEach(u => { if (u.rol === 'advisor' && u.hoja && out.indexOf(u.hoja) < 0) out.push(u.hoja); }); return out; }
function sesion(u) {
  const exp = Date.now() + DIAS_SESION * 864e5;
  return { perfil: u.perfil, rol: u.rol, hoja: u.rol === 'advisor' ? u.hoja : '', token: exp + '.' + firmar(u.perfil + '|' + exp + '|' + u.llave), exp: exp };
}
function entrar(llave) {
  llave = String(llave || '').trim();
  const u = llave.length >= 20 && usuarios().find(x => x.llave === llave && ROLES.indexOf(x.rol) >= 0);
  if (!u) throw new Error('Este link ya no funciona. Pídele a Andrea tu link nuevo.');
  return sesion(u);
}
function verificar(perfil, token) {
  const [exp, sig] = String(token || '').split('.'); const u = usuarios().find(x => x.perfil === perfil);
  if (!u || ROLES.indexOf(u.rol) < 0 || !u.llave || !exp || +exp < Date.now() || sig !== firmar(perfil + '|' + exp + '|' + u.llave)) throw new Error('Tu sesión se cerró. Vuelve a abrir tu link personal.');
  return u;
}

/* ---------- equipo: solo admin ---------- */
function soloAdmin(u) { if (u.rol !== 'admin') throw new Error('Solo Andrea puede administrar perfiles.'); }
function equipo(u) {
  soloAdmin(u);
  return { equipo: usuarios().map(x => ({ perfil: x.perfil, hoja: x.hoja, rol: x.rol, llave: x.llave })), hojas: hojas() };
}
function crearPerfil(u, p) {
  soloAdmin(u);
  const nombre = String(p.nombre || '').trim().replace(/\s+/g, ' '), rol = String(p.rol || 'advisor').toLowerCase();
  if (!nombre || nombre.length > 40 || /^_/.test(nombre)) throw new Error('Escribe un nombre de hasta 40 letras.');
  if (ROLES.indexOf(rol) < 0) throw new Error('Rol no válido.');
  const ss = SpreadsheetApp.getActive();
  if (usuarios().some(x => x.perfil.toLowerCase() === nombre.toLowerCase())) throw new Error('Ya existe un perfil con ese nombre.');
  let hoja = '';
  if (rol === 'advisor') {
    hoja = nombre;
    if (!ss.getSheetByName(hoja)) {
      // La pestaña nueva copia columnas, listas desplegables, casillas y formato de una advisor existente, sin sus datos.
      const base = ss.getSheetByName(String(p.plantilla || '')) || ss.getSheetByName(hojas()[0] || '');
      if (!base) throw new Error('No hay una pestaña de advisor para copiar las columnas.');
      const nueva = base.copyTo(ss).setName(hoja);
      const n = Math.max(nueva.getMaxRows() - 1, 1);
      nueva.getRange(2, 1, n, nueva.getMaxColumns()).clearContent();
      ss.setActiveSheet(nueva); ss.moveActiveSheet(ss.getSheetByName(base.getName()).getIndex() + 1);
    }
  }
  const llave = llaveNueva(), sh = hojaUsuarios();
  sh.appendRow([nombre, hoja, rol, llave, linkDe(llave)]);
  return { perfil: { perfil: nombre, hoja: hoja, rol: rol, llave: llave } };
}
function nuevoLink(u, perfil) {
  soloAdmin(u);
  const x = usuarios().find(y => y.perfil === perfil); if (!x) throw new Error('No encuentro ese perfil.');
  const llave = llaveNueva(); hojaUsuarios().getRange(x.fila, 4, 1, 2).setValues([[llave, linkDe(llave)]]);
  return { perfil: { perfil: x.perfil, hoja: x.hoja, rol: x.rol, llave: llave } };
}

/* ---------- lectura ---------- */
function encab(sh) {
  const last = Math.max(sh.getLastColumn(), 1); let h = sh.getRange(1, 1, 1, last).getValues()[0].map((v, i) => { const t = String(v === null ? '' : v).trim(); return (i === 0 && (t === '' || t === '0')) ? 'Puntaje' : t; });
  if (h.indexOf(IDC) < 0) { sh.getRange(1, last + 1).setValue(IDC); sh.hideColumns(last + 1); h.push(IDC); }
  return h;
}
function vacia(r, h) { return CLAVES.every(k => { const i = h.indexOf(k); return i < 0 || r[i] === '' || r[i] === null; }); }
function limpiar(v) { return v instanceof Date ? Utilities.formatDate(v, ZONA(), 'yyyy-MM-dd') : v; }
function leer(nombre) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nombre); if (!sh) throw new Error('No encuentro la hoja ' + nombre + '.');
  const h = encab(sh), idc = h.indexOf(IDC), n = Math.max(sh.getLastRow() - 1, 0);
  const vals = n ? sh.getRange(2, 1, n, h.length).getValues() : [];
  const filas = [];
  vals.forEach((r, k) => {
    if (vacia(r, h)) return;
    if (!r[idc]) { r[idc] = nombre.slice(0, 3).toLowerCase() + Utilities.getUuid().slice(0, 8); sh.getRange(k + 2, idc + 1).setValue(r[idc]); }
    const v = {}; h.forEach((name, i) => { if (name && name !== IDC && r[i] !== '' && r[i] !== null) v[name] = limpiar(r[i]); });
    filas.push({ id: String(r[idc]), hoja: nombre, fila: k + 2, v: v });
  });
  const listas = {};
  if (n) h.forEach((name, i) => { if (!name || name === IDC) return; const dv = sh.getRange(2, i + 1).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) listas[name] = dv.getCriteriaValues()[0].map(x => String(x).trim()); });
  Object.keys(LISTAS_STD).forEach(k => { if (!listas[k] && h.indexOf(k) >= 0) listas[k] = LISTAS_STD[k]; });
  return { sh, h, filas, cfg: { encabezados: h.filter(x => x !== IDC), listas } };
}
function histTodo() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HISTORIAL); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().map(r => { const d = r[0] instanceof Date ? r[0] : new Date(r[0]); const ts = +r[8] || d.getTime();
    return { ts, fecha: Utilities.formatDate(new Date(ts), ZONA(), "yyyy-MM-dd'T'HH:mm:ss"), perfil: r[1], hoja: r[2], id: String(r[3]), cuenta: r[4], campo: r[5], antes: limpiar(r[6]), despues: limpiar(r[7]) }; });
}
function datos(u) {
  const cfg = { hojas: {} }; let campanas = [];
  hojas().forEach(n => { const L = leer(n); cfg.hojas[n] = L.cfg; if (u.rol !== 'advisor' || n === u.hoja) campanas = campanas.concat(L.filas); });
  const mias = new Set(campanas.map(c => c.id));
  let hist = histTodo(); if (u.rol === 'advisor') hist = hist.filter(x => mias.has(x.id) || (x.perfil === u.perfil && x.campo === 'Traspaso'));
  return { config: cfg, campanas, historial: hist.slice(-1500).reverse() };
}
function histCampana(u, id) {
  if (u.rol === 'advisor' && !leer(u.hoja).filas.some(f => f.id === String(id))) throw new Error('Esta campaña no es tuya.');
  return histTodo().filter(x => x.id === String(id)).reverse();
}

/* ---------- escritura ---------- */
function etiqueta(v) { return [v.Marca, v['Campaña']].filter(Boolean).join(', '); }
function convertir(campo, valor) { if (typeof valor === 'string' && /fecha/i.test(campo) && /^\d{4}-\d{2}-\d{2}$/.test(valor)) { const p = valor.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); } return valor; }
function anotar(u, hoja, id, cuenta, campo, antes, despues) {
  const now = new Date(); SpreadsheetApp.getActive().getSheetByName(HISTORIAL).appendRow([now, u.perfil, hoja, id, cuenta, campo, antes == null ? '' : antes, despues == null ? '' : despues, now.getTime()]);
  return { ts: now.getTime(), fecha: Utilities.formatDate(now, ZONA(), "yyyy-MM-dd'T'HH:mm:ss"), perfil: u.perfil, hoja, id, cuenta, campo, antes: antes == null ? '' : antes, despues: despues == null ? '' : despues };
}
function soloAdvisor(u) { if (u.rol !== 'advisor') throw new Error('Este perfil es de solo lectura.'); }
function primeraLibre(sh, h) { const n = Math.max(sh.getLastRow() - 1, 0); const vals = n ? sh.getRange(2, 1, n, h.length).getValues() : []; let k = vals.length; for (let i = vals.length - 1; i >= 0; i--) { if (!vacia(vals[i], h)) break; k = i; } return k + 2; }
function cambiar(u, id, campo, valor) {
  soloAdvisor(u); const L = leer(u.hoja); const f = L.filas.find(x => x.id === String(id)); if (!f) throw new Error('Esta campaña no es tuya o ya no existe; recarga la página.');
  const col = L.h.indexOf(campo); if (col < 0 || campo === IDC) throw new Error('Campo no válido.');
  const antes = f.v[campo]; L.sh.getRange(f.fila, col + 1).setValue(convertir(campo, valor));
  f.v[campo] = valor; const h = anotar(u, u.hoja, f.id, etiqueta(f.v), campo, antes, valor); SpreadsheetApp.flush();
  return { campana: f, hist: h };
}
function agregar(u, valores) {
  soloAdvisor(u); const L = leer(u.hoja); const fila = primeraLibre(L.sh, L.h); const id = u.hoja.slice(0, 3).toLowerCase() + Utilities.getUuid().slice(0, 8);
  const v = {}; Object.keys(valores).forEach(k => { if (valores[k] !== '' && valores[k] != null && L.h.indexOf(k) >= 0) v[k] = valores[k]; });
  L.sh.getRange(fila, 1, 1, L.h.length).setValues([L.h.map(n => n === IDC ? id : (n in v ? convertir(n, v[n]) : ''))]);
  const h = anotar(u, u.hoja, id, etiqueta(v), 'Nueva campaña', '', etiqueta(v)); SpreadsheetApp.flush();
  return { campana: { id, hoja: u.hoja, fila, v }, hist: h };
}
function transferir(u, id, dest) {
  soloAdvisor(u); if (hojas().indexOf(dest) < 0 || dest === u.hoja) throw new Error('Elige a quién pasarla.');
  const O = leer(u.hoja); const f = O.filas.find(x => x.id === String(id)); if (!f) throw new Error('Esta campaña no es tuya.');
  const D = leer(dest); const nz = t => String(t).trim().toLowerCase(); const alias = { 'link de asana': 'link asana', 'link asana': 'link de asana', 'notas': 'notas' };
  const v = {}; Object.keys(f.v).forEach(k => { const t = D.h.find(h => h && nz(h) === nz(k)) || D.h.find(h => h && nz(h) === alias[nz(k)]); if (t) v[t] = f.v[k]; });
  const fila = primeraLibre(D.sh, D.h);
  D.sh.getRange(fila, 1, 1, D.h.length).setValues([D.h.map(n => n === IDC ? f.id : (n in v ? convertir(n, v[n]) : ''))]);
  O.sh.deleteRow(f.fila);
  const h = anotar(u, dest, f.id, etiqueta(v), 'Traspaso', u.hoja, dest); SpreadsheetApp.flush();
  return { campana: { id: f.id, hoja: dest, fila, v }, hist: h };
}

/* ---------- configuración inicial: se ejecuta una vez desde el editor ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  let us = ss.getSheetByName(USUARIOS);
  if (!us) {
    us = ss.insertSheet(USUARIOS);
    const fila = (perfil, hoja, rol) => { const k = llaveNueva(); return [perfil, hoja, rol, k, linkDe(k)]; };
    const filas = [COLS]
      .concat(ADVISORS_INICIALES.map(n => fila(n, n, 'advisor')))
      .concat([fila('Andrea', '', 'admin'), fila('Presentación', '', 'presentacion')]);
    us.getRange(1, 1, filas.length, COLS.length).setValues(filas); us.setFrozenRows(1); us.autoResizeColumns(1, 3);
  }
  let hs = ss.getSheetByName(HISTORIAL);
  if (!hs) { hs = ss.insertSheet(HISTORIAL); hs.getRange(1, 1, 1, 9).setValues([['Fecha y hora','Perfil','Hoja','ID','Campaña','Campo','Antes','Después','ts']]); hs.setFrozenRows(1); }
  // Protege _Usuarios e _Historial: solo la dueña de la hoja y quien ejecuta pueden editarlas.
  // Quien ejecuta se agrega primero como editor; si no, al quitar la edición por dominio se quitaría a sí mismo.
  [us, hs].forEach(sh => { const p = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || sh.protect().setDescription('Solo la dueña de la hoja'); const yo = Session.getEffectiveUser(); p.addEditor(yo); const fijos = [ss.getOwner() ? ss.getOwner().getEmail() : '', yo.getEmail()]; p.removeEditors(p.getEditors().filter(e => fijos.indexOf(e.getEmail()) < 0)); if (p.canDomainEdit()) p.setDomainEdit(false); });
  const faltan = hojas().filter(n => !ss.getSheetByName(n));
  if (faltan.length) throw new Error('No encuentro estas pestañas: ' + faltan.join(', ') + '. Revisa la columna Hoja de _Usuarios.');
  hojas().forEach(n => leer(n)); secreto();
  const andrea = usuarios().find(x => x.rol === 'admin');
  Logger.log('Listo. Los links personales están en la pestaña _Usuarios.' + (andrea ? ' Link de ' + andrea.perfil + ': ' + linkDe(andrea.llave) : ''));
}
