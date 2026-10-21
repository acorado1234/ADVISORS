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
    case 'nota': return conLock(() => nota(u, p.id, p.texto));
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
/** Lee _Usuarios por el NOMBRE de cada columna (no por su posición), así no se rompe si alguien mueve columnas. */
function usuarios() {
  const vals = hojaUsuarios().getDataRange().getValues(); const h = (vals[0] || []).map(x => String(x).trim().toLowerCase());
  const c = n => h.indexOf(n.toLowerCase());
  if (c('Perfil') < 0 || c('Llave') < 0) throw new Error('La pestaña _Usuarios está en el formato anterior. Ejecuta configurar() en Apps Script.');
  const v = (r, n) => c(n) < 0 ? '' : String(r[c(n)] == null ? '' : r[c(n)]).trim();
  return vals.slice(1).map((r, i) => ({ fila: i + 2, perfil: v(r, 'Perfil'), hoja: v(r, 'Hoja'), rol: (v(r, 'Rol') || 'advisor').toLowerCase(), llave: v(r, 'Llave') })).filter(u => u.perfil);
}
function colUsuarios(n) { const h = hojaUsuarios().getRange(1, 1, 1, hojaUsuarios().getLastColumn()).getValues()[0].map(x => String(x).trim().toLowerCase()); return h.indexOf(n.toLowerCase()) + 1; }
function filaUsuario(o) { const sh = hojaUsuarios(); const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim()); return h.map(n => (n in o ? o[n] : '')); }
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
  sh.appendRow(filaUsuario({ Perfil: nombre, Hoja: hoja, Rol: rol, Llave: llave, Link: linkDe(llave) }));
  return { perfil: { perfil: nombre, hoja: hoja, rol: rol, llave: llave }, equipo: equipo(u) };
}
function nuevoLink(u, perfil) {
  soloAdmin(u);
  const x = usuarios().find(y => y.perfil === perfil); if (!x) throw new Error('No encuentro ese perfil.');
  const llave = llaveNueva(), sh = hojaUsuarios(); sh.getRange(x.fila, colUsuarios('Llave')).setValue(llave);
  const cl = colUsuarios('Link'); if (cl) sh.getRange(x.fila, cl).setValue(linkDe(llave));
  return { perfil: { perfil: x.perfil, hoja: x.hoja, rol: x.rol, llave: llave }, equipo: equipo(u) };
}

/* ---------- modelo estándar de campaña ----------
 * Cada advisor arma su pestaña a su manera: columnas con otro nombre, en otro orden, o que no tiene.
 * El tablero trabaja con UN solo modelo para todas (CAMPOS) y aquí se traduce cada pestaña a ese modelo:
 *  - los nombres se comparan sin importar mayúsculas, tildes ni espacios ("notas" = "Notas");
 *  - ALIAS cubre nombres distintos para lo mismo ("link de asana" = "Link Asana");
 *  - si a una pestaña le falta un campo, se crea esa columna al final la primera vez que alguien lo llena;
 *  - si una pestaña tiene dos columnas para lo mismo (Cintia: "notas" y "Notas"), se leen las dos
 *    y se escribe en la principal (la que se llama exactamente igual que el campo).
 * Nunca se borra, se mueve ni se renombra una columna de nadie.
 */
const CAMPOS = ['Puntaje', 'Errores operativos', 'Status', 'Agencia', 'Marca', 'Campaña', 'Propuesta', 'Fecha último envio de propuesta', 'Link propuesta final', 'Link Asana',
  'Activación', 'Fecha de activación', 'Fecha de finalización', 'Creación de comunidad', 'Contratos firmados', 'Link contratos', 'Status calendario', 'Link calendario',
  'Status de producción', 'Notas', 'Link contenidos', 'Contenido reposteado', 'Solicitud de reporte', 'Dashboard', 'Fecha envio reporte',
  'ODC enviada', 'Facturas recibidas', 'Influencers pagados', 'Link documentación de pago', 'Campaña facturada', 'Pago recibido'];
const ALIAS = { 'link de asana': 'Link Asana', 'asana': 'Link Asana', 'nota': 'Notas', 'comentarios': 'Notas', 'link de calendario': 'Link calendario', 'fecha de envio de reporte': 'Fecha envio reporte', 'link de propuesta final': 'Link propuesta final' };
const BOOLS = ['Creación de comunidad', 'Contratos firmados', 'Contenido reposteado', 'Solicitud de reporte', 'ODC enviada', 'Facturas recibidas', 'Influencers pagados', 'Campaña facturada', 'Pago recibido'];
const nz = t => String(t == null ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
const CANON = {}; CAMPOS.forEach(c => { CANON[nz(c)] = c; }); Object.keys(ALIAS).forEach(a => { CANON[nz(a)] = ALIAS[a]; });
function canonDe(nombre, i) { if (i === 0 && (nombre === '' || nombre === '0')) return 'Puntaje'; return CANON[nz(nombre)] || null; }

/* ---------- lectura ---------- */
function encab(sh) {
  const last = Math.max(sh.getLastColumn(), 1); let h = sh.getRange(1, 1, 1, last).getValues()[0].map((v, i) => { const t = String(v === null ? '' : v).trim(); return (i === 0 && (t === '' || t === '0')) ? 'Puntaje' : t; });
  if (h.indexOf(IDC) < 0) { sh.getRange(1, last + 1).setValue(IDC); sh.hideColumns(last + 1); h.push(IDC); }
  return h;
}
/** Columnas de cada campo en la pestaña: { campo: [col, ...] } (base 0), la principal primero. */
function mapa(h) {
  const m = {};
  h.forEach((nombre, i) => { if (!nombre || nombre === IDC) return; const c = canonDe(nombre, i) || nombre; (m[c] = m[c] || []).push(i); });
  Object.keys(m).forEach(c => { if (m[c].length > 1) m[c].sort((a, b) => ((h[b] === c) - (h[a] === c)) || a - b); });
  return m;
}
function limpiar(v) { return v instanceof Date ? Utilities.formatDate(v, ZONA(), 'yyyy-MM-dd') : (typeof v === 'string' ? v.trim() : v); }
function lleno(x) { return x !== '' && x !== null && x !== undefined && !(typeof x === 'string' && !x.trim()); }
function vacia(r, m) { return CLAVES.every(k => !(m[k] || []).some(i => lleno(r[i]))); }
/** Valor de un campo; si la pestaña tiene dos columnas para él, se unen (separadas por una línea en blanco). */
function valorDe(r, cols) { const xs = cols.map(i => limpiar(r[i])).filter(lleno); if (!xs.length) return undefined; if (xs.length === 1 || typeof xs[0] !== 'string') return xs[0]; return xs.join('\n\n'); }
function leer(nombre) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nombre); if (!sh) throw new Error('No encuentro la hoja ' + nombre + '.');
  const h = encab(sh), m = mapa(h), idc = h.indexOf(IDC), n = Math.max(sh.getLastRow() - 1, 0);
  const vals = n ? sh.getRange(2, 1, n, h.length).getValues() : [];
  const filas = []; let nuevos = false;
  vals.forEach((r, k) => {
    if (vacia(r, m)) return;
    if (!r[idc]) { r[idc] = nombre.slice(0, 3).toLowerCase() + Utilities.getUuid().slice(0, 8); nuevos = true; }
    const v = {}; Object.keys(m).forEach(c => { const x = valorDe(r, m[c]); if (x !== undefined) v[c] = x; });
    filas.push({ id: String(r[idc]), hoja: nombre, fila: k + 2, v: v });
  });
  // Los ID nuevos se escriben de una sola vez (una escritura por campaña hacía muy lenta la primera carga).
  if (nuevos) sh.getRange(2, idc + 1, vals.length, 1).setValues(vals.map(r => [r[idc] || '']));
  // Las listas son las mismas para todas (LISTAS_STD). Solo una columna propia, fuera del modelo, usa su propia lista.
  const listas = {}; Object.keys(LISTAS_STD).forEach(k => { listas[k] = LISTAS_STD[k]; });
  const extras = Object.keys(m).filter(c => CAMPOS.indexOf(c) < 0);
  if (n) extras.forEach(c => { const dv = sh.getRange(2, m[c][0] + 1).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) listas[c] = dv.getCriteriaValues()[0].map(x => String(x).trim()); });
  return { sh, h, m, filas, cfg: { encabezados: CAMPOS.concat(extras), listas } };
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
/**
 * Columna (base 1) donde se escribe un campo. Si la pestaña no la tiene y crear = true, la agrega al final
 * con el nombre estándar (con casillas o lista desplegable si corresponde). libre = true permite crear
 * columnas fuera del modelo (solo en traspasos, para no perder datos).
 */
function colDe(L, campo, crear, libre) {
  if (L.m[campo]) return L.m[campo][0] + 1;
  if (!crear) return 0;
  if (!libre && CAMPOS.indexOf(campo) < 0) throw new Error('Campo no válido.');
  const c = L.sh.getLastColumn() + 1; L.sh.getRange(1, c).setValue(campo);
  try {
    const rg = L.sh.getRange(2, c, Math.max(L.sh.getMaxRows() - 1, 1), 1);
    if (BOOLS.indexOf(campo) >= 0) rg.insertCheckboxes();
    else if (LISTAS_STD[campo]) rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LISTAS_STD[campo], true).setAllowInvalid(true).build());
  } catch (e) { /* el formato es opcional: el dato se guarda igual */ }
  L.h[c - 1] = campo; L.m[campo] = [c - 1]; return c;
}
function primeraLibre(sh, m) { const n = Math.max(sh.getLastRow() - 1, 0); const vals = n ? sh.getRange(2, 1, n, Math.max(sh.getLastColumn(), 1)).getValues() : []; let k = vals.length; for (let i = vals.length - 1; i >= 0; i--) { if (!vacia(vals[i], m)) break; k = i; } return k + 2; }
function cambiar(u, id, campo, valor) {
  soloAdvisor(u); const L = leer(u.hoja); const f = L.filas.find(x => x.id === String(id)); if (!f) throw new Error('Esta campaña no es tuya o ya no existe; recarga la página.');
  if (campo === 'Notas') throw new Error('Las notas se agregan desde la bitácora.');
  if (!campo || campo === IDC || (CAMPOS.indexOf(campo) < 0 && !L.m[campo])) throw new Error('Campo no válido.');
  const col = colDe(L, campo, true);
  const antes = f.v[campo]; L.sh.getRange(f.fila, col).setValue(convertir(campo, valor));
  f.v[campo] = valor; const h = anotar(u, u.hoja, f.id, etiqueta(f.v), campo, antes, valor); SpreadsheetApp.flush();
  return { campana: f, hist: h };
}
/* Bitácora: cada nota se suma ARRIBA en la columna Notas principal, con fecha, hora y quién la escribió.
 * Formato de cada línea: "[20/10/2026 14:30, Mafer] texto". Lo que ya estaba escrito se conserva debajo. */
function nota(u, id, texto) {
  soloAdvisor(u); texto = String(texto || '').replace(/\s*\n+\s*/g, ' / ').trim();
  if (!texto) throw new Error('Escribe la nota primero.'); if (texto.length > 1000) throw new Error('La nota es muy larga; resúmela en menos de 1000 letras.');
  const L = leer(u.hoja); const f = L.filas.find(x => x.id === String(id)); if (!f) throw new Error('Esta campaña no es tuya o ya no existe; recarga la página.');
  const col = colDe(L, 'Notas', true), antes = String(L.sh.getRange(f.fila, col).getValue() || '').trim();
  const linea = '[' + Utilities.formatDate(new Date(), ZONA(), 'dd/MM/yyyy HH:mm') + ', ' + u.perfil + '] ' + texto;
  L.sh.getRange(f.fila, col).setValue(antes ? linea + '\n' + antes : linea);
  const r = L.sh.getRange(f.fila, 1, 1, L.sh.getLastColumn()).getValues()[0]; f.v.Notas = valorDe(r, L.m.Notas);
  const h = anotar(u, u.hoja, f.id, etiqueta(f.v), 'Nota', '', texto); SpreadsheetApp.flush();
  return { campana: f, hist: h };
}
/** Escribe una campaña completa en una fila, respetando lo que ya haya en las demás columnas (casillas, formato). */
function escribirFila(L, fila, id, v, libre) {
  Object.keys(v).forEach(k => colDe(L, k, true, libre));
  const ncol = L.sh.getLastColumn(), row = L.sh.getRange(fila, 1, 1, ncol).getValues()[0];
  row[L.h.indexOf(IDC)] = id; Object.keys(v).forEach(k => { row[L.m[k][0]] = convertir(k, v[k]); });
  L.sh.getRange(fila, 1, 1, ncol).setValues([row]);
}
function agregar(u, valores) {
  soloAdvisor(u); const L = leer(u.hoja); const fila = primeraLibre(L.sh, L.m); const id = u.hoja.slice(0, 3).toLowerCase() + Utilities.getUuid().slice(0, 8);
  const v = {}; Object.keys(valores).forEach(k => { if (lleno(valores[k]) && k !== 'Notas' && (CAMPOS.indexOf(k) >= 0 || L.m[k])) v[k] = valores[k]; });
  escribirFila(L, fila, id, v);
  const h = anotar(u, u.hoja, id, etiqueta(v), 'Nueva campaña', '', etiqueta(v)); SpreadsheetApp.flush();
  return { campana: { id, hoja: u.hoja, fila, v }, hist: h };
}
function transferir(u, id, dest) {
  soloAdvisor(u); if (hojas().indexOf(dest) < 0 || dest === u.hoja) throw new Error('Elige a quién pasarla.');
  const O = leer(u.hoja); const f = O.filas.find(x => x.id === String(id)); if (!f) throw new Error('Esta campaña no es tuya.');
  // Se pasa campo por campo con el modelo estándar; si a la otra pestaña le falta una columna, se crea para no perder nada.
  const D = leer(dest), v = Object.assign({}, f.v), fila = primeraLibre(D.sh, D.m);
  escribirFila(D, fila, f.id, v, true);
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
  } else {
    // Migración: si _Usuarios viene del formato con PIN (Perfil, PIN, Hoja, Rol), se pasa a links
    // personales conservando perfiles, pestañas y roles. El PIN se elimina.
    const vals = us.getDataRange().getValues(); const h = vals[0].map(x => String(x).trim().toLowerCase());
    if (h.indexOf('llave') < 0) {
      const c = n => h.indexOf(n);
      const filas = [COLS].concat(vals.slice(1).filter(r => String(r[c('perfil')] || '').trim()).map(r => {
        const k = llaveNueva(); return [String(r[c('perfil')]).trim(), c('hoja') < 0 ? '' : String(r[c('hoja')] || '').trim(), c('rol') < 0 ? 'advisor' : (String(r[c('rol')] || '').trim().toLowerCase() || 'advisor'), k, linkDe(k)];
      }));
      us.clearContents(); us.getRange(1, 1, us.getMaxRows(), us.getMaxColumns()).setNumberFormat('General');
      us.getRange(1, 1, filas.length, COLS.length).setValues(filas); us.setFrozenRows(1); us.autoResizeColumns(1, 3);
      Logger.log('_Usuarios pasó del formato con PIN a links personales (' + (filas.length - 1) + ' perfiles).');
    }
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

/* ---------- limpieza de pruebas: se ejecuta a mano desde el editor ----------
 * Borra TODO lo que hicieron los perfiles cuyo nombre empieza con "Prueba":
 *  - su fila en _Usuarios (su link deja de servir),
 *  - su pestaña (solo si el nombre de la pestaña también empieza con "Prueba"),
 *  - las campañas que ellos crearon, aunque las hayan pasado a otra advisor,
 *  - sus movimientos en _Historial.
 * No toca nada de Mafer, Adriana, Cintia ni de ningún otro perfil.
 */
function borrarPruebas() {
  const ss = SpreadsheetApp.getActive(), es = n => /^prueba/i.test(String(n || '').trim());
  const U = usuarios(), pruebas = U.filter(u => es(u.perfil));
  const hs = ss.getSheetByName(HISTORIAL);
  const H = hs && hs.getLastRow() > 1 ? hs.getRange(2, 1, hs.getLastRow() - 1, 9).getValues() : [];
  const nombres = pruebas.map(u => u.perfil);
  // Campañas creadas por un perfil de prueba (por su ID), estén donde estén ahora.
  const ids = new Set(H.filter(r => es(r[1]) && r[5] === 'Nueva campaña').map(r => String(r[3])));
  let campanas = 0;
  hojas().filter(n => !es(n)).forEach(n => {
    const sh = ss.getSheetByName(n); if (!sh) return; const h = encab(sh), c = h.indexOf(IDC), last = sh.getLastRow(); if (last < 2) return;
    const v = sh.getRange(2, c + 1, last - 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) if (ids.has(String(v[i][0]))) { sh.deleteRow(i + 2); campanas++; }
  });
  let movs = 0;
  for (let i = H.length - 1; i >= 0; i--) if (es(H[i][1]) || ids.has(String(H[i][3]))) { hs.deleteRow(i + 2); movs++; }
  let pestanas = 0;
  pruebas.map(u => u.hoja).filter(n => n && es(n)).forEach(n => { const sh = ss.getSheetByName(n); if (sh) { ss.deleteSheet(sh); pestanas++; } });
  const us = hojaUsuarios();
  pruebas.map(u => u.fila).sort((a, b) => b - a).forEach(f => us.deleteRow(f));
  Logger.log('Pruebas borradas. Perfiles: ' + (nombres.join(', ') || 'ninguno') + '. Pestañas: ' + pestanas + '. Campañas en otras pestañas: ' + campanas + '. Movimientos del historial: ' + movs + '.');
}
