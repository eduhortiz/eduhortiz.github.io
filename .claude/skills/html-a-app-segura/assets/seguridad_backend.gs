// EJEMPLO probado (app Uniformes/EPPs). Pegar al final del Code.gs de la app y adaptar datos_, ejecutar_ y puede_ a sus acciones.
// Requiere de la app: libro_(), hoja_(), escribir_() y las funciones de negocio usadas en ejecutar_().
// SEGURIDAD: usuarios con clave cifrada (hash), sesión con token y
// permisos por rol validados AQUÍ en el servidor (no en el celular).
// ======================================================================
var USUARIOS = 'Usuarios', HEAD_USR = ['NOMBRE', 'CORREO', 'CLAVE', 'ESTADO', 'ROL', 'ACCESOS', 'CELULAR', 'ACTUALIZADO'];
var SESION_SEG = 21600;          // duración de la sesión: 6 horas (máximo de CacheService)
var MAX_INTENTOS = 5;            // intentos fallidos antes de bloquear temporalmente
var BLOQUEO_SEG = 900;           // bloqueo temporal: 15 minutos

function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function hashClave_(clave, sal) {
  var h = String(clave);
  for (var i = 0; i < 1000; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + h, Utilities.Charset.UTF_8));
  return 'h1$' + sal + '$' + h;
}
function nuevaSal_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function esHash_(v) { return /^h1\$[0-9a-f]{16}\$[0-9a-f]{64}$/.test(String(v || '')); }
function claveOk_(guardada, clave) {
  guardada = String(guardada || '');
  if (esHash_(guardada)) return hashClave_(clave, guardada.split('$')[1]) === guardada;
  return guardada !== '' && guardada === String(clave); // clave antigua sin cifrar (se cifra al entrar)
}

function leerUsuarios_(ss) {
  var sh = hoja_(ss, USUARIOS); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getDisplayValues(), H = v[0].map(function (c) { return String(c).trim().toUpperCase(); });
  var c = function () { var ks = [].slice.call(arguments); return H.findIndex(function (x) { return ks.some(function (k) { return x.indexOf(k) >= 0; }); }); };
  var cN = c('NOMBRE'), cC = c('CORREO', 'EMAIL', 'USUARIO'), cK = c('CLAVE', 'PIN', 'CONTRASE'), cE = c('ESTADO', 'ACTIVO'), cR = c('ROL'), cA = c('ACCESO'), cT = c('CELULAR', 'TELEF');
  var g = function (r, i) { return i >= 0 ? String(r[i]).replace(/^'/, '').trim() : ''; };
  return v.slice(1).filter(function (r) { return g(r, cC); }).map(function (r) {
    var e = g(r, cE);
    return { nombre: g(r, cN), correo: g(r, cC).toLowerCase(), clave: g(r, cK), estado: /bloq|^no$|inactiv|false/i.test(e) ? 'Bloqueado' : 'Activo',
      rol: /admin/i.test(g(r, cR)) ? 'Administrador' : /entrega/i.test(g(r, cR)) ? 'Operador de Entrega' : /ingreso/i.test(g(r, cR)) ? 'Operador de Ingreso' : 'Operador', accesos: g(r, cA).split(',').map(function (x) { return x.trim(); }).filter(String), celular: g(r, cT) };
  });
}
function escribirUsuarios_(ss, lista) {
  var sh = hoja_(ss, USUARIOS);
  if (!sh) sh = ss.insertSheet(USUARIOS);
  var last = Math.max(sh.getLastRow(), 1), cols = Math.max(sh.getLastColumn(), HEAD_USR.length);
  sh.getRange(1, 1, last, cols).clearContent();
  try { sh.getRange(1, 3, sh.getMaxRows(), 1).setNumberFormat('@'); sh.getRange(1, 7, sh.getMaxRows(), 1).setNumberFormat('@'); } catch (e) { }
  var rows = [HEAD_USR].concat(lista.map(function (u) { return [u.nombre, String(u.correo).toLowerCase(), String(u.clave), u.estado === 'Bloqueado' ? 'BLOQUEADO' : 'ACTIVO', u.rol, (u.accesos || []).join(','), String(u.celular || ''), new Date()]; }));
  escribir_(sh, 1, rows);
  try { sh.getRange(1, 1, 1, HEAD_USR.length).setFontWeight('bold'); sh.setFrozenRows(1); } catch (e) { }
}
// Lo que ve la app: nunca incluye la clave
function publico_(u) { return { nombre: u.nombre, correo: u.correo, estado: u.estado, rol: u.rol, accesos: u.accesos, celular: u.celular }; }

// Recibe la lista editada por el administrador. "clave" solo viene cuando se crea o cambia (6 números);
// si viene vacía se conserva la clave cifrada que ya existía.
function guardarUsuarios_(ss, lista, yo) {
  if (!Array.isArray(lista)) throw new Error('Lista de usuarios inválida.');
  var prev = {}; leerUsuarios_(ss).forEach(function (u) { prev[u.correo] = u; });
  var vistos = {};
  var nueva = lista.map(function (u) {
    var correo = String(u.correo || '').trim().toLowerCase(), orig = String(u.orig || correo).trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) throw new Error('Correo inválido: ' + correo);
    if (vistos[correo]) throw new Error('Correo repetido: ' + correo);
    vistos[correo] = true;
    var clave = String(u.clave || '').trim(), antes = prev[orig] || prev[correo];
    if (clave) { if (!/^\d{6}$/.test(clave)) throw new Error('La clave debe tener 6 números.'); clave = hashClave_(clave, nuevaSal_()); }
    else if (antes) clave = esHash_(antes.clave) ? antes.clave : hashClave_(antes.clave, nuevaSal_());
    else throw new Error('Falta la clave de ' + correo);
    var rol = ['Administrador', 'Operador de Entrega', 'Operador de Ingreso', 'Operador'].indexOf(u.rol) >= 0 ? u.rol : 'Operador de Entrega';
    return { nombre: String(u.nombre || '').trim().slice(0, 120), correo: correo, clave: clave, estado: u.estado === 'Bloqueado' ? 'Bloqueado' : 'Activo', rol: rol,
      accesos: (Array.isArray(u.accesos) ? u.accesos : []).map(String).slice(0, 20), celular: String(u.celular || '').replace(/\D/g, '').slice(0, 15) };
  });
  if (!nueva.some(function (u) { return u.rol === 'Administrador' && u.estado !== 'Bloqueado'; })) throw new Error('Debe quedar al menos un administrador activo.');
  if (!nueva.some(function (u) { return u.correo === yo.correo && u.rol === 'Administrador' && u.estado !== 'Bloqueado'; })) throw new Error('No puedes quitarte tu propio acceso de administrador.');
  escribirUsuarios_(ss, nueva);
  return { ok: true, total: nueva.length, usuarios: nueva.map(publico_) };
}

// Ejecutar UNA VEZ desde el editor de Apps Script: cifra todas las claves que estén en texto plano
function cifrarClavesExistentes() {
  var ss = libro_(), lista = leerUsuarios_(ss), n = 0;
  lista.forEach(function (u) { if (u.clave && !esHash_(u.clave)) { u.clave = hashClave_(u.clave, nuevaSal_()); n++; } });
  if (n) escribirUsuarios_(ss, lista);
  Logger.log('Claves cifradas: ' + n);
}

// ---- Sesiones
function crearSesion_(u) {
  var tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('ses_' + tok, JSON.stringify({ correo: u.correo }), SESION_SEG);
  return tok;
}
// Devuelve el usuario vigente del token (relee la hoja: si lo bloquean o cambian de rol, aplica al instante)
function sesion_(tok) {
  if (!/^[0-9a-f]{64}$/.test(String(tok || ''))) return null;
  var c = CacheService.getScriptCache(), raw = c.get('ses_' + tok); if (!raw) return null;
  var s = JSON.parse(raw), u = leerUsuarios_(libro_()).filter(function (x) { return x.correo === s.correo; })[0];
  if (!u || u.estado === 'Bloqueado') { c.remove('ses_' + tok); return null; }
  c.put('ses_' + tok, raw, SESION_SEG); // renueva mientras se use
  return u;
}

function login_(d) {
  var correo = String(d.correo || '').trim().toLowerCase(), clave = String(d.clave || '').trim();
  if (!correo || !/^\d{6}$/.test(clave)) return { ok: false, error: 'Correo o clave incorrectos.' };
  var c = CacheService.getScriptCache(), kI = 'int_' + Utilities.base64EncodeWebSafe(correo), n = Number(c.get(kI) || 0);
  if (n >= MAX_INTENTOS) return { ok: false, error: 'Demasiados intentos. Espera 15 minutos e inténtalo de nuevo.' };
  var ss = libro_(), lista = leerUsuarios_(ss), u = lista.filter(function (x) { return x.correo === correo; })[0];
  if (!u || !claveOk_(u.clave, clave)) { c.put(kI, String(n + 1), BLOQUEO_SEG); Utilities.sleep(800); return { ok: false, error: 'Correo o clave incorrectos.' }; }
  if (u.estado === 'Bloqueado') return { ok: false, error: 'Tu acceso está bloqueado. Comunícate con el administrador.' };
  c.remove(kI);
  if (!esHash_(u.clave)) { u.clave = hashClave_(clave, nuevaSal_()); escribirUsuarios_(ss, lista); }
  return { ok: true, token: crearSesion_(u), user: publico_(u), claveDebil: /^(\d)\1{5}$|^123456$|^654321$/.test(clave) };
}

// ---- Permisos por rol (en el servidor)
function puede_(u, accion, d) {
  if (u.rol === 'Administrador') return true;
  var ent = u.rol === 'Operador de Entrega' || u.rol === 'Operador', ing = u.rol === 'Operador de Ingreso' || u.rol === 'Operador';
  if (accion === 'datos' || accion === 'logout' || accion === 'cambiarClave') return true;
  if (accion === 'entrega') return ent && !(d && d.regularizacion);
  if (accion === 'ingreso' || accion === 'nuevoProducto') return ing;
  return false; // usuarios, eliminar*, ajustarStock, eliminarProducto, completarTVU: solo administrador
}

// Evita que un texto escrito en la app se ejecute como fórmula en la hoja (=, +, @)
function limpiar_(v) {
  if (typeof v === 'string') return /^[=+@\t\r]/.test(v) ? "'" + v : v.slice(0, 2000);
  if (Array.isArray(v)) return v.map(limpiar_);
  if (v && typeof v === 'object') { var o = {}; Object.keys(v).forEach(function (k) { o[k] = limpiar_(v[k]); }); return o; }
  return v;
}

function registrarIngreso_(ss, d) {
  var ing = hoja_(ss, 'Ingresos') || ss.insertSheet('Ingresos');
  if (ing.getLastRow() === 0) {
    ing.appendRow(['N° Ingreso','Fecha','Tipo ingreso','Recibido por',
      'TIPO','Codigo SAP','Descripcion','Talla','GENERO','Cantidad',
      'Costo S/.','Subtotal S/.','Observaciones','Registrado']);
    try { ing.getRange('F:F').setNumberFormat('@'); ing.getRange('H:H').setNumberFormat('@'); } catch (e) { }
  }
  var st = hojaStock_(ss);
  var lastT = ultimaFila_(st);
  var data = lastT > 1 ? [[]].concat(leerFilas_(st, 2, lastT - 1)) : [[]], fila = {};
  for (var i = 1; i < data.length; i++) fila[clave_(data[i][1], data[i][2], data[i][3])] = i + 1;
  (d.items || []).forEach(function (it) {
    var cant = Number(it.cantidad);
    if (!(cant > 0) || cant > 100000) throw new Error('Cantidad inválida en ' + it.descripcion);
    ing.appendRow([d.id, d.fecha, d.motivo, d.recibido, it.tipo, String(it.codigo),
      it.descripcion, String(it.talla), it.genero, cant, num_(it.costo),
      cant * num_(it.costo), d.obs, new Date()]);
    var r = fila[clave_(it.codigo, it.descripcion, it.talla)];
    if (r) {
      var c = st.getRange(r, 6), nuevo = num_(c.getValue()) + cant;
      c.setValue(nuevo);
      st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue()));
      st.getRange(r, 9).setValue(new Date());
    }
  });
}

function datos_(ss, u) {
  var sh = hojaStock_(ss), last = ultimaFila_(sh), stock = [];
  if (last > 1) stock = leerFilas_(sh, 2, last - 1).map(function (r) {
    return { codigo: String(r[1]).trim(), descripcion: String(r[2]).trim(), talla: txt_(r[3]), stock: num_(r[5]) };
  });
  var out = { ok: true, version: 7, user: publico_(u), catalogo: leerCatalogo_(ss), stock: stock, personal: leerPersonal_(ss), entregas: leerEntregas_(ss), uxa: leerUxA_(ss), ingresos: leerIngresos_(ss) };
  if (u.rol === 'Administrador') out.usuarios = leerUsuarios_(ss).map(publico_);
  return out;
}

// Todas las acciones que modifican la hoja se ejecutan con bloqueo para evitar choques entre equipos
function ejecutar_(accion, d, u) {
  var ss = libro_();
  if (accion === 'datos') return datos_(ss, u);
  if (accion === 'logout') { CacheService.getScriptCache().remove('ses_' + d.__tok); return { ok: true }; }
  var lk = LockService.getScriptLock(); lk.waitLock(20000);
  try {
    switch (accion) {
      case 'cambiarClave': {
        var lista = leerUsuarios_(ss), me = lista.filter(function (x) { return x.correo === u.correo; })[0];
        if (!claveOk_(me.clave, String(d.actual || ''))) return { ok: false, error: 'La clave actual no es correcta.' };
        if (!/^\d{6}$/.test(String(d.nueva || ''))) return { ok: false, error: 'La nueva clave debe tener 6 números.' };
        me.clave = hashClave_(String(d.nueva), nuevaSal_()); escribirUsuarios_(ss, lista); return { ok: true };
      }
      case 'entrega': d.entregadoPor = u.nombre; registrarEntrega_(ss, d); return { ok: true };
      case 'ingreso': d.recibido = u.nombre; registrarIngreso_(ss, d); return { ok: true };
      case 'nuevoProducto': return agregarProducto_(ss, d);
      case 'ajustarStock': return ajustarStock_(ss, d);
      case 'eliminarProducto': return eliminarProducto_(ss, d);
      case 'completarTVU': var h = completarTVU(); return { ok: true, completados: h.length, detalle: h };
      case 'eliminarEntrega': return { ok: true, borradas: eliminarPorId_(ss, ENTREGAS, String(d.id), +1) };
      case 'eliminarIngreso': return { ok: true, borradas: eliminarPorId_(ss, 'Ingresos', String(d.id), -1) };
      case 'eliminarItemEntrega': return { ok: true, borradas: eliminarItemEntrega_(ss, d) };
      case 'eliminarItemIngreso': return { ok: true, borradas: eliminarItemIngreso_(ss, d) };
      case 'guardarUsuarios': return guardarUsuarios_(ss, d.usuarios, u);
    }
    return { ok: false, error: 'Acción desconocida.' };
  } finally { lk.releaseLock(); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// La app envía todo por POST: { accion, token, data }. Sin sesión válida no se lee ni modifica nada.
function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}'), accion = String(b.accion || '');
    if (accion === 'login') return json_(login_(b.data || {}));
    var u = sesion_(b.token);
    if (!u) return json_({ ok: false, auth: false, error: 'Tu sesión expiró. Vuelve a ingresar.' });
    var d = limpiar_(b.data || {});
    if (!puede_(u, accion, d)) return json_({ ok: false, error: 'No tienes permiso para esta acción.' });
    d.__tok = b.token;
    return json_(ejecutar_(accion, d, u));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Error del servidor: ' + String(err.message || err).slice(0, 200) });
  }
}

// Las lecturas por GET ya no entregan datos (antes cualquiera con el enlace podía verlos)
function doGet() { return json_({ ok: true, app: 'Registro de Uniformes y EPPs', version: 7 }); }
