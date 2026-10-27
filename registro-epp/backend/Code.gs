// ID de tu hoja de Google (la parte entre /d/ y /edit en la URL de la hoja)
var SHEET_ID = 'PEGA_AQUI_EL_ID_DE_TU_HOJA';
var CATALOGO = 'Uniformes/EPPs';
var STOCK = 'Stock disponible';
var PERSONAL = 'data de personal';
var ENTREGAS = 'Entregas';
var HEAD_STOCK = ['TIPO','Codigo SAP','Descripcion','Talla','GENERO',
  'Stock','Costo S/.','Valor S/.','Ultima actualizacion'];

function libro_() { return SpreadsheetApp.openById(SHEET_ID); }

// Busca una hoja ignorando mayúsculas, tildes y espacios al inicio/final ("entregas " = "Entregas")
function hoja_(ss, nombre) {
  var sh = ss.getSheetByName(nombre); if (sh) return sh;
  var k = function (s) { return String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); }, n = k(nombre), hs = ss.getSheets();
  for (var i = 0; i < hs.length; i++) if (k(hs[i].getName()) === n) return hs[i];
  return null;
}

// Convierte 1.2 / "1,2" / "51.9" a número real
// Convierte "S/ 1,234.50", "12,50", "S/. 4.5" o 12.5 en número
function num_(v) {
  if (typeof v === 'number') return v;
  if (v instanceof Date || v === null || v === undefined) return 0;
  var s = String(v).replace(/[^0-9,.\-]/g, '');
  if (!s) return 0;
  var c = s.lastIndexOf(','), d = s.lastIndexOf('.');
  if (c > -1 && d > -1) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (c > -1) s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
// La talla siempre como texto (evita que 8.5 se vuelva fecha)
function txt_(v) {
  if (v instanceof Date) return v.getDate() + '.' + (v.getMonth() + 1);
  return String(v === null || v === undefined ? '' : v).trim();
}
function clave_(codigo, desc, talla) {
  codigo = String(codigo || '').trim();
  return codigo ? codigo : (String(desc).trim() + '|' + txt_(talla));
}

function hojaCatalogo_(ss) {
  var sh = ss.getSheetByName(CATALOGO);
  if (sh) return sh;
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var n = hojas[i].getName();
    if (n === STOCK || n === 'Ingresos') continue;
    var v = hojas[i].getRange(1, 1, Math.min(10, hojas[i].getMaxRows()), 10).getValues();
    for (var j = 0; j < v.length; j++)
      if (v[j].some(function (c) { return String(c).trim().toUpperCase() === 'TIPO'; })) return hojas[i];
  }
  throw new Error('No se encontró la hoja del catálogo "' + CATALOGO + '"');
}

function leerCatalogo_(ss) {
  var cat = hojaCatalogo_(ss).getDataRange().getValues();
  var hi = cat.findIndex(function (r) {
    return r.some(function (c) { return String(c).trim().toUpperCase() === 'TIPO'; });
  });
  var H = cat[hi].map(function (c) { return String(c).toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var col_tipo = col('TIPO'), col_cod = col('CODIGO'), col_desc = col('DESCRIP'),
      col_talla = col('TALLA'), col_gen = col('GENERO'), col_tvu = col('TVU'), col_costo = col('COSTO');
  var disp = hojaCatalogo_(ss).getDataRange().getDisplayValues().slice(hi + 1);
  return cat.slice(hi + 1).map(function (r, i) { return [r, i]; }).filter(function (x) { return x[0][col_desc]; }).map(function (x) {
    var r = x[0], i = x[1];
    return { tipo: String(r[col_tipo]).trim(), codigo: String(r[col_cod]).trim(),
      descripcion: String(r[col_desc]).trim(), talla: txt_(r[col_talla]),
      genero: String(r[col_gen]).trim(), tvu: num_(r[col_tvu]), costo: num_(r[col_costo]) || num_(disp[i][col_costo]) };
  });
}

// Las hojas con formato de "Tabla" solo permiten operar columna por columna
function formatos_(sh) {
  var n = Math.max(sh.getMaxRows() - 1, 1);
  var f = [[2, '@'], [4, '@'], [6, '0'], [7, '0.00'], [8, '0.00'], [9, 'dd/mm/yyyy hh:mm']];
  f.forEach(function (x) {
    try { sh.getRange(2, x[0], n, 1).setNumberFormat(x[1]); } catch (e) { }
  });
}

// Escribe un bloque de filas columna por columna (compatible con Tablas)
function escribir_(sh, fila, rows) {
  if (!rows.length) return;
  for (var c = 0; c < rows[0].length; c++) {
    sh.getRange(fila, c + 1, rows.length, 1).setValues(rows.map(function (r) { return [r[c]]; }));
  }
}

function leerFilas_(sh, desde, n) {
  var out = [];
  for (var i = 0; i < n; i++) out.push([]);
  for (var c = 1; c <= 9; c++) {
    var v = sh.getRange(desde, c, n, 1).getValues();
    for (var i = 0; i < n; i++) out[i].push(v[i][0]);
  }
  return out;
}

function hojaStock_(ss) {
  var sh = ss.getSheetByName(STOCK);
  if (!sh) {
    sh = ss.insertSheet(STOCK);
    escribir_(sh, 1, [HEAD_STOCK]);
    try { sh.getRange(1, 1, 1, 9).setFontWeight('bold'); sh.setFrozenRows(1); } catch (e) { }
  }
  formatos_(sh);
  var lastS = ultimaFila_(sh);
  var existe = {}, cur = lastS > 1 ? [[]].concat(leerFilas_(sh, 2, lastS - 1)) : [[]];
  for (var i = 1; i < cur.length; i++) if (cur[i][2]) existe[clave_(cur[i][1], cur[i][2], cur[i][3])] = true;
  var nuevas = [];
  leerCatalogo_(ss).forEach(function (it) {
    var k = clave_(it.codigo, it.descripcion, it.talla);
    if (existe[k]) return;
    existe[k] = true;
    nuevas.push([it.tipo, it.codigo, it.descripcion, it.talla, it.genero, 0, it.costo, 0, '']);
  });
  if (nuevas.length) {
    var ini = ultimaFila_(sh) + 1;
    escribir_(sh, ini, nuevas);
  }
  recalcularValor_(sh);
  return sh;
}

function ultimaFila_(sh) {
  var v = sh.getRange(1, 3, sh.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim()) return i + 1;
  return 1;
}

function recalcularValor_(sh) {
  var last = ultimaFila_(sh);
  if (last < 2) return;
  var st = sh.getRange(2, 6, last - 1, 1).getValues();
  var co = sh.getRange(2, 7, last - 1, 1).getValues();
  sh.getRange(2, 8, last - 1, 1).setValues(st.map(function (x, i) { return [num_(x[0]) * num_(co[i][0])]; }));
}

// 1) Ejecutar una vez: crea o completa "Stock disponible"
function crearStockDisponible() { hojaStock_(libro_()); }

// 2) Ejecutar si la hoja ya existe con datos dañados (costos/tallas como fecha)
function repararStockDisponible() {
  var ss = libro_(), sh = ss.getSheetByName(STOCK);
  if (!sh) return crearStockDisponible();
  formatos_(sh);
  var cat = {};
  leerCatalogo_(ss).forEach(function (it) { cat[it.codigo || (it.descripcion + '|' + it.talla)] = it; });
  var last = ultimaFila_(sh);
  if (last < 2) return crearStockDisponible();
  var rows = leerFilas_(sh, 2, last - 1).map(function (r) {
    var k = String(r[1]).trim() || (String(r[2]).trim() + '|' + txt_(r[3]));
    var it = cat[k] || cat[String(r[1]).trim()];
    var talla = it ? it.talla : txt_(r[3]);
    var costo = it ? it.costo : num_(r[6]);
    var stock = num_(r[5]);
    return [r[0], String(r[1]).trim(), r[2], talla, r[4], stock, costo, stock * costo, r[8]];
  });
  escribir_(sh, 2, rows);
}

function leerPersonal_(ss) {
  var sh = ss.getSheetByName(PERSONAL);
  if (!sh) {
    var hs = ss.getSheets();
    for (var i = 0; i < hs.length; i++) if (hs[i].getName().toLowerCase().indexOf('personal') >= 0) { sh = hs[i]; break; }
  }
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  var hi = v.findIndex(function (r) {
    return r.some(function (c) { return String(c).toUpperCase().indexOf('NOMBRE') >= 0; });
  });
  if (hi < 0) return [];
  var H = v[hi].map(function (c) { return String(c).toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cN = col('PERS'), cD = col('ID NUMBER'), cNom = col('NOMBRE'), cF = col('FUNCI'), cS = col('SEXO'), cA = H.lastIndexOf('AREA');
  if (cA < 0) cA = col('AREA');
  return v.slice(hi + 1).filter(function (r) { return r[cNom]; }).map(function (r) {
    return { codigo: String(r[cN]).trim(), dni: String(r[cD]).trim(), nombre: String(r[cNom]).trim(),
      funcion: String(r[cF]).trim(), sexo: String(r[cS]).trim(), area: String(r[cA]).trim() };
  });
}

function hojaPersonal_(ss) {
  var sh = ss.getSheetByName(PERSONAL);
  if (sh) return sh;
  var hs = ss.getSheets();
  for (var i = 0; i < hs.length; i++) if (hs[i].getName().toLowerCase().indexOf('personal') >= 0) return hs[i];
  return null;
}

// Devuelve encabezados y fila completa del colaborador en "data de personal"
function filaPersonal_(ss, codigo) {
  var sh = hojaPersonal_(ss);
  if (!sh) return { head: [], row: [] };
  var v = sh.getDataRange().getDisplayValues();
  var hi = v.findIndex(function (r) { return r.some(function (c) { return String(c).toUpperCase().indexOf('NOMBRE') >= 0; }); });
  var head = v[hi].map(function (c) { return String(c).trim(); }).filter(function (c) { return c; });
  var cN = v[hi].findIndex(function (c) { return String(c).toUpperCase().indexOf('PERS') >= 0; });
  var row = v.slice(hi + 1).find(function (r) { return String(r[cN]).trim() === String(codigo).trim(); }) || [];
  return { head: head, row: head.map(function (_, i) { return row[i] || ''; }) };
}

var HEAD_ENT_A = ['N° Entrega', 'Fecha', 'Hora'];
var HEAD_ENT_B = ['Motivo entrega', 'TIPO', 'Codigo SAP', 'Descripcion', 'Talla', 'GENERO', 'Cantidad',
  'Devolvio anterior', 'Motivo no devolucion', 'Costo S/.', 'Subtotal S/.', 'Entregado por', 'Firma', 'Registrado'];

function hojaEntregas_(ss, headPers) {
  var head = HEAD_ENT_A.concat(headPers.length ? headPers : ['Nº pers.', 'ID Number', 'Nombre del Colaborador', 'Función', 'Area']).concat(HEAD_ENT_B);
  var sh = hoja_(ss, ENTREGAS);
  if (sh && sh.getLastRow() > 0) {
    var h1 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (String(h1[2]).trim() !== 'Hora') { sh.setName(ENTREGAS + ' (anterior)'); sh = null; }
  }
  if (!sh) sh = ss.insertSheet(ENTREGAS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    try { sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); sh.setFrozenRows(1); } catch (e) { }
  }
  return sh;
}

function registrarEntrega_(ss, d) {
  var pers = filaPersonal_(ss, d.codigo);
  var sh = hojaEntregas_(ss, pers.head);
  var datosPers = pers.row.length && pers.row.some(function (c) { return c; }) ? pers.row : [d.codigo, d.dni, d.colaborador, d.funcion, d.area];
  var firmaUrl = '';
  if (d.firma) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(d.firma.split(',')[1]), 'image/png', 'firma_' + d.id + '.png');
      firmaUrl = DriveApp.createFile(blob).getUrl();
    } catch (e) { firmaUrl = 'error: ' + e.message; }
  }
  var st = hojaStock_(ss), lastT = ultimaFila_(st);
  var data = lastT > 1 ? [[]].concat(leerFilas_(st, 2, lastT - 1)) : [[]], fila = {};
  for (var i = 1; i < data.length; i++) fila[clave_(data[i][1], data[i][2], data[i][3])] = i + 1;
  // Escribe cada valor en la columna cuyo encabezado coincide (no depende del orden ni de columnas extra)
  var K = function (s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, ''); };
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var headPers = pers.head.length ? pers.head : ['Nº pers.', 'ID Number', 'Nombre del Colaborador', 'Función', 'Area'];
  d.items.forEach(function (it) {
    var val = {};
    headPers.forEach(function (k, i) { val[K(k)] = '\'' + (datosPers[i] || ''); });
    var set = function (k, v) { val[K(k)] = v; };
    set('N° Entrega', '\'' + d.id); set('Fecha', d.fecha); set('Hora', d.hora); set('Motivo entrega', d.motivo); set('TIPO', it.tipo);
    set('Codigo SAP', '\'' + String(it.codigo)); set('Descripcion', it.descripcion); set('Talla', '\'' + String(it.talla || 'Generico')); set('GENERO', it.genero || '');
    set('Cantidad', Number(it.cantidad)); set('Devolvio anterior', it.devolvio); set('Motivo no devolucion', it.motivoNo); set('Costo S/.', num_(it.costo));
    set('Subtotal S/.', Number(it.cantidad) * num_(it.costo)); set('Entregado por', d.entregadoPor); set('Firma', firmaUrl); set('Registrado', new Date());
    sh.appendRow(H.map(function (hd) { var k = K(hd); return val.hasOwnProperty(k) ? val[k] : ''; }));
    var r = fila[clave_(it.codigo, it.descripcion, it.talla)];
    if (r && !d.regularizacion) {
      var c = st.getRange(r, 6), nuevo = num_(c.getValue()) - Number(it.cantidad);
      c.setValue(nuevo);
      st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue()));
      st.getRange(r, 9).setValue(new Date());
    }
  });
}

// Borra las filas de un registro (por N°) y ajusta el stock (signo: +1 devuelve, -1 descuenta)
function eliminarRegistro_(ss, hoja, id, items, signo) {
  var sh = hoja_(ss, hoja);
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
    for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]).replace(/^'/, '').trim() === id) sh.deleteRow(i + 1);
  }
  var st = hojaStock_(ss), lastT = ultimaFila_(st);
  var data = lastT > 1 ? [[]].concat(leerFilas_(st, 2, lastT - 1)) : [[]], fila = {};
  for (var j = 1; j < data.length; j++) fila[clave_(data[j][1], data[j][2], data[j][3])] = j + 1;
  (items || []).forEach(function (it) {
    var talla = it.talla === 'Generico' || it.talla === 'Genérico' ? '' : it.talla;
    var r = fila[clave_(it.codigo, it.descripcion, talla)];
    if (r) {
      var c = st.getRange(r, 6), nuevo = num_(c.getValue()) + signo * Number(it.cantidad);
      c.setValue(nuevo);
      st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue()));
      st.getRange(r, 9).setValue(new Date());
    }
  });
}

// Borra todas las filas con ese N° (col. A) leyendo los productos de la propia hoja y ajusta el stock
function eliminarPorId_(ss, hoja, id, signo) {
  var sh = hoja_(ss, hoja);
  if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getDataRange().getDisplayValues();
  var H = v[0].map(function (c) { return String(c).trim().toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cC = col('CODIGO SAP'), cD = col('DESCRIPCION'), cT = H.indexOf('TALLA'), cQ = H.indexOf('CANTIDAD'), cM = col('MOTIVO ENTREGA');
  var limpio = function (s) { return String(s).replace(/^'/, '').trim(); };
  var filas = [], items = [];
  for (var i = 1; i < v.length; i++) {
    if (limpio(v[i][0]) === limpio(id)) {
      filas.push(i + 1);
      if (cM >= 0 && /REGULARIZ/i.test(v[i][cM])) continue;
      var t = limpio(v[i][cT]);
      items.push({ codigo: limpio(v[i][cC]), descripcion: limpio(v[i][cD]), talla: /^gen[eé]rico$/i.test(t) ? '' : t, cantidad: num_(v[i][cQ]) });
    }
  }
  for (var k = filas.length - 1; k >= 0; k--) sh.deleteRow(filas[k]);
  if (!items.length) return 0;
  var st = hojaStock_(ss), lastT = ultimaFila_(st);
  var data = lastT > 1 ? [[]].concat(leerFilas_(st, 2, lastT - 1)) : [[]], fila = {};
  for (var j = 1; j < data.length; j++) fila[clave_(data[j][1], data[j][2], data[j][3])] = j + 1;
  items.forEach(function (it) {
    var r = fila[clave_(it.codigo, it.descripcion, it.talla)];
    if (r) {
      var c = st.getRange(r, 6), nuevo = num_(c.getValue()) + signo * it.cantidad;
      c.setValue(nuevo);
      st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue()));
      st.getRange(r, 9).setValue(new Date());
    }
  });
  return filas.length;
}

// Borra un solo ítem (fila) de un ingreso y descuenta su cantidad del stock
function eliminarItemIngreso_(ss, d) {
  var sh = hoja_(ss, 'Ingresos'); if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getDataRange().getDisplayValues(), H = v[0].map(function (c) { return String(c).trim().toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cC = col('CODIGO SAP'), cD = col('DESCRIPCION'), cT = H.indexOf('TALLA'), cQ = H.indexOf('CANTIDAD');
  var L = function (s) { return String(s).replace(/^'/, '').trim(); }, T = function (s) { s = L(s); return /^gen[eé]rico$/i.test(s) ? '' : s; };
  for (var i = v.length - 1; i >= 1; i--) {
    if (L(v[i][0]) !== L(d.id)) continue;
    var okC = d.codigo ? L(v[i][cC]) === L(d.codigo) : L(v[i][cD]) === L(d.descripcion);
    if (!okC || T(v[i][cT]) !== T(d.talla)) continue;
    var cant = num_(v[i][cQ]), it = { codigo: L(v[i][cC]), descripcion: L(v[i][cD]), talla: T(v[i][cT]) };
    sh.deleteRow(i + 1);
    var st = hojaStock_(ss), lastT = ultimaFila_(st), rows = lastT > 1 ? leerFilas_(st, 2, lastT - 1) : [];
    for (var j = 0; j < rows.length; j++) if (clave_(rows[j][1], rows[j][2], rows[j][3]) === clave_(it.codigo, it.descripcion, it.talla)) {
      var r = j + 2, c = st.getRange(r, 6), nuevo = num_(c.getValue()) - cant;
      c.setValue(nuevo); st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue())); st.getRange(r, 9).setValue(new Date()); break;
    }
    return 1;
  }
  return 0;
}

// Borra un solo ítem (fila) de una entrega y lo devuelve al stock (salvo regularización)
function eliminarItemEntrega_(ss, d) {
  var sh = hoja_(ss, ENTREGAS); if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getDataRange().getDisplayValues(), H = v[0].map(function (c) { return String(c).trim().toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cC = col('CODIGO SAP'), cD = col('DESCRIPCION'), cT = H.indexOf('TALLA'), cQ = H.indexOf('CANTIDAD'), cM = col('MOTIVO ENTREGA');
  var L = function (s) { return String(s).replace(/^'/, '').trim(); }, T = function (s) { s = L(s); return /^gen[eé]rico$/i.test(s) ? '' : s; };
  for (var i = v.length - 1; i >= 1; i--) {
    if (L(v[i][0]) !== L(d.id)) continue;
    var okC = d.codigo ? L(v[i][cC]) === L(d.codigo) : L(v[i][cD]) === L(d.descripcion);
    if (!okC || T(v[i][cT]) !== T(d.talla)) continue;
    var cant = num_(v[i][cQ]), reg = cM >= 0 && /REGULARIZ/i.test(v[i][cM]), it = { codigo: L(v[i][cC]), descripcion: L(v[i][cD]), talla: T(v[i][cT]) };
    sh.deleteRow(i + 1);
    if (reg) return 1;
    var st = hojaStock_(ss), lastT = ultimaFila_(st), rows = lastT > 1 ? leerFilas_(st, 2, lastT - 1) : [];
    for (var j = 0; j < rows.length; j++) if (clave_(rows[j][1], rows[j][2], rows[j][3]) === clave_(it.codigo, it.descripcion, it.talla)) {
      var r = j + 2, c = st.getRange(r, 6), nuevo = num_(c.getValue()) + cant;
      c.setValue(nuevo); st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue())); st.getRange(r, 9).setValue(new Date()); break;
    }
    return 1;
  }
  return 0;
}

// Entregas registradas (para la pestaña "Entregas Registradas")
function leerEntregas_(ss) {
  var sh = hoja_(ss, ENTREGAS);
  if (!sh || sh.getLastRow() < 2) return [];
  var rg = sh.getDataRange(), v = rg.getDisplayValues(), raw = rg.getValues(), H = v[0].map(function (c) { return String(c).trim(); });
  var tz = Session.getScriptTimeZone();
  return v.slice(1).map(function (r, j) {
    var o = {}; H.forEach(function (k, i) { var x = raw[j + 1][i]; o[k] = (x instanceof Date && /^(fecha|registrado)$/i.test(k)) ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(r[i]).replace(/^'/, '').trim(); }); return o;
  }).filter(function (o) { return o[H[0]]; });
}


// ---- Hoja "uniformes x area" (relación área/función → uniformes y EPPs)
var UXA = 'uniformes x area';
function leerUxA_(ss) {
  var sh = ss.getSheetByName(UXA);
  if (!sh) {
    var hs = ss.getSheets();
    for (var i = 0; i < hs.length; i++) if (/x\s*area/i.test(hs[i].getName())) { sh = hs[i]; break; }
  }
  return sh ? sh.getDataRange().getDisplayValues() : null;
}

// ---- TVU (vida útil en días) sugerida para productos sin TVU
var TVU_REGLAS = [[/TAPON/,30],[/GUANTE.*(PANAD|HORNO|NEOPREN|POLLO)/,30],[/GUANTE.*(NITRIL|CONGEL|TEJIDO)/,60],[/GUANTE.*ACERO/,180],[/GUANTE/,120],[/MANGA/,60],[/MASCARILLA|RESPIRADOR|CUBREBOCA|FILTRO/,30],[/BARBIQUEJO/,90],[/BOTA/,90],[/BOTIN|ZAPATO|CALZADO/,150],[/CASCO/,180],[/LENTE|GAFA|ANTEOJO|CARETA|FACIAL/,120],[/OREJERA/,180],[/ARNES|LINEA DE VIDA/,365],[/FAJA/,180],[/CHALECO/,180],[/CHOMPA|POLAR/,365],[/CASACA|CAMISACO|CAMISA|CRISTINA/,180],[/PANTALON/,150],[/POLO PIQUE/,150],[/POLO/,90],[/MANDIL.*(PVC|PU )/,150],[/MANDIL/,180],[/COLLET|GORRO|TOCA|MALLA|COFIA/,120]];
function tvuEstimado_(desc, tipo) {
  var d = String(desc).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  for (var i = 0; i < TVU_REGLAS.length; i++) if (TVU_REGLAS[i][0].test(d)) return TVU_REGLAS[i][1];
  return /EPP/i.test(tipo) ? 90 : 180;
}

// Ejecutar (o usar el botón de la app): completa la columna TVU vacía en "Uniformes/EPPs"
function completarTVU() {
  var sh = hojaCatalogo_(libro_()), v = sh.getDataRange().getValues();
  var hi = v.findIndex(function (r) { return r.some(function (c) { return String(c).trim().toUpperCase() === 'TIPO'; }); });
  var H = v[hi].map(function (c) { return String(c).toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cT = col('TIPO'), cD = col('DESCRIP'), cV = col('TVU'), cTa = col('TALLA');
  if (cV < 0) throw new Error('No existe la columna TVU en la hoja del catálogo.');
  var conocido = {};
  v.slice(hi + 1).forEach(function (r) { var n = num_(r[cV]); if (r[cD] && n > 0) conocido[String(r[cD]).trim()] = n; });
  var hechos = [];
  for (var i = hi + 1; i < v.length; i++) {
    var r = v[i]; if (!r[cD] || num_(r[cV]) > 0) continue;
    var d = String(r[cD]).trim(), t = conocido[d] || tvuEstimado_(d, r[cT]);
    var c = sh.getRange(i + 1, cV + 1); c.setValue(t);
    try { c.setBackground('#fff4c2'); } catch (e) { }
    hechos.push({ descripcion: d, talla: txt_(r[cTa]), tvu: t });
  }
  return hechos;
}


// Alta de producto nuevo desde la app (no estaba en el catálogo)
function agregarProducto_(ss, d) {
  var sh = hojaCatalogo_(ss), v = sh.getDataRange().getValues();
  var hi = v.findIndex(function (r) { return r.some(function (c) { return String(c).trim().toUpperCase() === 'TIPO'; }); });
  var H = v[hi].map(function (c) { return String(c).toUpperCase(); });
  var col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cT = col('TIPO'), cC = col('CODIGO'), cD = col('DESCRIP'), cTa = col('TALLA'), cG = col('GENERO'), cV = col('TVU'), cCo = col('COSTO');
  var cod = String(d.codigo || '').trim();
  if (cod) for (var i = hi + 1; i < v.length; i++) if (String(v[i][cC]).trim() === cod) return { ok: true, existe: true };
  var fila = hi + 1;
  for (var j = v.length - 1; j > hi; j--) if (String(v[j][cD]).trim()) { fila = j + 1; break; }
  fila++;
  var put = function (c, val, fmt) { if (c < 0) return; var r = sh.getRange(fila, c + 1); if (fmt) try { r.setNumberFormat(fmt); } catch (e) { } r.setValue(val); };
  put(cT, d.tipo); put(cC, cod, '@'); put(cD, d.descripcion); put(cTa, String(d.talla || ''), '@');
  put(cG, d.genero); if (num_(d.tvu) > 0) put(cV, num_(d.tvu)); put(cCo, num_(d.costo), '0.00');
  hojaStock_(ss);
  var ax = (d.areas && d.areas.length) ? agregarUxA_(ss, d) : 0;
  return { ok: true, fila: fila, areas: ax };
}

// Agrega una fila por área en la hoja/tabla "uniformes x area"
function agregarUxA_(ss, d) {
  var sh = ss.getSheetByName(UXA);
  if (!sh) { var hs = ss.getSheets(); for (var i = 0; i < hs.length; i++) if (/x\s*area/i.test(hs[i].getName())) { sh = hs[i]; break; } }
  if (!sh) throw new Error('No existe la hoja "' + UXA + '".');
  var v = sh.getDataRange().getDisplayValues();
  var nrm = function (s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); };
  var hi = v.findIndex(function (r) { return r.some(function (c) { return /\barea\b/.test(nrm(c)); }); });
  if (hi < 0) throw new Error('La hoja "' + UXA + '" no tiene columna Área.');
  var H = v[hi].map(nrm), col = function () { var ks = [].slice.call(arguments); return H.findIndex(function (x) { return ks.some(function (k) { return x.indexOf(k) >= 0; }); }); };
  var cA = H.findIndex(function (x) { return /\barea\b/.test(x); }), cC = col('codigo', 'sap'), cD = col('desc', 'uniforme', 'epp', 'item', 'producto', 'prenda'),
      cT = H.findIndex(function (x) { return x === 'tipo'; }), cG = col('genero', 'sexo'), cQ = col('cant'), cTa = col('talla');
  if (cD < 0 && cC < 0) throw new Error('La hoja "' + UXA + '" no tiene columna Descripción ni Código SAP.');
  var ya = {};
  v.slice(hi + 1).forEach(function (r) { ya[nrm(r[cA]) + '|' + nrm(cD >= 0 ? r[cD] : r[cC])] = true; });
  var last = hi + 1;
  for (var j = v.length - 1; j > hi; j--) if (String(v[j][cA]).trim() || (cD >= 0 && String(v[j][cD]).trim())) { last = j + 1; break; }
  var n = 0;
  d.areas.forEach(function (a) {
    if (ya[nrm(a) + '|' + nrm(cD >= 0 ? d.descripcion : d.codigo)]) return;
    last++; n++;
    var put = function (c, val, fmt) { if (c < 0) return; var r = sh.getRange(last, c + 1); if (fmt) try { r.setNumberFormat(fmt); } catch (e) { } r.setValue(val); };
    put(cA, a); put(cC, String(d.codigo || ''), '@'); put(cD, d.descripcion); put(cT, d.tipo); put(cG, d.genero); put(cQ, 1); put(cTa, String(d.talla || ''), '@');
  });
  return n;
}


// Ajusta el stock de un producto a una cantidad exacta (desde la pestaña Stock disponible)
function ajustarStock_(ss, d) {
  var st = hojaStock_(ss), last = ultimaFila_(st);
  var rows = last > 1 ? leerFilas_(st, 2, last - 1) : [], k = clave_(d.codigo, d.descripcion, d.talla);
  for (var i = 0; i < rows.length; i++) if (clave_(rows[i][1], rows[i][2], rows[i][3]) === k) {
    var r = i + 2, nuevo = num_(d.stock);
    st.getRange(r, 6).setValue(nuevo);
    st.getRange(r, 8).setValue(nuevo * num_(st.getRange(r, 7).getValue()));
    st.getRange(r, 9).setValue(new Date());
    return { ok: true, fila: r };
  }
  return { ok: false, error: 'No se encontró el producto en "Stock disponible".' };
}

// Elimina un producto del catálogo, de "Stock disponible" y (si no quedan otras tallas) de "uniformes x area"
function eliminarProducto_(ss, d) {
  var T = function (s) { return String(s === null || s === undefined ? '' : s).trim(); };
  var nrm = function (s) { return T(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); };
  var cod = T(d.codigo), desc = T(d.descripcion), talla = T(d.talla);
  var es = function (c, de, ta) { return cod ? T(c) === cod : (T(de) === desc && txt_(ta) === talla); };
  var n = { catalogo: 0, stock: 0, uxa: 0 };
  var sh = hojaCatalogo_(ss), v = sh.getDataRange().getValues();
  var hi = v.findIndex(function (r) { return r.some(function (c) { return T(c).toUpperCase() === 'TIPO'; }); });
  var H = v[hi].map(function (c) { return T(c).toUpperCase(); }), col = function (t) { return H.findIndex(function (x) { return x.indexOf(t) >= 0; }); };
  var cC = col('CODIGO'), cD = col('DESCRIP'), cTa = col('TALLA'), quedan = false;
  for (var i = v.length - 1; i > hi; i--) {
    if (es(v[i][cC], v[i][cD], v[i][cTa])) { sh.deleteRow(i + 1); n.catalogo++; }
    else if (T(v[i][cD]) === desc) quedan = true;
  }
  var st = hoja_(ss, STOCK);
  if (st) { var last = ultimaFila_(st); if (last > 1) { var rows = leerFilas_(st, 2, last - 1);
    for (var j = rows.length - 1; j >= 0; j--) if (es(rows[j][1], rows[j][2], rows[j][3])) { st.deleteRow(j + 2); n.stock++; } } }
  if (!quedan) {
    var ux = hoja_(ss, UXA);
    if (ux && ux.getLastRow() > 1) {
      var w = ux.getDataRange().getDisplayValues();
      var h2 = w.findIndex(function (r) { return r.some(function (c) { return /\barea\b/.test(nrm(c)); }); });
      if (h2 >= 0) {
        var HH = w[h2].map(nrm), f = function () { var ks = [].slice.call(arguments); return HH.findIndex(function (x) { return ks.some(function (k) { return x.indexOf(k) >= 0; }); }); };
        var uC = f('codigo', 'sap'), uD = f('desc', 'uniforme', 'epp', 'item', 'producto', 'prenda');
        for (var q = w.length - 1; q > h2; q--) {
          var hitC = cod && uC >= 0 && T(w[q][uC]) === cod, hitD = uD >= 0 && nrm(w[q][uD]) === nrm(desc);
          if (hitC || hitD) { ux.deleteRow(q + 1); n.uxa++; }
        }
      }
    }
  }
  return { ok: true, borradas: n };
}


// Historial de ingresos (para que todos los equipos vean lo mismo)
function leerIngresos_(ss) {
  var sh = hoja_(ss, 'Ingresos'); if (!sh || sh.getLastRow() < 2) return [];
  var rg = sh.getDataRange(), v = rg.getDisplayValues(), raw = rg.getValues(), H = v[0].map(function (c) { return String(c).trim(); }), tz = Session.getScriptTimeZone();
  return v.slice(1).map(function (r, j) {
    var o = {}; H.forEach(function (k, i) { var x = raw[j + 1][i]; o[k] = (x instanceof Date && /^fecha$/i.test(k)) ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(r[i]).replace(/^'/, '').trim(); }); return o;
  }).filter(function (o) { return o[H[0]]; });
}



// ======================================================================
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
