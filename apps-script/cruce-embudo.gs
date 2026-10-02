/**
 * CRUCE DE EMBUDO — ATLAB (Lya + Clientes) — v5
 * ------------------------------------------------------------
 * CAMBIOS RESPECTO A v4
 * 1) Teléfonos en notación científica ("5,7324E+11") o con menos de 10
 *    dígitos se descartan del cruce (quedan vacíos). Antes se convertían en
 *    "57324" y cientos de contactos distintos terminaban con el mismo
 *    "teléfono" — por eso Lya mostraba muy pocos contactos. El script NO
 *    puede recuperar esos dígitos: hay que volver a pegar la columna desde
 *    la fuente con la columna en formato "Texto sin formato".
 * 2) Fechas con año de 2 dígitos ("30/07/26 0:00") ahora se leen bien. Antes
 *    caían a new Date() y se leían en formato gringo (mes/día) o quedaban
 *    vacías.
 * 3) Lya ya se sube en hora Bogotá (LYA_FECHAS_EN_UTC: false). La v4 le
 *    restaba 5 horas en cada corrida. Si se activa el modo UTC, el ajuste se
 *    aplica una sola vez por fila y se marca en la columna "_bogota".
 * 4) Clientes: también se normaliza DTINICIOCONTRATO (inicio de contrato).
 * 5) Los meses/fechas de las tablas resumen se escriben como texto
 *    ("2026-09"), para que Sheets no los convierta en "2026-9".
 * 6) Al final muestra cuántos teléfonos inválidos encontró en cada hoja.
 *
 * QUÉ HACE
 * 1) Normaliza en el sitio Clientes y Lya: teléfono → texto de 10 dígitos,
 *    correo → minúsculas, fechas → Date reales en hora Bogotá.
 * 2) Maestro_Cruce: una fila por teléfono con FECHA_LYA, FECHA_CLIENTE y
 *    DIAS_A_CONVERSION.
 * 3) Resumen_General: diario (Fecha, Contactos_Nuevos_Lya, Clientes_Nuevos).
 * 4) Resumen_Mensual: volumen mes a mes.
 * 5) Cohorte_Mensual: por mes de entrada a Lya, % convertido a hoy y días.
 *
 * NO TOCA la pestaña "Ventas mes" (manual) ni FB_Atlab.
 * El tablero lee directo Lya, FB_Atlab y Ventas mes; Maestro_Cruce solo se usa
 * en la vista Detalle (búsqueda por teléfono o nombre).
 *
 * SUPUESTO: fechas de EVO (Clientes) y de Lya ya en hora Bogotá. Si algún día
 * se pega la exportación cruda de Lya (en UTC), poner LYA_FECHAS_EN_UTC: true.
 *
 * CÓMO USARLO
 * 1) Extensiones > Apps Script, pega este código (reemplaza el anterior).
 * 2) Ejecuta procesarTodo().
 * 3) Corre procesarTodo() cada vez que alimentes Clientes o Lya, o programa
 *    crearTriggerDiario() para que corra solo.
 * ------------------------------------------------------------
 */

var CONFIG = {
  SHEET_CLIENTES: 'Clientes',
  SHEET_LYA: 'Lya',
  SHEET_MAESTRO: 'Maestro_Cruce',
  SHEET_RESUMEN: 'Resumen_General',
  SHEET_RESUMEN_MENSUAL: 'Resumen_Mensual',
  SHEET_COHORTE_MENSUAL: 'Cohorte_Mensual',
  ZONA_HORARIA: 'America/Bogota',
  COL_MARCA_UTC: '_bogota', // marca en Lya: fila ya convertida de UTC a Bogotá
  // false = la base de Lya que se sube ya viene en hora Bogotá (como el histórico
  // limpio). Poner true solo si se pega la exportación cruda de Lya, que viene en UTC.
  LYA_FECHAS_EN_UTC: false,
};

function procesarTodo() {
  var reporte = normalizarFuentes_();
  construirCruce_();
  var msg = 'Cruce actualizado.';
  if (reporte.length) msg += ' Teléfonos inválidos (fuera del cruce): ' + reporte.join(' · ');
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'ATLAB', 10);
}

// =================================================================
// PASO 1 — NORMALIZACIÓN DE LAS HOJAS FUENTE (en el sitio)
// =================================================================

function normalizarFuentes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reporte = [];

  var invCli = normalizarHoja_(ss, CONFIG.SHEET_CLIENTES, {
    telefono: ['TELÉFONO', 'TELEFONO'],
    correo: ['CORREO ELECTRÓNICO', 'CORREO ELECTRONICO'],
    fechasLocales: ['FECHA REGISTRO', 'FECHA CONVERSIÓN', 'DATACONVERSAO', 'DTINICIOCONTRATO',
                    'FECHA VENCIMIENTO CONTRATO', 'DATAVENCIMENTOCONTRATO'],
  });
  if (invCli) reporte.push('Clientes ' + invCli);

  var specLya = { telefono: ['telefono'] };
  if (CONFIG.LYA_FECHAS_EN_UTC) specLya.fechasUTC = ['primera_vez', 'ultima_vez'];
  else specLya.fechasLocales = ['primera_vez', 'ultima_vez'];
  var invLya = normalizarHoja_(ss, CONFIG.SHEET_LYA, specLya);
  if (invLya) reporte.push('Lya ' + invLya);

  return reporte;
}

/** Devuelve cuántas filas tienen un teléfono que no se pudo normalizar. */
function normalizarHoja_(ss, nombreHoja, spec) {
  var sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) throw new Error('No encuentro la pestaña "' + nombreHoja + '".');

  // Columna de marca para no convertir dos veces UTC → Bogotá
  if (spec.fechasUTC) {
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    if (buscarColumna_(hdr, [CONFIG.COL_MARCA_UTC]) < 0) {
      var ultima = 0;
      hdr.forEach(function (h, i) { if (h) ultima = i + 1; });
      sheet.getRange(1, ultima + 1).setValue(CONFIG.COL_MARCA_UTC);
    }
  }

  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length < 2) return 0;

  var headers = values[0].map(function (h) { return String(h).trim(); });

  var colTel = spec.telefono ? buscarColumna_(headers, spec.telefono) : -1;
  var colCorreo = spec.correo ? buscarColumna_(headers, spec.correo) : -1;
  var colsFechaLocal = (spec.fechasLocales || []).map(function (n) { return buscarColumna_(headers, [n]); }).filter(function (i) { return i >= 0; });
  var colsFechaUTC = (spec.fechasUTC || []).map(function (n) { return buscarColumna_(headers, [n]); }).filter(function (i) { return i >= 0; });
  var colMarca = spec.fechasUTC ? buscarColumna_(headers, [CONFIG.COL_MARCA_UTC]) : -1;

  var invalidos = 0;
  for (var i = 1; i < values.length; i++) {
    if (colTel >= 0) {
      var original = values[i][colTel];
      var tel = normalizarTelefono_(original);
      if (!tel && original !== '' && original !== null) invalidos++;
      // Si el teléfono es inválido se deja el valor original visible para poder corregirlo,
      // pero leerHoja_/construirCruce_ lo vuelven a validar y lo dejan fuera del cruce.
      values[i][colTel] = tel || original;
    }
    if (colCorreo >= 0) values[i][colCorreo] = normalizarCorreo_(values[i][colCorreo]);
    colsFechaLocal.forEach(function (c) { values[i][c] = parsearFechaCruda_(values[i][c], false); });

    if (colsFechaUTC.length) {
      var yaConvertida = colMarca >= 0 && String(values[i][colMarca]).trim() === 'ok';
      var filaVacia = colsFechaUTC.every(function (c) { return values[i][c] === '' || values[i][c] === null; });
      colsFechaUTC.forEach(function (c) { values[i][c] = parsearFechaCruda_(values[i][c], !yaConvertida); });
      if (colMarca >= 0 && !filaVacia) values[i][colMarca] = 'ok';
    }
  }

  // Teléfono como texto ANTES de escribir, para que Sheets no lo vuelva número
  var numFilas = values.length - 1;
  if (colTel >= 0) sheet.getRange(2, colTel + 1, numFilas, 1).setNumberFormat('@');
  if (colCorreo >= 0) sheet.getRange(2, colCorreo + 1, numFilas, 1).setNumberFormat('@');

  range.setValues(values);

  colsFechaLocal.concat(colsFechaUTC).forEach(function (c) {
    sheet.getRange(2, c + 1, numFilas, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  });
  return invalidos;
}

function buscarColumna_(headers, candidatos) {
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < candidatos.length; j++) {
      if (headers[i].toUpperCase() === candidatos[j].toUpperCase()) return i;
    }
  }
  return -1;
}

/**
 * Devuelve el teléfono en 10 dígitos, o '' si no es válido.
 * - Notación científica ("5,7324E+11") → '' (los dígitos ya se perdieron).
 * - Menos de 10 dígitos → '' (evita que números truncados choquen entre sí).
 * - Con indicativo (57...) o espacios → últimos 10 dígitos.
 */
function normalizarTelefono_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  var s;
  if (typeof valor === 'number') {
    s = valor.toFixed(0); // número real completo, sin notación científica
  } else {
    s = String(valor).trim();
    if (/e\+?\d+$/i.test(s)) return '';
  }
  var digitos = s.replace(/\D/g, '');
  if (digitos.length < 10) return '';
  return digitos.slice(-10);
}

function normalizarCorreo_(valor) {
  if (!valor) return '';
  return String(valor).trim().toLowerCase();
}

/**
 * Convierte a Date. Si aplicarUTC es true, resta 5 horas (UTC → Bogotá).
 * Acepta Date, dd/mm/yyyy[ hh:mm], dd/mm/yy[ hh:mm] y yyyy-mm-dd hh:mm.
 */
function parsearFechaCruda_(valor, aplicarUTC) {
  if (!valor && valor !== 0) return '';
  var fecha = null;

  if (valor instanceof Date) {
    fecha = valor;
  } else {
    var s = String(valor).trim();
    if (!s) return '';

    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})(?!\d)(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
      var d = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      var hh = m[4] ? parseInt(m[4], 10) : 0, mm = m[5] ? parseInt(m[5], 10) : 0;
      fecha = new Date(y, mo, d, hh, mm);
    } else {
      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[\sT]+(\d{1,2}):(\d{2}))?/);
      if (m) {
        fecha = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10),
                         m[4] ? parseInt(m[4], 10) : 0, m[5] ? parseInt(m[5], 10) : 0);
      }
      // Cualquier otro formato se deja vacío en vez de adivinar (new Date() lo lee mes/día).
    }
  }

  if (!fecha || isNaN(fecha.getTime())) return '';
  if (aplicarUTC) fecha = new Date(fecha.getTime() - 5 * 60 * 60 * 1000);
  return fecha;
}

// =================================================================
// PASO 2 — CRUCE (Maestro_Cruce) Y RESÚMENES
// =================================================================

function construirCruce_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var clientes = leerHoja_(ss, CONFIG.SHEET_CLIENTES);
  var lya = leerHoja_(ss, CONFIG.SHEET_LYA);

  var porTelefono = {};
  function asegurar(tel) {
    if (!porTelefono[tel]) {
      porTelefono[tel] = {
        TEL_NORM: tel, NOMBRE: '', CORREO: '', PLAN: '',
        FECHA_LYA: '', FECHA_CLIENTE: '', DIAS_A_CONVERSION: '',
      };
    }
    return porTelefono[tel];
  }

  lya.forEach(function (r) {
    var tel = normalizarTelefono_(r['telefono']);
    if (!tel) return;
    var e = asegurar(tel);
    e.NOMBRE = e.NOMBRE || r['nombre'] || '';
    var f = r['primera_vez'];
    if (f instanceof Date && (!(e.FECHA_LYA instanceof Date) || f < e.FECHA_LYA)) e.FECHA_LYA = f;
  });

  clientes.forEach(function (r) {
    var tel = normalizarTelefono_(r['TELÉFONO'] || r['TELEFONO']);
    if (!tel) return;
    var e = asegurar(tel);
    e.NOMBRE = r['NOMBRE'] || e.NOMBRE || '';
    e.CORREO = r['CORREO ELECTRÓNICO'] || e.CORREO;
    e.PLAN = r['PLAN'] || e.PLAN;
    var fc = r['FECHA CONVERSIÓN'] || r['DATACONVERSAO'];
    if (fc instanceof Date && (!(e.FECHA_CLIENTE instanceof Date) || fc < e.FECHA_CLIENTE)) e.FECHA_CLIENTE = fc;
  });

  // DIAS_A_CONVERSION: solo si hay ambas fechas y el cliente es posterior al
  // contacto Lya (si es anterior, ya era cliente antes de escribir por WhatsApp).
  Object.keys(porTelefono).forEach(function (k) {
    var e = porTelefono[k];
    if (e.FECHA_LYA instanceof Date && e.FECHA_CLIENTE instanceof Date && e.FECHA_CLIENTE >= e.FECHA_LYA) {
      var ms = e.FECHA_CLIENTE.getTime() - e.FECHA_LYA.getTime();
      e.DIAS_A_CONVERSION = Math.round(ms / (1000 * 60 * 60 * 24));
    }
  });

  var maestro = Object.keys(porTelefono).map(function (k) { return porTelefono[k]; });
  maestro.sort(function (a, b) { return a.TEL_NORM < b.TEL_NORM ? -1 : 1; });

  escribirMaestro_(ss, maestro);
  escribirResumenDiario_(ss, maestro);
  escribirResumenMensual_(ss, maestro);
  escribirCohorteMensual_(ss, maestro);
}

function leerHoja_(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) throw new Error('No encuentro la pestaña "' + nombre + '".');
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) if (headers[j]) row[headers[j]] = values[i][j];
    rows.push(row);
  }
  return rows;
}

// ---- escritura Maestro_Cruce ----

var COLS_MAESTRO = ['TEL_NORM', 'NOMBRE', 'CORREO', 'PLAN', 'FECHA_LYA', 'FECHA_CLIENTE', 'DIAS_A_CONVERSION'];

function escribirMaestro_(ss, maestro) {
  var sheet = obtenerOCrearHoja_(ss, CONFIG.SHEET_MAESTRO);
  var filas = maestro.map(function (e) { return COLS_MAESTRO.map(function (c) { return e[c] === undefined ? '' : e[c]; }); });
  sheet.getRange(1, 1, 1, COLS_MAESTRO.length).setValues([COLS_MAESTRO]);
  var colTel = COLS_MAESTRO.indexOf('TEL_NORM') + 1;
  if (filas.length) {
    sheet.getRange(2, colTel, filas.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, filas.length, COLS_MAESTRO.length).setValues(filas);
  }
  sheet.getRange(1, 1, 1, COLS_MAESTRO.length).setFontWeight('bold').setBackground('#1F1F1F').setFontColor('#FFC000');
  sheet.setFrozenRows(1);
  ['FECHA_LYA', 'FECHA_CLIENTE'].forEach(function (col) {
    var idx = COLS_MAESTRO.indexOf(col) + 1;
    if (filas.length) sheet.getRange(2, idx, filas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  });
  var idxDias = COLS_MAESTRO.indexOf('DIAS_A_CONVERSION') + 1;
  if (filas.length) sheet.getRange(2, idxDias, filas.length, 1).setNumberFormat('0');
  sheet.autoResizeColumns(1, COLS_MAESTRO.length);
}

// ---- Resumen_General: diario, tabular puro ----

function escribirResumenDiario_(ss, maestro) {
  var tz = CONFIG.ZONA_HORARIA;
  function clave(fecha) { return fecha instanceof Date ? Utilities.formatDate(fecha, tz, 'yyyy-MM-dd') : null; }

  var diaLya = {}, diaCliente = {};
  maestro.forEach(function (e) {
    var kL = clave(e.FECHA_LYA); if (kL) diaLya[kL] = (diaLya[kL] || 0) + 1;
    var kC = clave(e.FECHA_CLIENTE); if (kC) diaCliente[kC] = (diaCliente[kC] || 0) + 1;
  });

  var todasLasFechas = {};
  [diaLya, diaCliente].forEach(function (obj) { Object.keys(obj).forEach(function (k) { todasLasFechas[k] = true; }); });
  var fechasOrdenadas = Object.keys(todasLasFechas).sort();

  var filas = [['Fecha', 'Contactos_Nuevos_Lya', 'Clientes_Nuevos']];
  fechasOrdenadas.forEach(function (f) {
    filas.push([f, diaLya[f] || 0, diaCliente[f] || 0]);
  });

  var sheet = escribirTabla_(ss, CONFIG.SHEET_RESUMEN, filas);
  if (filas.length > 1) sheet.getRange(2, 2, filas.length - 1, 2).setNumberFormat('0');
}

// ---- Resumen_Mensual: volumen simple mes a mes ----
// Pct_Conversion_Mensual = clientes del mes ÷ contactos Lya del mes. Son poblaciones
// distintas (puede pasar de 100%); el tablero no lo usa. Se deja para no mover columnas.

function escribirResumenMensual_(ss, maestro) {
  var tz = CONFIG.ZONA_HORARIA;
  function claveMes(fecha) { return fecha instanceof Date ? Utilities.formatDate(fecha, tz, 'yyyy-MM') : null; }

  var mesLya = {}, mesCliente = {}, mesDiasSuma = {}, mesDiasN = {};
  maestro.forEach(function (e) {
    var mL = claveMes(e.FECHA_LYA); if (mL) mesLya[mL] = (mesLya[mL] || 0) + 1;
    var mC = claveMes(e.FECHA_CLIENTE);
    if (mC) {
      mesCliente[mC] = (mesCliente[mC] || 0) + 1;
      if (typeof e.DIAS_A_CONVERSION === 'number') {
        mesDiasSuma[mC] = (mesDiasSuma[mC] || 0) + e.DIAS_A_CONVERSION;
        mesDiasN[mC] = (mesDiasN[mC] || 0) + 1;
      }
    }
  });

  var todosLosMeses = {};
  [mesLya, mesCliente].forEach(function (obj) { Object.keys(obj).forEach(function (k) { todosLosMeses[k] = true; }); });
  var mesesOrdenados = Object.keys(todosLosMeses).sort();

  var filas = [['Mes', 'Contactos_Nuevos_Lya', 'Clientes_Nuevos', 'Pct_Conversion_Mensual', 'Dias_Promedio_Conversion']];
  mesesOrdenados.forEach(function (m) {
    var lya = mesLya[m] || 0, cli = mesCliente[m] || 0;
    var pct = lya ? cli / lya : 0;
    var diasProm = mesDiasN[m] ? Math.round((mesDiasSuma[m] / mesDiasN[m]) * 10) / 10 : '';
    filas.push([m, lya, cli, pct, diasProm]);
  });

  var sheet = escribirTabla_(ss, CONFIG.SHEET_RESUMEN_MENSUAL, filas);
  if (filas.length > 1) {
    sheet.getRange(2, 2, filas.length - 1, 2).setNumberFormat('0');
    sheet.getRange(2, 4, filas.length - 1, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 5, filas.length - 1, 1).setNumberFormat('0.0');
  }
}

// ---- Cohorte_Mensual: cohorte real por mes de entrada (Lya) ----
// De los contactos que llegaron por Lya en el mes X, % que ya convirtió a HOY y
// cuántos días tardó. Los meses recientes muestran % bajo: aún no ha pasado tiempo.

function escribirCohorteMensual_(ss, maestro) {
  var tz = CONFIG.ZONA_HORARIA;
  function claveMes(fecha) { return fecha instanceof Date ? Utilities.formatDate(fecha, tz, 'yyyy-MM') : null; }

  var cohortes = {};
  maestro.forEach(function (e) {
    var m = claveMes(e.FECHA_LYA);
    if (!m) return;
    if (!cohortes[m]) cohortes[m] = { contactos: 0, convertidos: 0, dias: [] };
    cohortes[m].contactos++;
    if (e.FECHA_CLIENTE instanceof Date) {
      cohortes[m].convertidos++;
      if (typeof e.DIAS_A_CONVERSION === 'number') cohortes[m].dias.push(e.DIAS_A_CONVERSION);
    }
  });

  var meses = Object.keys(cohortes).sort();
  var filas = [['Mes_Cohorte_Lya', 'Contactos_Del_Mes', 'Convertidos_A_Hoy', 'Pct_Convertido_A_Hoy', 'Dias_Promedio', 'Dias_Mediana']];
  meses.forEach(function (m) {
    var c = cohortes[m];
    var dias = c.dias.slice().sort(function (a, b) { return a - b; });
    var prom = dias.length ? Math.round((dias.reduce(function (a, b) { return a + b; }, 0) / dias.length) * 10) / 10 : '';
    var mediana = dias.length ? (dias.length % 2 ? dias[(dias.length - 1) / 2] : Math.round(((dias[dias.length / 2 - 1] + dias[dias.length / 2]) / 2) * 10) / 10) : '';
    filas.push([m, c.contactos, c.convertidos, c.contactos ? c.convertidos / c.contactos : 0, prom, mediana]);
  });

  var sheet = escribirTabla_(ss, CONFIG.SHEET_COHORTE_MENSUAL, filas);
  if (filas.length > 1) {
    sheet.getRange(2, 2, filas.length - 1, 2).setNumberFormat('0');
    sheet.getRange(2, 4, filas.length - 1, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 5, filas.length - 1, 2).setNumberFormat('0.0');
  }
}

/**
 * Escribe una tabla pura: fila 1 = encabezados, resto = datos.
 * La primera columna (fecha/mes) se fuerza a texto ANTES de escribir, para que
 * Sheets no convierta "2026-09" en fecha ni lo muestre como "2026-9".
 */
function escribirTabla_(ss, nombreHoja, filas) {
  var sheet = obtenerOCrearHoja_(ss, nombreHoja);
  if (filas.length > 1) sheet.getRange(2, 1, filas.length - 1, 1).setNumberFormat('@');
  sheet.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  sheet.getRange(1, 1, 1, filas[0].length).setFontWeight('bold').setBackground('#1F1F1F').setFontColor('#FFC000');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, filas[0].length);
  return sheet;
}

/** Solo se usa con las pestañas que genera el script (nunca con Clientes, Lya, FB_Atlab ni Ventas mes). */
function obtenerOCrearHoja_(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) sheet = ss.insertSheet(nombre);
  sheet.clear();
  return sheet;
}

function crearTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'procesarTodo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarTodo').timeBased().everyDays(1).atHour(6).create();
}
