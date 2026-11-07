/**
 * VALHALLA · API sobre Google Sheets
 * -----------------------------------
 * Pegá este código en Extensiones > Apps Script dentro de la planilla de Valhalla
 * y desplegalo como "Aplicación web" (ver README).
 *
 * La app web habla siempre por POST {action, usuario, pin, payload}: 'login', 'all' o una acción de escritura.
 * Los operadores no reciben las hojas de stock ni de contactos.
 * Solo escribe en las columnas de carga manual; las columnas calculadas del Sheet
 * siguen funcionando con sus fórmulas.
 */

// Los usuarios, PIN y roles se cargan en la hoja "Usuarios".
// Acciones que solo puede hacer un admin:
const ADMIN_ACTIONS = ['addProducto', 'updateProducto', 'addContacto', 'updateContacto', 'importContactos'];

const FIRST_ROW = 5; // los datos empiezan en la fila 5 (fila 4 = encabezados)

// Columnas de cada hoja, en orden. Las que empiezan con "_" son calculadas: se leen pero nunca se escriben.
const SCHEMA = {
  Servicios:      ['id', 'nombre', 'categoria', 'dias', 'precio', 'desc'],
  Marcas_Modelos: ['marca', 'modelo', '_vehiculo', 'segmento', '_ftSeg', 'ajuste', '_ft', '_fp'],
  Clientes:       ['id', 'nombre', 'telefono', 'email', 'direccion', 'alta', 'notas'],
  Vehiculos:      ['patente', 'idCliente', '_cliente', 'vehiculo', 'anio', 'color'],
  Presupuestos:   ['n', 'fecha', 'patente', '_cliente', '_vehiculo', 'servicio', '_db', '_ft', '_de', '_pb', '_fp', '_pe', 'estado'],
  Turnos:         ['id', 'n', '_patente', '_cliente', '_vehiculo', 'ingreso', '_dias', '_entrega', 'estado', 'notas'],
  Trabajos:       ['id', 'n', 'inicio', 'fin', 'patente', '_idCliente', '_cliente', '_vehiculo', 'servicio', '_de', '_dr', '_dv', 'precio', 'medio', 'notas'],
  Productos:      ['codigo', 'nombre', 'categoria', 'unidad', 'inicial', 'minimo', 'costo', 'proveedor', '_entradas', '_salidas', '_stock', '_estado', '_valor'],
  Movimientos:    ['fecha', 'codigo', '_producto', 'tipo', 'cantidad', 'motivo', 'usuario'],
  Contactos:      ['id', 'nombre', 'tipo', 'telefono', 'email', 'empresa', 'notas'],
  Solicitudes:    ['id', 'fecha', 'nombre', 'telefono', 'vehiculo', 'patente', 'servicios', 'deseada', 'dias', 'precio', 'notas', 'estado', 'n']
};
// Acciones del link público de reservas: no piden usuario ni PIN
const PUBLIC_ACTIONS = ['catalogo', 'solicitar'];
const ADMIN_SHEETS = ['Productos', 'Movimientos', 'Contactos'];
// Las solicitudes que llegan del link de reservas quedan en este estado hasta que el negocio las confirme
const ESPERA = 'En espera de aprobación';

/* ---------------- entrada ---------------- */
function doGet() {
  return json({ ok: true, info: 'API Valhalla activa. La app usa POST.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const body = JSON.parse(e.postData.contents);
    if (PUBLIC_ACTIONS.indexOf(body.action) >= 0) {
      const res = body.action === 'catalogo' ? { data: catalogo() } : { result: solicitar(body.payload || {}) };
      SpreadsheetApp.flush();
      return json(Object.assign({ ok: true }, res));
    }
    const user = auth(body.usuario, body.pin);
    let result = {};
    if (body.action !== 'all' && body.action !== 'login') {
      const fn = ACTIONS[body.action];
      if (!fn) throw new Error('Acción desconocida: ' + body.action);
      if (ADMIN_ACTIONS.indexOf(body.action) >= 0 && user.rol !== 'admin') throw new Error('Solo un admin puede hacer esto');
      result = fn(body.payload || {}, user);
      SpreadsheetApp.flush();
    }
    return json({ ok: true, user: user, result: result, data: readAll(user.rol) });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- acciones ---------------- */
const ACTIONS = {
  addCliente: function (p) {
    const id = nextId('Clientes', 'id', 'C', 3);
    insertRow('Clientes', { id: id, nombre: p.nombre, telefono: p.telefono, email: p.email, direccion: p.direccion, alta: new Date(), notas: p.notas });
    return { id: id };
  },
  addVehiculo: function (p) {
    if (findRows('Vehiculos', 'patente', p.patente).length) throw new Error('La patente ya existe');
    insertRow('Vehiculos', { patente: p.patente, idCliente: p.idCliente, vehiculo: p.vehiculo, anio: p.anio, color: p.color });
    return { patente: p.patente };
  },
  addPresupuesto: function (p) {
    const n = nextId('Presupuestos', 'n', 'P', 4);
    const estado = p.estado || 'Pendiente';
    (p.servicios || []).forEach(function (s) {
      insertRow('Presupuestos', { n: n, fecha: new Date(), patente: p.patente, servicio: s, estado: estado });
    });
    return { n: n, estado: estado };
  },
  setPresupuestoEstado: function (p) {
    setField('Presupuestos', 'n', p.n, 'estado', p.estado);
    return {};
  },
  addTurno: function (p) {
    const id = nextId('Turnos', 'id', 'T', 3);
    insertRow('Turnos', { id: id, n: p.n, ingreso: toDate(p.ingreso), estado: p.estado || 'Agendado', notas: p.notas || '' });
    setField('Presupuestos', 'n', p.n, 'estado', 'Aprobado');
    return { id: id };
  },
  setTurnoEstado: function (p) {
    setField('Turnos', 'id', p.id, 'estado', p.estado);
    return {};
  },
  addProducto: function (p) {
    const codigo = nextId('Productos', 'codigo', 'PR', 2);
    insertRow('Productos', { codigo: codigo, nombre: p.nombre, categoria: p.categoria, unidad: p.unidad, inicial: Number(p.inicial) || 0,
      minimo: Number(p.minimo) || 0, costo: Number(p.costo) || 0, proveedor: p.proveedor || '' });
    return { codigo: codigo };
  },
  updateProducto: function (p) {
    ['nombre', 'categoria', 'unidad', 'minimo', 'costo', 'proveedor'].forEach(function (k) {
      if (p.hasOwnProperty(k)) setField('Productos', 'codigo', p.codigo, k, p[k]);
    });
    return {};
  },
  addMovimiento: function (p, user) {
    if (p.tipo !== 'Entrada' && p.tipo !== 'Salida') throw new Error('Tipo inválido');
    if (p.tipo === 'Entrada' && user.rol !== 'admin') throw new Error('Solo un admin carga entradas de stock');
    insertRow('Movimientos', { fecha: toDate(p.fecha) || new Date(), codigo: p.codigo, tipo: p.tipo, cantidad: Number(p.cantidad) || 0,
      motivo: p.motivo || '', usuario: user.usuario });
    return {};
  },
  addContacto: function (p) {
    const id = nextId('Contactos', 'id', 'K', 3);
    insertRow('Contactos', { id: id, nombre: p.nombre, tipo: p.tipo, telefono: p.telefono, email: p.email, empresa: p.empresa, notas: p.notas });
    return { id: id };
  },
  updateContacto: function (p) {
    ['nombre', 'tipo', 'telefono', 'email', 'empresa', 'notas'].forEach(function (k) {
      if (p.hasOwnProperty(k)) setField('Contactos', 'id', p.id, k, p[k]);
    });
    return {};
  },
  importContactos: function (p) {
    let n = 0;
    (p.items || []).forEach(function (c) {
      if (!c.nombre) return;
      ACTIONS.addContacto(c); n++;
    });
    return { importados: n };
  },
  confirmarSolicitud: function (p) {
    const rows = findRows('Solicitudes', 'id', p.id);
    if (!rows.length) throw new Error('Solicitud no encontrada');
    const sh = SpreadsheetApp.getActive().getSheetByName('Solicitudes');
    const v = sh.getRange(rows[0], 1, 1, SCHEMA.Solicitudes.length).getValues()[0];
    const s = {}; SCHEMA.Solicitudes.forEach(function (k, i) { s[k] = v[i]; });
    if (s.estado !== ESPERA) throw new Error('La solicitud ya fue ' + String(s.estado).toLowerCase());
    // cliente: se busca por teléfono (últimos 8 dígitos); si no existe, se crea
    const tel = String(s.telefono).replace(/\D/g, '').slice(-8);
    const cs = SpreadsheetApp.getActive().getSheetByName('Clientes');
    const nc = Math.max(0, cs.getLastRow() - FIRST_ROW + 1);
    let idCliente = '';
    if (nc) cs.getRange(FIRST_ROW, 1, nc, 3).getValues().forEach(function (r) {
      if (!idCliente && r[0] && tel && String(r[2]).replace(/\D/g, '').slice(-8) === tel) idCliente = r[0];
    });
    if (!idCliente) idCliente = ACTIONS.addCliente({ nombre: s.nombre, telefono: s.telefono, email: '', direccion: '', notas: 'Llegó por el link de reservas' }).id;
    if (!findRows('Vehiculos', 'patente', s.patente).length) insertRow('Vehiculos', { patente: s.patente, idCliente: idCliente, vehiculo: s.vehiculo, anio: '', color: '' });
    const servicios = String(s.servicios).split(';').map(function (x) { return x.trim(); }).filter(String);
    const n = ACTIONS.addPresupuesto({ patente: s.patente, servicios: servicios, estado: 'Aprobado' }).n;
    ACTIONS.addTurno({ n: n, ingreso: p.ingreso || s.deseada, notas: 'Reserva online ' + s.id });
    setField('Solicitudes', 'id', p.id, 'estado', 'Confirmada');
    setField('Solicitudes', 'id', p.id, 'n', n);
    return { n: n, idCliente: idCliente };
  },
  rechazarSolicitud: function (p) {
    setField('Solicitudes', 'id', p.id, 'estado', 'Rechazada');
    return {};
  },
  addTrabajos: function (p) {
    (p.items || []).forEach(function (it) {
      const id = nextId('Trabajos', 'id', 'W', 3);
      insertRow('Trabajos', { id: id, n: it.n || '', inicio: toDate(it.inicio), fin: toDate(it.fin), patente: it.patente,
        servicio: it.servicio, precio: Number(it.precio) || 0, medio: it.medio || '', notas: it.notas || '' });
    });
    return {};
  }
};

/* ---------------- lectura ---------------- */
function readAll(rol) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const out = {};
  Object.keys(SCHEMA).forEach(function (name) {
    if (ADMIN_SHEETS.indexOf(name) >= 0 && rol !== 'admin') { out[name] = []; return; }
    const cols = SCHEMA[name];
    const sh = ss.getSheetByName(name);
    const n = sh.getLastRow() - FIRST_ROW + 1;
    const rows = n > 0 ? sh.getRange(FIRST_ROW, 1, n, cols.length).getValues() : [];
    out[name] = rows.filter(function (r) { return r[0] !== '' && r[0] !== null; }).map(function (r) {
      const o = {};
      cols.forEach(function (k, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        o[k.replace(/^_/, '')] = v;
      });
      return o;
    });
  });
  const par = ss.getSheetByName('Parametros');
  const list = function (a1) { return par.getRange(a1).getValues().map(function (r) { return r[0]; }).filter(String); };
  return {
    servicios: out.Servicios,
    modelos: out.Marcas_Modelos,
    clientes: out.Clientes,
    vehiculos: out.Vehiculos,
    presupuestos: out.Presupuestos,
    turnos: out.Turnos,
    trabajos: out.Trabajos,
    productos: out.Productos,
    movimientos: out.Movimientos,
    contactos: out.Contactos,
    solicitudes: out.Solicitudes,
    params: {
      pattern: par.getRange('B13').getDisplayValue(),
      capacidad: par.getRange('B14').getValue(),
      telefonoNegocio: par.getRange('B15').getDisplayValue(),
      feriados: list('K5:K44').map(function (d) { return d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d); }),
      estadosTurno: list('G5:G14'),
      medios: list('H5:H14'),
      categorias: list('I5:I14'),
      tiposContacto: list('N5:N14'),
      unidades: list('O5:O14')
    }
  };
}

/* ---------------- escritura ---------------- */
// Escribe en la primera fila libre (columna A vacía). Si se pasa del rango con fórmulas,
// copia la fila anterior para arrastrar las fórmulas y limpia las columnas de carga.
function insertRow(name, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const cols = SCHEMA[name];
  const maxRows = sh.getMaxRows();
  const colA = sh.getRange(FIRST_ROW, 1, maxRows - FIRST_ROW + 1, 1).getValues();
  let idx = colA.findIndex(function (r) { return r[0] === '' || r[0] === null; });
  let row;
  if (idx === -1) {
    sh.insertRowAfter(maxRows);
    row = maxRows + 1;
  } else {
    row = FIRST_ROW + idx;
  }
  const lastCol = sh.getLastColumn();
  const hasFormula = sh.getRange(row, 1, 1, lastCol).getFormulas()[0].some(String);
  const prevHasFormula = row > FIRST_ROW && sh.getRange(row - 1, 1, 1, lastCol).getFormulas()[0].some(String);
  if (!hasFormula && prevHasFormula) {
    sh.getRange(row - 1, 1, 1, lastCol).copyTo(sh.getRange(row, 1, 1, lastCol));
  }
  cols.forEach(function (k, i) {
    if (k.charAt(0) === '_') return;
    const v = obj.hasOwnProperty(k) ? obj[k] : '';
    sh.getRange(row, i + 1).setValue(v === undefined || v === null ? '' : v);
  });
  return row;
}

function findRows(name, key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const col = SCHEMA[name].indexOf(key) + 1;
  const n = sh.getLastRow() - FIRST_ROW + 1;
  if (n <= 0) return [];
  const vals = sh.getRange(FIRST_ROW, col, n, 1).getValues();
  const rows = [];
  vals.forEach(function (r, i) { if (String(r[0]).toUpperCase() === String(value).toUpperCase() && r[0] !== '') rows.push(FIRST_ROW + i); });
  return rows;
}

function setField(name, key, value, field, newValue) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const col = SCHEMA[name].indexOf(field) + 1;
  const rows = findRows(name, key, value);
  if (!rows.length) throw new Error('No se encontró ' + value + ' en ' + name);
  rows.forEach(function (r) { sh.getRange(r, col).setValue(newValue); });
}

/* ---------------- link público de reservas ---------------- */
// Solo lo necesario para cotizar: servicios, modelos, calendario y ocupación por día (sin datos de clientes)
function catalogo() {
  const d = readAll('operador');
  const ocup = {};
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  d.turnos.forEach(function (t) {
    if (t.estado === 'Cancelado' || t.estado === 'Entregado' || !t.ingreso || !t.entrega) return;
    const a = new Date(t.ingreso + 'T12:00:00'), b = new Date(t.entrega + 'T12:00:00');
    for (let x = a; x <= b; x.setDate(x.getDate() + 1)) {
      const k = Utilities.formatDate(x, tz, 'yyyy-MM-dd'); ocup[k] = (ocup[k] || 0) + 1;
    }
  });
  return { servicios: d.servicios, modelos: d.modelos, params: d.params, ocupacion: ocup };
}

function solicitar(p) {
  if (p.website) throw new Error('No se pudo enviar');
  const d = readAll('operador');
  const m = d.modelos.filter(function (x) { return x.vehiculo === p.vehiculo; })[0];
  if (!m) throw new Error('Elegí la marca y el modelo');
  const patente = String(p.patente || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!/^[A-Z0-9]{6,7}$/.test(patente)) throw new Error('Revisá la patente');
  const servicios = (p.servicios || []).slice(0, 12);
  if (!servicios.length) throw new Error('Elegí al menos un servicio');
  let dias = 0, precio = 0;
  servicios.forEach(function (n) {
    const sv = d.servicios.filter(function (x) { return x.nombre === n; })[0];
    if (!sv) throw new Error('Servicio no disponible: ' + n);
    dias += Math.ceil(Number(sv.dias) * Number(m.ft || 1) * 2) / 2;
    precio += Number(sv.precio) * Number(m.fp || 1);
  });
  const nombre = String(p.nombre || '').trim().slice(0, 80), telefono = String(p.telefono || '').trim().slice(0, 30);
  if (nombre.length < 2) throw new Error('Falta tu nombre');
  if (telefono.replace(/\D/g, '').length < 8) throw new Error('Revisá tu WhatsApp');
  if (!p.deseada) throw new Error('Elegí un día');
  const id = nextId('Solicitudes', 'id', 'R', 4);
  insertRow('Solicitudes', { id: id, fecha: new Date(), nombre: nombre, telefono: telefono, vehiculo: p.vehiculo, patente: patente,
    servicios: servicios.join('; '), deseada: toDate(p.deseada), dias: dias, precio: Math.round(precio), notas: String(p.notas || '').slice(0, 300), estado: ESPERA, n: '' });
  return { id: id, dias: dias, precio: Math.round(precio) };
}

function getField(name, key, value, field) {
  const rows = findRows(name, key, value);
  if (!rows.length) return '';
  const col = SCHEMA[name].indexOf(field) + 1;
  return SpreadsheetApp.getActive().getSheetByName(name).getRange(rows[0], col).getValue();
}

function nextId(name, key, prefix, len) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const col = SCHEMA[name].indexOf(key) + 1;
  const n = sh.getLastRow() - FIRST_ROW + 1;
  let max = 0;
  if (n > 0) sh.getRange(FIRST_ROW, col, n, 1).getValues().forEach(function (r) {
    const num = parseInt(String(r[0]).replace(/\D/g, ''), 10);
    if (num > max) max = num;
  });
  return prefix + ('0000000' + (max + 1)).slice(-len);
}

/* ---------------- helpers ---------------- */
function toDate(s) {
  if (!s) return '';
  const p = String(s).slice(0, 10).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}
// Valida usuario + PIN contra la hoja Usuarios. Nunca devuelve los PIN.
function auth(usuario, pin) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const rows = sh.getRange(FIRST_ROW, 1, 20, 4).getDisplayValues();
  const u = String(usuario || '').trim().toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].trim().toLowerCase() === u && rows[i][2].trim() === String(pin || '').trim()) {
      return { usuario: rows[i][0].trim(), nombre: rows[i][1], rol: (rows[i][3] || 'operador').trim().toLowerCase() };
    }
  }
  Utilities.sleep(800); // frena intentos repetidos
  throw new Error('Usuario o PIN incorrecto');
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Ejecutá esta función una vez desde el editor para autorizar el acceso y probar la lectura.
function probar() {
  const d = readAll('admin');
  Logger.log('Clientes: %s · Servicios: %s · Modelos: %s · Productos: %s · Contactos: %s',
    d.clientes.length, d.servicios.length, d.modelos.length, d.productos.length, d.contactos.length);
}
