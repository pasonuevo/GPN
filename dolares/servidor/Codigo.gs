// Servidor de la app Divisas: guarda las operaciones y las cuentas en la hoja de Google donde se instala.
// Se pega en Extensiones → Apps Script de esa hoja y se implementa como aplicación web ("Cualquier usuario").
const ZONA = 'America/Caracas';
const COLS_OP = ['Fecha', 'Tipo', 'Cliente', 'Monto USD', 'Tasa', 'Total Bs', 'Banco', 'Ganancia USD', 'Nota',
  'Tasa costo', 'Ganancia Bs', 'Margen', 'Anotado por', 'Anulado', 'Registrado', 'ID'];
const COLS_CTA = ['Grupo', 'Nombre', 'Saldo', 'Nota', 'Actualizado', 'Borrada', 'ID'];
const OP = Object.fromEntries(COLS_OP.map((c, i) => [c, i]));
const CTA = Object.fromEntries(COLS_CTA.map((c, i) => [c, i]));
const TIPOS = ['Compra', 'Venta', 'USDT/EF'];
const GRUPOS = ['Banco', 'Bolívares', 'Por cobrar', 'Por pagar'];

function doGet(e) {
  return responder(() => leer(e.parameter.k));
}

function doPost(e) {
  return responder(() => {
    const datos = JSON.parse(e.postData.contents);
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const yo = usuario(libro, datos.k);
    const dueno = yo.rol === 'Dueño';
    const rechazados = [];
    const candado = LockService.getScriptLock();
    candado.waitLock(20000);
    try {
      // Operaciones: se agregan; solo el dueño anula.
      const hOp = hoja(libro, 'Operaciones', COLS_OP);
      const ops = filasDe(hOp, COLS_OP.length);
      const ids = new Set(ops.map(f => String(f[OP.ID])));
      (datos.add || []).forEach(r => {
        if (!r.id || ids.has(r.id) || TIPOS.indexOf(r.tipo) < 0) return;
        ids.add(r.id);
        const f = new Array(COLS_OP.length).fill('');
        f[OP.Fecha] = fechaDe(r.fecha);
        f[OP.Tipo] = r.tipo;
        f[OP.Cliente] = String(r.cliente || '').trim();
        f[OP['Monto USD']] = Number(r.usd) || 0;
        f[OP.Tasa] = r.tipo === 'USDT/EF' ? '' : Number(r.tasa) || 0;
        f[OP['Total Bs']] = r.tipo === 'USDT/EF' ? '' : Number(r.bs) || 0;
        f[OP.Banco] = r.tipo === 'USDT/EF' ? '' : String(r.banco || '');
        f[OP['Ganancia USD']] = r.tipo === 'USDT/EF' ? Number(r.gan) || 0 : '';
        f[OP.Nota] = String(r.nota || '');
        f[OP['Anotado por']] = yo.nombre;
        f[OP.Registrado] = new Date();
        f[OP.ID] = r.id;
        ops.push(f);
      });
      const anular = new Set(datos.anular || []);
      const marca = 'Sí · ' + yo.nombre + ' ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM HH:mm');
      ops.forEach(f => {
        if (!anular.has(String(f[OP.ID])) || f[OP.Anulado]) return;
        if (dueno) f[OP.Anulado] = marca; else rechazados.push(String(f[OP.ID]));
      });
      ordenarYCalcular(ops);
      escribir(hOp, ops, COLS_OP.length);
      if (ops.length) {
        hOp.getRange(2, OP.Fecha + 1, ops.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
        hOp.getRange(2, OP.Registrado + 1, ops.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
        hOp.getRange(2, OP.Margen + 1, ops.length, 1).setNumberFormat('0.00%');
      }

      // Cuentas: gana el cambio más reciente.
      const hCta = hoja(libro, 'Cuentas', COLS_CTA);
      const ctas = filasDe(hCta, COLS_CTA.length);
      (datos.cuentas || []).forEach(c => {
        if (!c.id || GRUPOS.indexOf(c.grupo) < 0) return;
        const mod = new Date(Number(c.mod) || Date.now());
        let f = ctas.find(x => String(x[CTA.ID]) === c.id);
        if (f && f[CTA.Actualizado] && new Date(f[CTA.Actualizado]) >= mod) return;
        if (!f) { f = new Array(COLS_CTA.length).fill(''); f[CTA.ID] = c.id; ctas.push(f); }
        f[CTA.Grupo] = c.grupo;
        f[CTA.Nombre] = String(c.nombre || '').trim();
        f[CTA.Saldo] = Number(c.saldo) || 0;
        f[CTA.Nota] = String(c.nota || '');
        f[CTA.Actualizado] = mod;
        f[CTA.Borrada] = c.borrada ? 'Sí' : '';
      });
      ctas.sort((a, b) => GRUPOS.indexOf(a[CTA.Grupo]) - GRUPOS.indexOf(b[CTA.Grupo]));
      escribir(hCta, ctas, COLS_CTA.length);
      if (ctas.length) hCta.getRange(2, CTA.Actualizado + 1, ctas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    } finally {
      candado.releaseLock();
    }
    const salida = leer(datos.k);
    salida.rechazados = rechazados;
    return salida;
  });
}

function leer(clave) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const yo = usuario(libro, clave);
  const ops = filasDe(hoja(libro, 'Operaciones', COLS_OP), COLS_OP.length);
  const ctas = filasDe(hoja(libro, 'Cuentas', COLS_CTA), COLS_CTA.length);
  return {
    ok: true,
    yo: yo,
    rows: ops.filter(f => f[OP.ID]).map(f => ({
      id: String(f[OP.ID]),
      fecha: Utilities.formatDate(new Date(f[OP.Fecha]), ZONA, "yyyy-MM-dd'T'HH:mm"),
      tipo: f[OP.Tipo], cliente: String(f[OP.Cliente]), usd: Number(f[OP['Monto USD']]) || 0,
      tasa: Number(f[OP.Tasa]) || 0, bs: Number(f[OP['Total Bs']]) || 0, banco: String(f[OP.Banco]),
      gan: Number(f[OP['Ganancia USD']]) || 0, nota: String(f[OP.Nota]), quien: String(f[OP['Anotado por']]),
      anulado: !!f[OP.Anulado]
    })),
    cuentas: ctas.filter(f => f[CTA.ID] && !f[CTA.Borrada]).map(f => ({
      id: String(f[CTA.ID]), grupo: f[CTA.Grupo], nombre: String(f[CTA.Nombre]),
      saldo: Number(f[CTA.Saldo]) || 0, nota: String(f[CTA.Nota]),
      mod: f[CTA.Actualizado] ? new Date(f[CTA.Actualizado]).getTime() : 0
    }))
  };
}

// Igual que RO: por cada mes, la tasa costo es el promedio ponderado de todas las compras del mes,
// y cada venta gana (tasa pactada − tasa costo) × monto.
function ordenarYCalcular(ops) {
  ops.sort((a, b) => new Date(a[OP.Fecha]) - new Date(b[OP.Fecha]));
  const mes = f => Utilities.formatDate(new Date(f[OP.Fecha]), ZONA, 'yyyy-MM');
  const compras = {};
  ops.forEach(f => {
    if (f[OP.Anulado] || f[OP.Tipo] !== 'Compra') return;
    const c = compras[mes(f)] = compras[mes(f)] || { usd: 0, bs: 0 };
    c.usd += Number(f[OP['Monto USD']]) || 0;
    c.bs += Number(f[OP['Total Bs']]) || 0;
  });
  ops.forEach(f => {
    f[OP['Tasa costo']] = ''; f[OP['Ganancia Bs']] = ''; f[OP.Margen] = '';
    if (f[OP.Anulado]) return;
    if (f[OP.Tipo] === 'Compra') {
      f[OP['Tasa costo']] = Number(f[OP.Tasa]) || 0; f[OP['Ganancia Bs']] = 0; f[OP.Margen] = 0;
    } else if (f[OP.Tipo] === 'Venta') {
      const c = compras[mes(f)];
      if (!c || !c.usd) return;
      const costo = c.bs / c.usd;
      const bs = Number(f[OP['Total Bs']]) || 0;
      const gan = bs - (Number(f[OP['Monto USD']]) || 0) * costo;
      f[OP['Tasa costo']] = costo; f[OP['Ganancia Bs']] = gan; f[OP.Margen] = bs ? gan / bs : 0;
    }
  });
}

// Busca la clave en la pestaña Usuarios (Nombre, Clave, Rol, Activo).
function usuario(libro, clave) {
  const k = String(clave == null ? '' : clave).trim();
  const filas = filasDe(hoja(libro, 'Usuarios', ['Nombre', 'Clave', 'Rol', 'Activo']), 4);
  const f = k && filas.find(f => String(f[1]).trim() === k && String(f[0]).trim() && String(f[3]).trim().toLowerCase() !== 'no');
  if (!f) throw new Error('clave');
  return { nombre: String(f[0]).trim(), rol: String(f[2]).trim() === 'Dueño' ? 'Dueño' : 'Operador' };
}

// La pestaña con sus títulos; si no existe, la crea.
function hoja(libro, nombre, cols) {
  let h = libro.getSheetByName(nombre);
  if (!h) {
    h = libro.insertSheet(nombre);
    h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function filasDe(h, n) {
  return h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, n).getValues() : [];
}

function escribir(h, filas, n) {
  if (filas.length) h.getRange(2, 1, filas.length, n).setValues(filas);
}

function fechaDe(texto) {
  const t = String(texto).slice(0, 16);
  return /^\d{4}-\d\d-\d\dT\d\d:\d\d$/.test(t) ? Utilities.parseDate(t, ZONA, "yyyy-MM-dd'T'HH:mm") : new Date();
}

function responder(fn) {
  let salida;
  try { salida = fn(); } catch (err) { salida = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

// Correr una vez desde el editor: da permisos, crea las pestañas y, si no hay usuarios, crea la clave del dueño.
function preparar() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  hoja(libro, 'Operaciones', COLS_OP);
  hoja(libro, 'Cuentas', COLS_CTA);
  const u = hoja(libro, 'Usuarios', ['Nombre', 'Clave', 'Rol', 'Activo']);
  if (u.getLastRow() < 2) {
    const clave = String(Math.floor(100000 + Math.random() * 900000));
    u.appendRow(['Guillermo', clave, 'Dueño', 'Sí']);
    u.getRange(2, 2).setNumberFormat('@').setValue(clave);
  }
  Logger.log('Listo. Su clave está en la pestaña Usuarios: ' + u.getRange(2, 2).getValue());
}
