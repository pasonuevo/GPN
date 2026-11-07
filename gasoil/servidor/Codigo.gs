// Servidor de la app de gasoil: guarda y entrega los movimientos y los servicios de la hoja "Gasoil Paso Nuevo".
// Se pega en Apps Script (script.google.com) y se implementa como aplicación web ("Cualquier usuario").
const HOJA_ID = '1WJUGq9Wufpa8zzgdklmaCW1GS8GKAgHpdgZC2O-Yx8w';
const ZONA = 'America/Caracas';
const VERSION = 3;
const COLS = ['Fecha', 'Tipo', 'Máquina', 'Litros', 'Comentario', 'Anotado por', 'Saldo', 'Anulado', 'Registrado', 'ID', 'Horas'];
const COLS_SERV = ['Fecha', 'Máquina', 'Horas', 'Trabajo', 'Comentario', 'Anotado por', 'Anulado', 'Registrado', 'ID'];
const COLS_PLAN = ['Máquina', 'Servicio', 'Cada (horas)', 'Activo'];
const MAQUINAS = ['Landini', 'John Deere', 'Fiat', 'Ford'];
// Intervalos provisionales: se cambian en la pestaña "Plan de servicios" según la ficha técnica de cada máquina.
const PLAN_INICIAL = [['Aceite de motor', 250], ['Filtro de aceite', 250], ['Filtro de gasoil', 500], ['Filtro de aire', 500], ['Aceite hidráulico', 1000], ['Filtro hidráulico', 1000]];

function doGet(e) {
  return responder(() => leer(e.parameter.k));
}

function doPost(e) {
  return responder(() => {
    const datos = JSON.parse(e.postData.contents);
    const libro = SpreadsheetApp.openById(HOJA_ID);
    const yo = usuario(libro, datos.k);
    const dueno = yo.rol === 'Dueño';
    const rechazados = [];
    const marca = 'Sí · ' + yo.nombre + ' ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM HH:mm');
    const candado = LockService.getScriptLock();
    candado.waitLock(20000);
    try {
      // Reiniciar desde cero (solo el dueño): el gasoil y los servicios van por separado.
      if (datos.reiniciar === 'gasoil' || datos.reiniciar === 'servicios') {
        if (!dueno) throw new Error('solo el dueño puede reiniciar');
        reiniciar(libro, datos.reiniciar);
      }
      const props = PropertiesService.getScriptProperties();
      // 10 minutos de margen por si el reloj de algún teléfono está atrasado.
      const desdeGas = (Number(props.getProperty('reinicio_gasoil')) || 0) - 600000;
      const desdeServ = (Number(props.getProperty('reinicio_servicios')) || 0) - 600000;

      // Movimientos de gasoil
      const hoja = hojas(libro).mov;
      const filas = leerFilas(hoja, COLS.length);
      const ids = new Set(filas.map(f => String(f[9])));
      (datos.add || []).forEach(r => {
        if (!r.id || ids.has(r.id)) return;
        // Lo anotado en un teléfono antes del último reinicio ya no entra.
        if (creadoEn(r.id) < desdeGas) { rechazados.push(r.id); return; }
        // Solo el dueño surte gasoil: tanquear, entradas y correcciones del tanque.
        if (!dueno) { rechazados.push(r.id); return; }
        ids.add(r.id);
        filas.push([fechaDe(r.fecha), r.tipo, r.maquina || '', Number(r.litros) || 0, r.comentario || '', yo.nombre, '', '', new Date(), r.id, Number(r.horas) > 0 ? Number(r.horas) : '']);
      });
      // Solo el dueño puede anular movimientos.
      const anular = new Set(datos.anular || []);
      filas.forEach(f => {
        if (!anular.has(String(f[9])) || f[7]) return;
        if (dueno) f[7] = marca; else rechazados.push(String(f[9]));
      });
      ordenarYSaldar(filas);
      escribir(hoja, filas, COLS.length, [1, 9]);

      // Servicios de las máquinas
      const hServ = hojas(libro).serv;
      const serv = leerFilas(hServ, COLS_SERV.length);
      const idsServ = new Set(serv.map(f => String(f[8])));
      (datos.addServ || []).forEach(r => {
        if (!r.id || idsServ.has(r.id)) return;
        if (creadoEn(r.id) < desdeServ) { rechazados.push(r.id); return; }
        idsServ.add(r.id);
        serv.push([fechaDe(r.fecha), r.maquina || '', Number(r.horas) || 0, (r.trabajo || []).join(', '), r.comentario || '', yo.nombre, '', new Date(), r.id]);
      });
      const anularServ = new Set(datos.anularServ || []);
      serv.forEach(f => {
        if (!anularServ.has(String(f[8])) || f[6]) return;
        if (dueno) f[6] = marca; else rechazados.push(String(f[8]));
      });
      serv.sort((a, b) => new Date(a[0]) - new Date(b[0]));
      escribir(hServ, serv, COLS_SERV.length, [1, 8]);
    } finally {
      candado.releaseLock();
    }
    const salida = leer(datos.k);
    salida.rechazados = rechazados;
    return salida;
  });
}

// Guarda una copia de la pestaña (Movimientos o Servicios) y la deja solo con los encabezados.
// No toca las máquinas, el plan de servicios, los usuarios ni la clave.
function reiniciar(libro, que) {
  const h = hojas(libro);
  const hoja = que === 'gasoil' ? h.mov : h.serv;
  if (hoja.getLastRow() > 1) {
    const nombre = 'Respaldo ' + (que === 'gasoil' ? 'gasoil' : 'servicios') + ' ' + Utilities.formatDate(new Date(), ZONA, 'dd-MM-yyyy HH.mm');
    hoja.copyTo(libro).setName(nombre).setTabColor('#999999');
    hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getMaxColumns()).clearContent();
  }
  PropertiesService.getScriptProperties().setProperty('reinicio_' + que, String(Date.now()));
}

// Momento en que el teléfono creó el registro (va al principio del ID).
function creadoEn(id) {
  return parseInt(String(id).split('-')[0], 36) || 0;
}

function leer(clave) {
  const libro = SpreadsheetApp.openById(HOJA_ID);
  const yo = usuario(libro, clave);
  const h = hojas(libro);
  const hora = f => Utilities.formatDate(new Date(f), ZONA, "yyyy-MM-dd'T'HH:mm");
  return {
    ok: true,
    version: VERSION,
    yo: yo,
    rows: leerFilas(h.mov, COLS.length).filter(f => f[9]).map(f => ({
      id: String(f[9]), fecha: hora(f[0]),
      tipo: f[1], maquina: f[2], litros: Number(f[3]) || 0, comentario: f[4], quien: f[5],
      anulado: !!f[7], horas: Number(f[10]) || 0
    })),
    servicios: leerFilas(h.serv, COLS_SERV.length).filter(f => f[8]).map(f => ({
      id: String(f[8]), fecha: hora(f[0]), maquina: String(f[1]), horas: Number(f[2]) || 0,
      trabajo: String(f[3]).split(',').map(s => s.trim()).filter(Boolean), comentario: f[4], quien: f[5],
      anulado: !!f[6]
    })),
    plan: leerFilas(h.plan, COLS_PLAN.length)
      .filter(f => String(f[0]).trim() && String(f[1]).trim() && Number(f[2]) > 0 && String(f[3]).trim().toLowerCase() !== 'no')
      .map(f => ({ maquina: String(f[0]).trim(), servicio: String(f[1]).trim(), cada: Number(f[2]) }))
  };
}

// Devuelve las tres pestañas; crea las que falten con sus encabezados.
function hojas(libro) {
  const mov = libro.getSheetByName('Movimientos');
  anchoMinimo(mov, COLS.length);
  if (!mov.getRange(1, 11).getValue()) mov.getRange(1, 11).setValue('Horas').setFontWeight('bold');
  let serv = libro.getSheetByName('Servicios');
  if (!serv) {
    serv = libro.insertSheet('Servicios');
    serv.getRange(1, 1, 1, COLS_SERV.length).setValues([COLS_SERV]).setFontWeight('bold');
    serv.setFrozenRows(1);
  }
  anchoMinimo(serv, COLS_SERV.length);
  let plan = libro.getSheetByName('Plan de servicios');
  if (!plan) {
    plan = libro.insertSheet('Plan de servicios');
    const filas = [COLS_PLAN];
    MAQUINAS.forEach(m => PLAN_INICIAL.forEach(([s, h]) => filas.push([m, s, h, 'Sí'])));
    plan.getRange(1, 1, filas.length, COLS_PLAN.length).setValues(filas);
    plan.getRange(1, 1, 1, COLS_PLAN.length).setFontWeight('bold');
    plan.setFrozenRows(1);
  }
  return { mov, serv, plan };
}

function anchoMinimo(hoja, n) {
  if (hoja.getMaxColumns() < n) hoja.insertColumnsAfter(hoja.getMaxColumns(), n - hoja.getMaxColumns());
}

function leerFilas(hoja, n) {
  return hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, n).getValues() : [];
}

function escribir(hoja, filas, n, colsFecha) {
  if (filas.length) hoja.getRange(2, 1, filas.length, n).setValues(filas);
  colsFecha.forEach(c => hoja.getRange(2, c, Math.max(filas.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm'));
}

// Ordena por fecha y calcula el saldo: Entrada suma, Salida resta, Corrección fija el nivel medido.
function ordenarYSaldar(filas) {
  filas.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  let saldo = 0;
  filas.forEach(f => {
    if (!f[7]) {
      if (f[1] === 'Entrada') saldo += Number(f[3]) || 0;
      else if (f[1] === 'Salida') saldo -= Number(f[3]) || 0;
      else if (f[1] === 'Corrección') saldo = Number(f[3]) || 0;
    }
    f[6] = saldo;
  });
}

// Busca la clave en la pestaña Usuarios (Nombre, Clave, Rol, Activo).
function usuario(libro, clave) {
  const k = String(clave == null ? '' : clave).trim();
  const hoja = libro.getSheetByName('Usuarios');
  const filas = hoja && hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues() : [];
  const f = k && filas.find(f => String(f[1]).trim() === k && String(f[0]).trim() && String(f[3]).trim().toLowerCase() !== 'no');
  if (!f) throw new Error('clave');
  return { nombre: String(f[0]).trim(), rol: String(f[2]).trim() === 'Dueño' ? 'Dueño' : 'Surtidor' };
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

// Correr una vez desde el editor para dar permisos (y crear las pestañas de servicios).
function probar() {
  const usuarios = SpreadsheetApp.openById(HOJA_ID).getSheetByName('Usuarios').getDataRange().getValues();
  const d = leer(usuarios[1][1]);
  Logger.log(d.rows.length + ' movimientos, ' + d.servicios.length + ' servicios, ' + d.plan.length + ' filas de plan, ' + (usuarios.length - 1) + ' usuarios');
}
