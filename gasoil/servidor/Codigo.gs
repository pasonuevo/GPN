// Servidor de la app de gasoil: guarda y entrega los movimientos de la hoja "Gasoil Paso Nuevo".
// Se pega en Apps Script (script.google.com) y se implementa como aplicación web ("Cualquier usuario").
const HOJA_ID = '1WJUGq9Wufpa8zzgdklmaCW1GS8GKAgHpdgZC2O-Yx8w';
const ZONA = 'America/Caracas';
const COLS = ['Fecha', 'Tipo', 'Máquina', 'Litros', 'Comentario', 'Anotado por', 'Saldo', 'Anulado', 'Registrado', 'ID'];

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
    const candado = LockService.getScriptLock();
    candado.waitLock(20000);
    try {
      const hoja = libro.getSheetByName('Movimientos');
      const filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS.length).getValues() : [];
      const ids = new Set(filas.map(f => String(f[9])));
      (datos.add || []).forEach(r => {
        if (!r.id || ids.has(r.id)) return;
        // Solo el dueño corrige el nivel del tanque.
        if (r.tipo === 'Corrección' && !dueno) { rechazados.push(r.id); return; }
        ids.add(r.id);
        filas.push([fechaDe(r.fecha), r.tipo, r.maquina || '', Number(r.litros) || 0, r.comentario || '', yo.nombre, '', '', new Date(), r.id]);
      });
      // Solo el dueño puede anular movimientos.
      const anular = new Set(datos.anular || []);
      const marca = 'Sí · ' + yo.nombre + ' ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM HH:mm');
      filas.forEach(f => {
        if (!anular.has(String(f[9])) || f[7]) return;
        if (dueno) f[7] = marca; else rechazados.push(String(f[9]));
      });
      ordenarYSaldar(filas);
      if (filas.length) hoja.getRange(2, 1, filas.length, COLS.length).setValues(filas);
      hoja.getRange(2, 1, Math.max(filas.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
      hoja.getRange(2, 9, Math.max(filas.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    } finally {
      candado.releaseLock();
    }
    const salida = leer(datos.k);
    salida.rechazados = rechazados;
    return salida;
  });
}

function leer(clave) {
  const libro = SpreadsheetApp.openById(HOJA_ID);
  const yo = usuario(libro, clave);
  const hoja = libro.getSheetByName('Movimientos');
  const filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS.length).getValues() : [];
  return {
    ok: true,
    yo: yo,
    rows: filas.filter(f => f[9]).map(f => ({
      id: String(f[9]),
      fecha: Utilities.formatDate(new Date(f[0]), ZONA, "yyyy-MM-dd'T'HH:mm"),
      tipo: f[1], maquina: f[2], litros: Number(f[3]) || 0, comentario: f[4], quien: f[5],
      anulado: !!f[7]
    }))
  };
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

// Correr una vez desde el editor para dar permisos.
function probar() {
  const usuarios = SpreadsheetApp.openById(HOJA_ID).getSheetByName('Usuarios').getDataRange().getValues();
  Logger.log(leer(usuarios[1][1]).rows.length + ' movimientos, ' + (usuarios.length - 1) + ' filas en Usuarios');
}
