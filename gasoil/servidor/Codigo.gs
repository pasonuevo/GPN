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
    revisarClave(libro, datos.k);
    const candado = LockService.getScriptLock();
    candado.waitLock(20000);
    try {
      const hoja = libro.getSheetByName('Movimientos');
      const filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS.length).getValues() : [];
      const ids = new Set(filas.map(f => String(f[9])));
      (datos.add || []).forEach(r => {
        if (!r.id || ids.has(r.id)) return;
        ids.add(r.id);
        filas.push([fechaDe(r.fecha), r.tipo, r.maquina || '', Number(r.litros) || 0, r.comentario || '', r.quien || '', '', '', new Date(), r.id]);
      });
      const anular = new Set(datos.anular || []);
      filas.forEach(f => { if (anular.has(String(f[9])) && !f[7]) f[7] = 'Sí'; });
      ordenarYSaldar(filas);
      if (filas.length) hoja.getRange(2, 1, filas.length, COLS.length).setValues(filas);
      hoja.getRange(2, 1, Math.max(filas.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
      hoja.getRange(2, 9, Math.max(filas.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    } finally {
      candado.releaseLock();
    }
    return leer(datos.k);
  });
}

function leer(clave) {
  const libro = SpreadsheetApp.openById(HOJA_ID);
  revisarClave(libro, clave);
  const hoja = libro.getSheetByName('Movimientos');
  const filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS.length).getValues() : [];
  return {
    ok: true,
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

function revisarClave(libro, clave) {
  const buena = String(libro.getSheetByName('Clave').getRange('B1').getValue()).trim();
  if (!buena || String(clave).trim() !== buena) throw new Error('clave');
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
  Logger.log(leer(SpreadsheetApp.openById(HOJA_ID).getSheetByName('Clave').getRange('B1').getValue()).rows.length + ' movimientos');
}
