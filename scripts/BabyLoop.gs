/**
 * ==================================================================
 * BABYLOOP · MENU DE OPERACION DEL INVENTARIO
 * ==================================================================
 * Vive dentro de tu propio Google Sheet. No cuesta nada, no hay
 * servidor que mantener y no depende de ningun servicio externo.
 *
 * Instalacion (una sola vez, desde computadora):
 *   1. Abre el Sheet
 *   2. Extensiones -> Apps Script
 *   3. Borra lo que haya y pega este archivo completo
 *   4. Guarda (icono de disco)
 *   5. Vuelve al Sheet y recarga la pagina
 *   6. Aparece el menu "BabyLoop" junto a Ayuda
 *   7. Menu BabyLoop -> Preparar hoja (crea la casilla de publicar)
 *      La primera vez Google te pide autorizar el script. Es tuyo,
 *      corre con tu propia cuenta y no comparte datos con nadie.
 * ==================================================================
 */

// ------------------------------------------------------------------
// CONFIGURACION
// ------------------------------------------------------------------

var HOJA = 'inventario';          // nombre de la pestaña del inventario
var DROPS_VIGENTES = 3;           // cuantos drops se quedan publicables
var MINIMO_POR_DROP = 5;          // piezas minimas para contar como drop
var COL_CASILLA = 'publicar';     // columna de casilla para aprobar desde el celular
var COL_VIGENCIA = 'vigencia';    // columna que maneja el sistema: vigente / archivado

// De aqui bajan las correcciones calculadas al releer los captions.
var URL_CORRECCIONES =
  'https://cdn.jsdelivr.net/gh/babyloopmx/babyloop-catalogo@main/datos/correcciones.json';


// ------------------------------------------------------------------
// MENU
// ------------------------------------------------------------------

/**
 * Se ejecuta solo al abrir el Sheet y dibuja el menu.
 * Nota: los menus personalizados solo aparecen en computadora.
 * Para el celular esta la casilla "publicar" (ver mas abajo).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BabyLoop')
    .addItem('Publicar seleccionados', 'publicarSeleccionados')
    .addItem('Marcar como apartado', 'marcarApartado')
    .addItem('Marcar como vendido', 'marcarVendido')
    .addItem('Regresar a borrador', 'regresarABorrador')
    .addSeparator()
    .addItem('Archivar drops viejos', 'archivarDropsViejos')
    .addItem('Aplicar correcciones', 'aplicarCorrecciones')
    .addItem('Resumen del inventario', 'mostrarResumen')
    .addSeparator()
    .addItem('Preparar hoja', 'prepararHoja')
    .addToUi();
}


// ------------------------------------------------------------------
// UTILIDADES INTERNAS
// ------------------------------------------------------------------

/** Devuelve la hoja del inventario, o avisa si no la encuentra. */
function hojaInventario_() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  if (!hoja) {
    throw new Error('No encuentro la pestaña "' + HOJA + '". Revisa el nombre.');
  }
  return hoja;
}

/** Mapa de nombre de columna -> numero de columna (1 = A). */
function columnas_(hoja) {
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var mapa = {};
  for (var i = 0; i < encabezados.length; i++) {
    var nombre = String(encabezados[i]).trim().toLowerCase();
    if (nombre) { mapa[nombre] = i + 1; }
  }
  if (!mapa['estado']) {
    throw new Error('La hoja no tiene columna "estado".');
  }
  return mapa;
}

/**
 * Filas que el usuario tiene seleccionadas, sin contar el encabezado.
 * Basta con tocar una celda de la fila: no hace falta seleccionar toda la fila.
 */
function filasSeleccionadas_(hoja) {
  var rangos = hoja.getActiveRangeList();
  if (!rangos) { return []; }
  var filas = {};
  var lista = rangos.getRanges();
  for (var i = 0; i < lista.length; i++) {
    var inicio = lista[i].getRow();
    var alto = lista[i].getNumRows();
    for (var f = inicio; f < inicio + alto; f++) {
      if (f > 1) { filas[f] = true; }   // la fila 1 es el encabezado
    }
  }
  return Object.keys(filas).map(Number).sort(function (a, b) { return a - b; });
}

/** Escribe un estado en todas las filas seleccionadas y avisa cuantas cambio. */
function aplicarEstado_(nuevoEstado) {
  var hoja = hojaInventario_();
  var cols = columnas_(hoja);
  var filas = filasSeleccionadas_(hoja);
  var ui = SpreadsheetApp.getUi();

  if (!filas.length) {
    ui.alert('Selecciona al menos una fila del inventario y vuelve a intentar.');
    return;
  }

  for (var i = 0; i < filas.length; i++) {
    hoja.getRange(filas[i], cols['estado']).setValue(nuevoEstado);
    // Si la pieza se va, la casilla de publicar deja de tener sentido
    if (cols[COL_CASILLA] && nuevoEstado !== 'disponible') {
      hoja.getRange(filas[i], cols[COL_CASILLA]).setValue(false);
    }
  }

  ui.alert(filas.length + (filas.length === 1 ? ' pieza marcada como ' : ' piezas marcadas como ') + nuevoEstado + '.');
}


// ------------------------------------------------------------------
// ACCIONES DEL MENU
// ------------------------------------------------------------------

function publicarSeleccionados() { aplicarEstado_('disponible'); }
function marcarApartado()        { aplicarEstado_('apartado'); }
function marcarVendido()         { aplicarEstado_('vendido'); }
function regresarABorrador()     { aplicarEstado_('borrador'); }


/**
 * ARCHIVAR DROPS VIEJOS
 * ------------------------------------------------------------------
 * Por que no toca la columna estado:
 *   La columna estado tiene una lista desplegable que solo admite
 *   disponible / apartado / vendido. Escribirle un cuarto valor la
 *   hace responder "Valor no permitido" y aborta la escritura.
 *
 *   En vez de pelearse con esa regla, el archivado vive en su propia
 *   columna: vigencia. estado se queda como lo que tu editas desde el
 *   celular; vigencia la maneja el sistema. Nunca chocan.
 *
 * Un drop es una fecha con volumen (MINIMO_POR_DROP piezas o mas):
 * asi una publicacion suelta no cuenta como drop.
 */
function archivarDropsViejos() {
  var hoja = hojaInventario_();
  var cols = columnas_(hoja);
  var ui = SpreadsheetApp.getUi();

  if (!cols['fecha_alta']) {
    ui.alert('La hoja no tiene columna "fecha_alta".');
    return;
  }

  var ultimaFila = hoja.getLastRow();

  // Si la columna vigencia no existe todavia, se crea al final.
  var colVig = cols[COL_VIGENCIA];
  if (!colVig) {
    colVig = hoja.getLastColumn() + 1;
    hoja.getRange(1, colVig).setValue(COL_VIGENCIA);
    SpreadsheetApp.flush();
  }

  var estados = hoja.getRange(2, cols['estado'], ultimaFila - 1, 1).getValues();
  var fechas = hoja.getRange(2, cols['fecha_alta'], ultimaFila - 1, 1).getValues();
  var vigencia = hoja.getRange(2, colVig, ultimaFila - 1, 1).getValues();

  var conteo = {};
  for (var i = 0; i < fechas.length; i++) {
    var f = normalizarFecha_(fechas[i][0]);
    if (f) { conteo[f] = (conteo[f] || 0) + 1; }
  }
  var vigentes = {};
  Object.keys(conteo)
    .filter(function (f) { return conteo[f] >= MINIMO_POR_DROP; })
    .sort()
    .reverse()
    .slice(0, DROPS_VIGENTES)
    .forEach(function (f) { vigentes[f] = true; });

  var listaVigentes = Object.keys(vigentes).sort().reverse();
  var respuesta = ui.alert(
    'Archivar drops viejos',
    'Se quedan vigentes solo estos drops:\n\n' + listaVigentes.join('\n') +
    '\n\nTodo lo anterior queda marcado como archivado en la columna ' +
    '"' + COL_VIGENCIA + '". La columna estado no se toca. ¿Continuamos?',
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) { return; }

  var archivadas = 0, activas = 0;
  for (var j = 0; j < estados.length; j++) {
    var estado = String(estados[j][0]).trim().toLowerCase();
    var fecha = normalizarFecha_(fechas[j][0]);

    // Vendido y apartado son historial real: se quedan vigentes
    // para que sigan contando en el resumen y en el catalogo.
    if (estado === 'vendido' || estado === 'apartado' || vigentes[fecha]) {
      vigencia[j][0] = 'vigente';
      activas++;
    } else {
      vigencia[j][0] = 'archivado';
      archivadas++;
    }
  }

  hoja.getRange(2, colVig, vigencia.length, 1).setValues(vigencia);
  SpreadsheetApp.flush();

  ui.alert(
    'Archivado listo\n\n' +
    'Vigentes  : ' + activas + '\n' +
    'Archivadas: ' + archivadas + '\n\n' +
    'El catalogo solo muestra las vigentes que ademas tengan foto.');
}


/** Cuenta rapida de como esta el inventario hoy. */
function mostrarResumen() {
  var hoja = hojaInventario_();
  var cols = columnas_(hoja);
  var ultimaFila = hoja.getLastRow();
  var estados = hoja.getRange(2, cols['estado'], ultimaFila - 1, 1).getValues();

  var conteo = {};
  for (var i = 0; i < estados.length; i++) {
    var e = String(estados[i][0]).trim().toLowerCase() || '(vacio)';
    conteo[e] = (conteo[e] || 0) + 1;
  }

  var lineas = ['Inventario: ' + (ultimaFila - 1) + ' piezas', ''];
  var orden = ['borrador', 'disponible', 'apartado', 'vendido', 'archivo'];
  orden.forEach(function (e) {
    if (conteo[e]) {
      lineas.push(e + ': ' + conteo[e] + (e === 'disponible' || e === 'apartado' ? '  (se publica)' : ''));
      delete conteo[e];
    }
  });
  Object.keys(conteo).forEach(function (e) { lineas.push(e + ': ' + conteo[e]); });

  // Valor del inventario que hoy esta a la venta
  if (cols['precio']) {
    var precios = hoja.getRange(2, cols['precio'], ultimaFila - 1, 1).getValues();
    var suma = 0;
    for (var k = 0; k < estados.length; k++) {
      if (String(estados[k][0]).trim().toLowerCase() === 'disponible') {
        suma += numero_(precios[k][0]);
      }
    }
    lineas.push('', 'Valor del inventario disponible: $' + suma.toLocaleString('es-MX'));
  }

  SpreadsheetApp.getUi().alert('Resumen del inventario', lineas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * Crea la columna de casilla "publicar" al final de la hoja.
 * Esa casilla es lo que te deja aprobar desde el celular, donde
 * los menus personalizados no existen.
 */
function prepararHoja() {
  var hoja = hojaInventario_();
  var cols = columnas_(hoja);
  var ui = SpreadsheetApp.getUi();

  var ultima = hoja.getLastRow();

  // Si la columna ya existe, le devolvemos las casillas.
  // Esto hace falta despues de importar un CSV sobre la hoja:
  // el import reemplaza los valores y se lleva las casillas de paso.
  if (cols[COL_CASILLA]) {
    if (ultima > 1) {
      var existente = hoja.getRange(2, cols[COL_CASILLA], ultima - 1, 1);
      existente.insertCheckboxes();
      // Todo lo que no sea exactamente true queda destildado
      var valores = existente.getValues();
      for (var k = 0; k < valores.length; k++) {
        valores[k][0] = (valores[k][0] === true || String(valores[k][0]).toLowerCase() === 'true');
      }
      existente.setValues(valores);
    }
    ui.alert('Listo', 'Se restauraron las casillas de la columna "' + COL_CASILLA + '".', ui.ButtonSet.OK);
    return;
  }

  var nueva = hoja.getLastColumn() + 1;
  hoja.getRange(1, nueva).setValue(COL_CASILLA);
  if (ultima > 1) {
    var rango = hoja.getRange(2, nueva, ultima - 1, 1);
    rango.insertCheckboxes();
    rango.setValue(false);
  }
  hoja.setColumnWidth(nueva, 90);

  ui.alert(
    'Listo',
    'Se agrego la columna "' + COL_CASILLA + '" al final de la hoja.\n\n' +
    'Desde el celular: filtra por estado = borrador y palomea la casilla de las piezas ' +
    'que quieras publicar. Cada palomita cambia el estado a disponible sola.',
    ui.ButtonSet.OK
  );
}


// ------------------------------------------------------------------
// APROBACION DESDE EL CELULAR
// ------------------------------------------------------------------

/**
 * Se dispara solo con cada edicion de la hoja, incluida la app del celular.
 * No hay que instalar nada: Google ejecuta esta funcion por su nombre.
 * Si palomeas la casilla "publicar", la pieza pasa a disponible y la
 * casilla se limpia sola. Un toque por pieza, sin escribir nada.
 */
function onEdit(e) {
  if (!e || !e.range) { return; }
  var hoja = e.range.getSheet();
  if (hoja.getName() !== HOJA) { return; }
  if (e.range.getRow() === 1) { return; }

  var cols = columnas_(hoja);
  if (!cols[COL_CASILLA]) { return; }
  if (e.range.getColumn() !== cols[COL_CASILLA]) { return; }
  if (e.range.getValue() !== true) { return; }

  hoja.getRange(e.range.getRow(), cols['estado']).setValue('disponible');
  e.range.setValue(false);
}


// ------------------------------------------------------------------
// AYUDANTES
// ------------------------------------------------------------------

/** Convierte cualquier forma de fecha a texto AAAA-MM-DD. */
function normalizarFecha_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(valor || '').trim().slice(0, 10);
}

/** Convierte "$1,200" a 1200. */
function numero_(valor) {
  if (typeof valor === 'number') { return valor; }
  var limpio = String(valor || '').replace(/[^0-9.]/g, '');
  return parseFloat(limpio) || 0;
}


// ==================================================================
// CANAL DE CORRECCIONES
// ==================================================================
// Para que sirve:
//   Las correcciones que se calculan releyendo los captions de
//   Instagram viven en un archivo publico del repositorio. Esta
//   funcion lo baja y lo aplica sobre la hoja.
//
//   Asi las correcciones viajan por el repositorio y no hay que
//   pegar listas gigantes dentro del editor cada vez.
//
// Regla de oro: solo escribe donde la celda esta VACIA o dice
// [REVISAR]. Nada que ya tenga un valor bueno se toca.
// La columna estado no se toca nunca desde aqui.
// ==================================================================

function aplicarCorrecciones() {
  var hoja = hojaInventario_();
  var cols = columnas_(hoja);
  var ui = SpreadsheetApp.getUi();
  var ultima = hoja.getLastRow();
  if (ultima < 2) { return; }

  // 1. Bajar el archivo de correcciones
  var doc;
  try {
    var resp = UrlFetchApp.fetch(URL_CORRECCIONES, {muteHttpExceptions: true});
    if (resp.getResponseCode() !== 200) {
      throw new Error('el servidor respondio ' + resp.getResponseCode());
    }
    doc = JSON.parse(resp.getContentText());
  } catch (e) {
    ui.alert('No pude bajar las correcciones.\n\n' + e.message +
             '\n\nSi el archivo se acaba de subir al repositorio, el CDN ' +
             'tarda unos minutos en verlo. Intenta de nuevo mas tarde.');
    return;
  }
  var cambios = doc.cambios || {};

  // 2. Leer de un solo golpe todas las columnas que podriamos tocar
  var campos = ['nombre', 'categoria', 'talla', 'genero', 'condicion',
                'precio', 'marca', 'color', 'temporada', 'revisar'];
  var presentes = [];
  for (var c = 0; c < campos.length; c++) {
    if (cols[campos[c]]) { presentes.push(campos[c]); }
  }

  var ids = hoja.getRange(2, cols['id'], ultima - 1, 1).getValues();
  var datos = {};
  for (var p = 0; p < presentes.length; p++) {
    datos[presentes[p]] = hoja.getRange(2, cols[presentes[p]], ultima - 1, 1).getValues();
  }

  // 3. Aplicar en memoria
  var celdas = 0, tocadas = {};
  for (var f = 0; f < ids.length; f++) {
    var id = String(ids[f][0]).trim();
    var corr = cambios[id];
    if (!corr) { continue; }

    for (var k = 0; k < presentes.length; k++) {
      var campo = presentes[k];
      if (!(campo in corr)) { continue; }

      var actual = String(datos[campo][f][0]).trim();

      // La columna revisar si se reescribe siempre: es un recordatorio,
      // no un dato, y tiene que reflejar lo que sigue pendiente.
      var pendiente = (actual === '' || actual.charAt(0) === '[');
      if (campo !== 'revisar' && !pendiente) { continue; }

      var nuevo = String(corr[campo]);
      if (nuevo === actual) { continue; }

      datos[campo][f][0] = nuevo;
      celdas++;
      tocadas[id] = true;
    }
  }

  // 4. Una sola escritura por columna
  for (var w = 0; w < presentes.length; w++) {
    var campo2 = presentes[w];
    hoja.getRange(2, cols[campo2], datos[campo2].length, 1).setValues(datos[campo2]);
  }
  SpreadsheetApp.flush();

  ui.alert(
    'Correcciones aplicadas\n\n' +
    'Generadas el      : ' + (doc.generado || 'sin fecha') + '\n' +
    'Piezas corregidas : ' + Object.keys(tocadas).length + '\n' +
    'Celdas escritas   : ' + celdas + '\n\n' +
    'Solo se llenaron celdas vacias o marcadas [REVISAR].');
}
