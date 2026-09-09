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

  // 1. La tabla de correcciones va embebida al final de este archivo.
  //    Se embebe en vez de bajarla por internet a proposito: bajarla
  //    obligaria a autorizar de nuevo el script con un permiso extra
  //    (UrlFetchApp), y no vale la pena por un archivo que cambia
  //    una vez por drop.
  var doc = CORRECCIONES;
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


// ==================================================================
// TABLA DE CORRECCIONES
// ==================================================================
// Generada al releer los captions del export de Instagram.
// La usa aplicarCorrecciones(). Para actualizarla se reemplaza este
// bloque completo; el resto del archivo no se toca.
// Fuente en el repositorio: datos/correcciones.json
// ==================================================================

var CORRECCIONES = {"generado":"2026-09-09","origen":"instagram-baby.loop.mx-2026-07-06 · relectura de captions","nota":"Solo se aplican a celdas vacias o marcadas [REVISAR]. Nada que ya tenga un valor bueno se toca.","cambios":{"BL-0740":{"revisar":""},"BL-0735":{"revisar":"categoria"},"BL-0728":{"condicion":"seminuevo","revisar":"categoria"},"BL-0716":{"condicion":"seminuevo","revisar":""},"BL-0710":{"condicion":"seminuevo","revisar":"categoria"},"BL-0704":{"condicion":"seminuevo","revisar":"categoria"},"BL-0691":{"nombre":"Mameluco Soda City. · 9m","categoria":"mameluco","revisar":""},"BL-0689":{"condicion":"seminuevo","revisar":"categoria"},"BL-0687":{"condicion":"seminuevo","marca":"Intacto","revisar":"categoria"},"BL-0686":{"condicion":"seminuevo","revisar":""},"BL-0684":{"condicion":"seminuevo","revisar":""},"BL-0683":{"condicion":"seminuevo","marca":"Intacto","revisar":""},"BL-0682":{"condicion":"seminuevo","revisar":""},"BL-0679":{"marca":"Sin detalles"},"BL-0675":{"condicion":"seminuevo","revisar":"categoria"},"BL-0674":{"marca":"Sin detalles"},"BL-0673":{"marca":"Sin detalles"},"BL-0672":{"condicion":"seminuevo","revisar":""},"BL-0671":{"condicion":"seminuevo","marca":"Intacto","revisar":""},"BL-0669":{"marca":"Sin detalles"},"BL-0668":{"condicion":"seminuevo","marca":"Intacto","revisar":"categoria"},"BL-0667":{"marca":"Sin detalles"},"BL-0664":{"marca":"Sin detalles"},"BL-0661":{"condicion":"seminuevo","revisar":""},"BL-0659":{"condicion":"seminuevo","revisar":"categoria"},"BL-0657":{"condicion":"seminuevo","marca":"Intacto","revisar":"categoria"},"BL-0651":{"nombre":"Sueter Pili Carrera · 4-5a","categoria":"sueter/chamarra","condicion":"seminuevo","revisar":""},"BL-0648":{"condicion":"seminuevo","revisar":""},"BL-0647":{"condicion":"seminuevo","revisar":""},"BL-0646":{"condicion":"seminuevo","revisar":"categoria"},"BL-0643":{"revisar":"categoria, condicion"},"BL-0633":{"condicion":"seminuevo","revisar":"categoria"},"BL-0631":{"condicion":"seminuevo","revisar":"categoria"},"BL-0622":{"revisar":"categoria"},"BL-0604":{"condicion":"seminuevo","revisar":"categoria"},"BL-0596":{"nombre":"Pantalon Bonnie Jean · 2-3a","categoria":"pantalon/short","revisar":"condicion"},"BL-0585":{"condicion":"seminuevo","marca":"Intacto","revisar":""},"BL-0582":{"nombre":"Mameluco Ralph Lauren · 9m","categoria":"mameluco","condicion":"seminuevo","revisar":""},"BL-0578":{"revisar":"categoria"},"BL-0575":{"revisar":"condicion"},"BL-0574":{"condicion":"seminuevo","revisar":"categoria"},"BL-0573":{"condicion":"seminuevo","revisar":"categoria"},"BL-0572":{"nombre":"Pantalon Bonnie Jean · 6-8a","categoria":"pantalon/short","revisar":""},"BL-0568":{"condicion":"seminuevo","revisar":""},"BL-0559":{"marca":"Sin detalles"},"BL-0555":{"condicion":"seminuevo","revisar":"categoria"},"BL-0548":{"nombre":"Pantalon Bonnie Jean. · 6-8a","categoria":"pantalon/short","revisar":""},"BL-0544":{"condicion":"seminuevo","revisar":"categoria"},"BL-0543":{"condicion":"seminuevo","revisar":"categoria"},"BL-0542":{"condicion":"seminuevo","revisar":"categoria"},"BL-0541":{"condicion":"seminuevo","revisar":"categoria"},"BL-0532":{"condicion":"seminuevo","revisar":"categoria"},"BL-0531":{"condicion":"seminuevo","revisar":"categoria"},"BL-0522":{"nombre":"Pantalon Bonnie Jean · 2-3a","categoria":"pantalon/short","revisar":""},"BL-0511":{"condicion":"seminuevo","revisar":"categoria"},"BL-0508":{"condicion":"seminuevo","revisar":""},"BL-0506":{"condicion":"seminuevo","revisar":"categoria"},"BL-0505":{"revisar":"condicion, talla"},"BL-0489":{"condicion":"seminuevo","revisar":"categoria"},"BL-0488":{"condicion":"seminuevo","revisar":""},"BL-0487":{"condicion":"seminuevo","revisar":"categoria"},"BL-0486":{"condicion":"seminuevo","revisar":"categoria"},"BL-0484":{"condicion":"seminuevo","revisar":"categoria"},"BL-0469":{"condicion":"seminuevo","revisar":"categoria"},"BL-0465":{"condicion":"seminuevo","revisar":"categoria"},"BL-0463":{"condicion":"seminuevo","revisar":"categoria"},"BL-0460":{"condicion":"seminuevo","revisar":"categoria"},"BL-0458":{"condicion":"seminuevo","revisar":"categoria"},"BL-0456":{"condicion":"seminuevo","revisar":"categoria"},"BL-0455":{"condicion":"seminuevo","revisar":"categoria"},"BL-0451":{"condicion":"seminuevo","revisar":"categoria"},"BL-0450":{"condicion":"seminuevo","revisar":"categoria"},"BL-0449":{"condicion":"seminuevo","revisar":"categoria"},"BL-0436":{"condicion":"seminuevo","revisar":"categoria"},"BL-0421":{"condicion":"seminuevo","revisar":"categoria"},"BL-0420":{"condicion":"seminuevo","revisar":"categoria"},"BL-0411":{"condicion":"seminuevo","revisar":"categoria"},"BL-0407":{"condicion":"seminuevo","revisar":"categoria"},"BL-0406":{"condicion":"seminuevo","revisar":"categoria"},"BL-0405":{"condicion":"seminuevo","revisar":""},"BL-0399":{"marca":"Sin detalles"},"BL-0389":{"condicion":"seminuevo","revisar":"categoria"},"BL-0381":{"marca":"Sin detalles"},"BL-0380":{"condicion":"seminuevo","revisar":"categoria"},"BL-0379":{"condicion":"seminuevo","revisar":"categoria"},"BL-0378":{"condicion":"seminuevo","revisar":"categoria"},"BL-0362":{"condicion":"seminuevo","revisar":"categoria"},"BL-0361":{"condicion":"seminuevo","revisar":"categoria"},"BL-0358":{"condicion":"seminuevo","revisar":"categoria"},"BL-0354":{"condicion":"seminuevo","revisar":""},"BL-0350":{"condicion":"seminuevo","revisar":""},"BL-0346":{"condicion":"seminuevo","revisar":"categoria"},"BL-0345":{"condicion":"seminuevo","revisar":"categoria"},"BL-0344":{"nombre":"Accesorio Cinturón Janie&Jack · 2-3a","categoria":"accesorio","revisar":""},"BL-0343":{"precio":"60","revisar":"categoria, condicion"},"BL-0342":{"condicion":"seminuevo","revisar":"talla"},"BL-0341":{"revisar":"categoria, condicion, talla"},"BL-0339":{"revisar":"condicion, talla"},"BL-0331":{"condicion":"seminuevo","revisar":"categoria"},"BL-0329":{"condicion":"seminuevo","revisar":"categoria"},"BL-0327":{"talla":"24m","revisar":"categoria, condicion"},"BL-0326":{"talla":"24m","revisar":"categoria"},"BL-0325":{"condicion":"seminuevo","revisar":"categoria"},"BL-0324":{"condicion":"seminuevo","revisar":"categoria"},"BL-0322":{"condicion":"seminuevo","revisar":""},"BL-0321":{"condicion":"seminuevo","revisar":"categoria"},"BL-0320":{"condicion":"seminuevo","revisar":"categoria"},"BL-0319":{"condicion":"seminuevo","revisar":"categoria"},"BL-0316":{"condicion":"seminuevo","revisar":"categoria"},"BL-0315":{"condicion":"seminuevo","revisar":"categoria"},"BL-0312":{"condicion":"seminuevo","revisar":"categoria"},"BL-0310":{"condicion":"seminuevo","revisar":"categoria"},"BL-0309":{"condicion":"seminuevo","revisar":"categoria"},"BL-0308":{"condicion":"seminuevo","revisar":""},"BL-0307":{"condicion":"seminuevo","revisar":"categoria"},"BL-0306":{"condicion":"seminuevo","revisar":""},"BL-0305":{"condicion":"seminuevo","revisar":""},"BL-0302":{"talla":"18m","condicion":"seminuevo","revisar":"categoria"},"BL-0301":{"precio":"149","revisar":"categoria"},"BL-0300":{"precio":"600","revisar":"categoria"},"BL-0299":{"condicion":"seminuevo","revisar":"categoria"},"BL-0298":{"talla":"24m","revisar":"condicion"},"BL-0297":{"condicion":"seminuevo","precio":"249","revisar":"categoria"},"BL-0296":{"precio":"199","revisar":"condicion"},"BL-0294":{"condicion":"seminuevo","precio":"119","revisar":""},"BL-0293":{"condicion":"seminuevo","precio":"129","revisar":"categoria"},"BL-0292":{"condicion":"seminuevo","revisar":"categoria"},"BL-0291":{"condicion":"seminuevo","precio":"169","revisar":""},"BL-0289":{"condicion":"seminuevo","revisar":"categoria"},"BL-0288":{"condicion":"seminuevo","precio":"169","revisar":""},"BL-0287":{"nombre":"Vestido Gap · 10-12a","categoria":"vestido","precio":"169","revisar":""},"BL-0286":{"condicion":"seminuevo","precio":"600","revisar":"categoria"},"BL-0285":{"precio":"249","revisar":"condicion"},"BL-0284":{"condicion":"seminuevo","precio":"249","revisar":""},"BL-0283":{"precio":"99","revisar":"categoria"},"BL-0282":{"precio":"249","revisar":"condicion"},"BL-0278":{"condicion":"nuevo con etiqueta","revisar":""},"BL-0265":{"revisar":"condicion"},"BL-0264":{"condicion":"seminuevo","revisar":""},"BL-0263":{"marca":"Conjunto español"},"BL-0261":{"revisar":"condicion"},"BL-0259":{"condicion":"seminuevo","revisar":""},"BL-0245":{"marca":"Conjunto armado"},"BL-0229":{"nombre":"Vestido Healthex · 9m","categoria":"vestido","revisar":"condicion"},"BL-0223":{"nombre":"Mameluco Nicole Miller · 6m","categoria":"mameluco","revisar":"condicion"},"BL-0220":{"nombre":"Mameluco VENDIDO Paz Rodriguez · 3m","categoria":"mameluco","revisar":"condicion"},"BL-0215":{"nombre":"Sueter Zara · 2-3a","categoria":"sueter/chamarra","revisar":""},"BL-0214":{"nombre":"Sueter Pili Carrera · 4-5a","categoria":"sueter/chamarra","condicion":"seminuevo","revisar":""},"BL-0209":{"revisar":"categoria"},"BL-0202":{"nombre":"Sueter H&M · 9m","categoria":"sueter/chamarra","revisar":"condicion"},"BL-0200":{"talla":"6m","revisar":""},"BL-0197":{"revisar":""},"BL-0193":{"talla":"6-8a","revisar":""},"BL-0191":{"talla":"4-5a","marca":"Conjunto armado","revisar":""},"BL-0186":{"talla":"4-5a","revisar":""},"BL-0184":{"nombre":"Mameluco Janie And Jack · 10-12a","categoria":"mameluco","revisar":""},"BL-0180":{"talla":"24m","revisar":""},"BL-0179":{"talla":"24m","revisar":"categoria"},"BL-0175":{"condicion":"seminuevo","revisar":""},"BL-0172":{"revisar":""},"BL-0169":{"nombre":"Mameluco Mayoral · 18m","categoria":"mameluco","condicion":"seminuevo","revisar":""},"BL-0167":{"nombre":"Mameluco Mayoral · 18m","categoria":"mameluco","condicion":"seminuevo","revisar":""},"BL-0163":{"nombre":"Vestido Angel dear · 9m","categoria":"vestido","revisar":""},"BL-0161":{"condicion":"seminuevo","revisar":""},"BL-0159":{"talla":"10-12a","condicion":"seminuevo","revisar":""},"BL-0157":{"condicion":"seminuevo","revisar":""},"BL-0155":{"revisar":"categoria"},"BL-0154":{"condicion":"seminuevo","revisar":""},"BL-0152":{"talla":"9m","condicion":"seminuevo","revisar":""},"BL-0147":{"nombre":"Vestido Mayoral · 18m","categoria":"vestido","revisar":""},"BL-0145":{"nombre":"Vestido Ralph Lauren · 18m","categoria":"vestido","revisar":"condicion"},"BL-0142":{"nombre":"Sueter VENDIDO Paz Rodríguez · 12m","categoria":"sueter/chamarra","revisar":"condicion"},"BL-0140":{"nombre":"Vestido Petit amie · 3m","categoria":"vestido","condicion":"nuevo con etiqueta","revisar":""},"BL-0138":{"nombre":"Sueter Neck&Neck · 12m","categoria":"sueter/chamarra","revisar":""},"BL-0137":{"condicion":"nuevo con etiqueta","revisar":""},"BL-0135":{"condicion":"seminuevo","revisar":""},"BL-0134":{"nombre":"Sueter Mayoral · 2-3a","categoria":"sueter/chamarra","revisar":""},"BL-0133":{"nombre":"Pijama Mon Caramel · 18m","categoria":"pijama","condicion":"seminuevo","revisar":""},"BL-0132":{"condicion":"seminuevo","revisar":"categoria"},"BL-0127":{"condicion":"seminuevo","revisar":""},"BL-0126":{"condicion":"seminuevo","revisar":""},"BL-0123":{"condicion":"seminuevo","revisar":""},"BL-0122":{"condicion":"seminuevo","revisar":""},"BL-0115":{"condicion":"seminuevo","revisar":""},"BL-0114":{"condicion":"seminuevo","revisar":""},"BL-0113":{"revisar":"categoria, condicion"},"BL-0112":{"condicion":"seminuevo","revisar":""},"BL-0111":{"condicion":"seminuevo","revisar":""},"BL-0110":{"condicion":"seminuevo","revisar":""},"BL-0107":{"condicion":"seminuevo","revisar":""},"BL-0105":{"condicion":"seminuevo","revisar":""},"BL-0104":{"condicion":"seminuevo","revisar":""},"BL-0103":{"revisar":"condicion"},"BL-0102":{"condicion":"seminuevo","revisar":""},"BL-0097":{"condicion":"seminuevo","revisar":""},"BL-0096":{"condicion":"seminuevo","revisar":""},"BL-0095":{"condicion":"seminuevo","revisar":""},"BL-0094":{"revisar":"categoria"},"BL-0093":{"condicion":"seminuevo","revisar":"categoria"},"BL-0090":{"condicion":"seminuevo","revisar":""},"BL-0089":{"precio":"250","revisar":""},"BL-0087":{"condicion":"seminuevo","revisar":""},"BL-0086":{"condicion":"seminuevo","revisar":""},"BL-0084":{"condicion":"seminuevo","precio":"400","revisar":""},"BL-0081":{"condicion":"nuevo con etiqueta","revisar":""},"BL-0080":{"nombre":"Sueter Janie & Jack · 12m","categoria":"sueter/chamarra","condicion":"seminuevo","precio":"200","revisar":""},"BL-0079":{"precio":"1200","revisar":"categoria, condicion"},"BL-0078":{"condicion":"seminuevo","revisar":""},"BL-0077":{"condicion":"seminuevo","revisar":""},"BL-0076":{"condicion":"seminuevo","revisar":""},"BL-0075":{"condicion":"seminuevo","revisar":""},"BL-0074":{"condicion":"seminuevo","revisar":""},"BL-0073":{"condicion":"seminuevo","revisar":""},"BL-0072":{"condicion":"seminuevo","revisar":""},"BL-0071":{"condicion":"seminuevo","revisar":""},"BL-0069":{"condicion":"seminuevo","revisar":""},"BL-0068":{"condicion":"seminuevo","revisar":""},"BL-0067":{"condicion":"seminuevo","revisar":""},"BL-0066":{"nombre":"Sueter Old Navy · 12m","categoria":"sueter/chamarra","condicion":"seminuevo","revisar":""},"BL-0064":{"condicion":"seminuevo","revisar":""},"BL-0063":{"condicion":"seminuevo","precio":"3000","revisar":""},"BL-0062":{"condicion":"seminuevo","revisar":""},"BL-0060":{"condicion":"seminuevo","revisar":""},"BL-0059":{"nombre":"Sueter H&M · 6m","categoria":"sueter/chamarra","condicion":"seminuevo","revisar":""},"BL-0058":{"condicion":"seminuevo","revisar":""},"BL-0057":{"nombre":"Vestido VENDIDO Trotter St. Kids · 12m","categoria":"vestido","revisar":""},"BL-0055":{"condicion":"seminuevo","revisar":"categoria"},"BL-0054":{"nombre":"Mameluco The Beaufort Bonnet Company · 18m","categoria":"mameluco","marca":"The Beaufort Bonnet Company","revisar":""},"BL-0053":{"revisar":"condicion"},"BL-0051":{"condicion":"seminuevo","revisar":""},"BL-0049":{"marca":"Baby Fashion Portugal"},"BL-0047":{"revisar":"condicion"},"BL-0043":{"condicion":"nuevo con etiqueta","revisar":""},"BL-0042":{"revisar":"categoria, condicion"},"BL-0040":{"revisar":"categoria, condicion"},"BL-0039":{"nombre":"Mameluco Mayoral · 12m","categoria":"mameluco","revisar":""},"BL-0038":{"revisar":"condicion"},"BL-0036":{"nombre":"Mameluco Carter's · 6m","categoria":"mameluco","revisar":"condicion"},"BL-0035":{"nombre":"Mameluco Little Me · 6m","categoria":"mameluco","revisar":"condicion"},"BL-0034":{"nombre":"Mameluco Little Me · 6m","categoria":"mameluco","revisar":"condicion"},"BL-0033":{"condicion":"seminuevo","revisar":""},"BL-0031":{"nombre":"Sueter VENDIDO Hecho a mano · 12m","categoria":"sueter/chamarra","revisar":"condicion"},"BL-0030":{"nombre":"Mameluco VENDIDO Busy Bees Smocked · 12m","categoria":"mameluco","condicion":"seminuevo","revisar":""},"BL-0023":{"nombre":"Sueter Gocco · 12m","categoria":"sueter/chamarra","condicion":"seminuevo","revisar":""},"BL-0021":{"nombre":"Pijama UNIQLO · 6m","revisar":""},"BL-0020":{"nombre":"Sueter Carter's · 9m","categoria":"sueter/chamarra","revisar":""},"BL-0019":{"condicion":"seminuevo","revisar":""},"BL-0017":{"nombre":"Pijama Chicco · 6m","revisar":"condicion"},"BL-0016":{"condicion":"seminuevo","revisar":""},"BL-0014":{"condicion":"seminuevo","revisar":""},"BL-0011":{"condicion":"seminuevo","revisar":""},"BL-0007":{"precio":"750","revisar":""},"BL-0005":{"precio":"1600","marca":"Tartine et chocolat","revisar":""},"BL-0001":{"precio":"150","revisar":""}}};
