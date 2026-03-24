const { getWorkbook, saveWorkbook } = require('./excel');

const HOJA_PEDIDOS = 'Pedidos';
const HOJA_PEDIDOS_LEGACY = 'Productos';
const HOJA_CUSTODIAS = 'Movimientos';

// La hoja historica era "Productos". Se migra automaticamente a "Pedidos".

const TIPO_REGISTRO_CUSTODIA = 'CUSTODIA';

const COLUMNA_ID_PEDIDO = 1;
const COLUMNA_FECHA_PEDIDO = 2;
const COLUMNA_CLIENTE_PEDIDO = 3;
const COLUMNA_LUGAR_PEDIDO = 4;
const COLUMNA_DESTINO_PEDIDO = 5;
const COLUMNA_TIPO_CUSTODIA_PEDIDO = 6;
const COLUMNA_GUIA_PEDIDO = 7;
const COLUMNA_JEFE_PATRULLA_PEDIDO = 8;
const COLUMNA_CUSTODIO2_PEDIDO = 9;
const COLUMNA_CUSTODIO3_PEDIDO = 10;
const COLUMNA_PLACA_PEDIDO = 11;
const COLUMNA_HORA_PEDIDO = 12;
const COLUMNA_ESTADO_PEDIDO = 13;

const CABECERAS_PEDIDOS = [
  'ID Pedido',
  'Fecha',
  'Cliente',
  'Lugar',
  'Destino',
  'Tipo de Custodia',
  '# Guía',
  'Jefe Patrulla',
  'Custodio 2',
  'Custodio 3',
  'Placa Vehículo',
  'Hora',
  'Estado'
];

function esEstadoPedido(valor) {
  const estado = normalizarTexto(valor).toUpperCase();
  return estado === 'PENDIENTE' || estado === 'EN PROCESO' || estado === 'FINALIZADO';
}

function getNombreHojaCustodias() {
  if (typeof HOJA_CUSTODIAS === 'string' && HOJA_CUSTODIAS.trim()) {
    return HOJA_CUSTODIAS;
  }
  return 'Movimientos';
}

// Normaliza cualquier dato de entrada a texto limpio.
function normalizarTexto(valor) {
  return (valor || '').toString().trim();
}

function formatearFecha(valor) {
  if (!valor) return '';
  const fecha = valor instanceof Date ? valor : new Date(valor);
  if (Number.isNaN(fecha.getTime())) return '';
  return fecha.toISOString().slice(0, 10);
}

function obtenerHoraPedidoDesdeFila(row) {
  const hora = normalizarTexto(row.getCell(COLUMNA_HORA_PEDIDO).value);
  return esEstadoPedido(hora) ? '' : hora;
}

function obtenerEstadoPedidoDesdeFila(row) {
  const estadoActual = normalizarTexto(row.getCell(COLUMNA_ESTADO_PEDIDO).value).toUpperCase();
  if (estadoActual) {
    return estadoActual;
  }

  const estadoLegacy = normalizarTexto(row.getCell(COLUMNA_HORA_PEDIDO).value).toUpperCase();
  if (esEstadoPedido(estadoLegacy)) {
    return estadoLegacy;
  }

  return 'PENDIENTE';
}

function mapearFilaLegacyPedido(row) {
  const estadoCol13 = normalizarTexto(row.getCell(13).value).toUpperCase();
  const valorCol12 = normalizarTexto(row.getCell(12).value);
  const hora = esEstadoPedido(valorCol12) ? '' : valorCol12;
  const estado = estadoCol13 || (esEstadoPedido(valorCol12) ? valorCol12.toUpperCase() : 'PENDIENTE');

  return {
    pedidoId: normalizarTexto(row.getCell(11).value),
    fecha: formatearFecha(row.getCell(6).value),
    codigo: normalizarTexto(row.getCell(1).value).toUpperCase(),
    nombre: normalizarTexto(row.getCell(2).value),
    unidad: normalizarTexto(row.getCell(3).value) || 'Unidades',
    grupo: normalizarTexto(row.getCell(4).value) || 'General',
    guia: normalizarTexto(row.getCell(5).value),
    jefePatrulla: normalizarTexto(row.getCell(7).value),
    custodio2: normalizarTexto(row.getCell(8).value),
    custodio3: normalizarTexto(row.getCell(9).value),
    placaVehiculo: normalizarTexto(row.getCell(10).value),
    hora,
    estado
  };
}

function crearFilaPedidoEnOrdenInventario(pedido, rowNumber) {
  return [
    normalizarTexto(pedido.pedidoId) || String(rowNumber),
    formatearFecha(pedido.fecha) || formatearFecha(new Date()),
    normalizarTexto(pedido.codigo).toUpperCase(),
    normalizarTexto(pedido.nombre),
    normalizarTexto(pedido.unidad) || 'Unidades',
    normalizarTexto(pedido.grupo) || 'General',
    normalizarTexto(pedido.guia),
    normalizarTexto(pedido.jefePatrulla),
    normalizarTexto(pedido.custodio2),
    normalizarTexto(pedido.custodio3),
    normalizarTexto(pedido.placaVehiculo),
    normalizarTexto(pedido.hora),
    normalizarTexto(pedido.estado).toUpperCase() || 'PENDIENTE'
  ];
}

function migrarSheetLegacyAInventario(workbook, sourceSheet) {
  const tempName = '__tmp_pedidos_migracion__';
  const previo = workbook.getWorksheet(tempName);
  if (previo) {
    workbook.removeWorksheet(previo.id);
  }

  const tempSheet = workbook.addWorksheet(tempName);
  tempSheet.addRow(CABECERAS_PEDIDOS);

  let destinoFila = 2;
  for (let i = 2; i <= sourceSheet.rowCount; i += 1) {
    const legacy = mapearFilaLegacyPedido(sourceSheet.getRow(i));
    const tieneDatos = legacy.codigo || legacy.nombre || legacy.pedidoId;
    if (!tieneDatos) continue;

    tempSheet.addRow(crearFilaPedidoEnOrdenInventario(legacy, destinoFila));
    destinoFila += 1;
  }

  workbook.removeWorksheet(sourceSheet.id);
  tempSheet.name = HOJA_PEDIDOS;

  return tempSheet;
}

function hojaPedidosEnFormatoInventario(sheet) {
  const cabecera = normalizarTexto(sheet.getRow(1).getCell(1).value).toUpperCase();
  return cabecera === 'ID PEDIDO';
}

// Sincroniza el ID Pedido con el número real de fila en Excel.
function sincronizarIdsPedidoPorFila(sheet) {
  let huboCambios = false;

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const idEsperado = String(i);
    const idActual = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);

    if (idActual !== idEsperado) {
      row.getCell(COLUMNA_ID_PEDIDO).value = idEsperado;
      huboCambios = true;
    }
  }

  return huboCambios;
}

// Ubica una fila de pedido por ID Pedido (preferido) o por código legacy.
function buscarFilaPedido(sheet, identificador) {
  const buscado = normalizarTexto(identificador);
  if (!buscado) return null;

  let encontrado = null;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || encontrado) return;
    const pedidoId = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);
    const codigo = normalizarTexto(row.getCell(COLUMNA_CLIENTE_PEDIDO).value).toUpperCase();
    if (pedidoId === buscado || codigo === buscado.toUpperCase()) {
      encontrado = row;
    }
  });

  return encontrado;
}

// Asegura que exista la hoja de pedidos con su cabecera.
function asegurarHojaPedidos(workbook) {
  let sheet = workbook.getWorksheet(HOJA_PEDIDOS);
  if (!sheet) {
    const legacy = workbook.getWorksheet(HOJA_PEDIDOS_LEGACY);
    if (legacy) {
      sheet = migrarSheetLegacyAInventario(workbook, legacy);
    }
  }

  if (!sheet) {
    sheet = workbook.addWorksheet(HOJA_PEDIDOS);
    sheet.addRow(CABECERAS_PEDIDOS);
    return sheet;
  }

  if (!hojaPedidosEnFormatoInventario(sheet)) {
    sheet = migrarSheetLegacyAInventario(workbook, sheet);
  }

  const headerRow = sheet.getRow(1);
  CABECERAS_PEDIDOS.forEach((title, index) => {
    const col = index + 1;
    const value = normalizarTexto(headerRow.getCell(col).value);
    if (!value || value.toUpperCase() !== title.toUpperCase()) {
      headerRow.getCell(col).value = title;
    }
  });

  return sheet;
}

// Asegura que exista la hoja historica de custodias con su cabecera.
function asegurarHojaCustodias(workbook) {
  const nombreHoja = getNombreHojaCustodias();
  let sheet = workbook.getWorksheet(nombreHoja);
  if (!sheet) {
    sheet = workbook.addWorksheet(nombreHoja);
    sheet.addRow(['ID Pedido', 'Fecha', 'Tipo Registro', 'Cantidad', 'Observaciones', 'Valor Calculado']);
  }
  return sheet;
}

// Asegura estructura de la hoja CUSTODIOS para gestion administrativa.
function asegurarHojaCatalogoCustodios(workbook) {
  let sheet = workbook.getWorksheet('CUSTODIOS');
  if (!sheet) {
    sheet = workbook.addWorksheet('CUSTODIOS');
    sheet.addRow(['Nombres', 'Apellidos', 'CI', 'Telefono']);
    return sheet;
  }

  const headerRow = sheet.getRow(1);
  const expectedHeaders = ['Nombres', 'Apellidos', 'CI', 'Telefono'];
  expectedHeaders.forEach((title, index) => {
    const col = index + 1;
    const value = normalizarTexto(headerRow.getCell(col).value);
    if (!value) {
      headerRow.getCell(col).value = title;
    }
  });

  return sheet;
}

function asegurarHojaVehiculos(workbook) {
  let sheet = workbook.getWorksheet('vehiculos') || workbook.getWorksheet('VEHICULOS');
  if (!sheet) {
    sheet = workbook.addWorksheet('vehiculos');
    sheet.addRow(['Placa Vehículo']);
    return sheet;
  }

  const headerRow = sheet.getRow(1);
  const value = normalizarTexto(headerRow.getCell(1).value);
  if (!value) {
    headerRow.getCell(1).value = 'Placa Vehículo';
  }

  return sheet;
}

function mapearFilaCustodio(row) {
  const id = String(row.number);
  const nombres = normalizarTexto(row.getCell(1).value);
  const apellidos = normalizarTexto(row.getCell(2).value);
  const ci = normalizarTexto(row.getCell(3).value);
  const telefono = normalizarTexto(row.getCell(4).value);

  return { id, nombres, apellidos, ci, telefono };
}

function nombreCompletoCustodio(custodio) {
  const fullName = `${normalizarTexto(custodio.nombres)} ${normalizarTexto(custodio.apellidos)}`.trim();
  return fullName || normalizarTexto(custodio.nombres);
}

function obtenerAliasesCustodio(custodio) {
  const nombres = normalizarTexto(custodio.nombres);
  const apellidos = normalizarTexto(custodio.apellidos);
  const full = `${nombres} ${apellidos}`.trim();

  return Array.from(new Set([
    nombres,
    apellidos,
    full,
    nombreCompletoCustodio(custodio)
  ].map((valor) => normalizarTexto(valor)).filter(Boolean)));
}

// Construye un indice flexible para convertir nombres historicos en el nombre
// completo vigente del custodio. Asi, al refrescar Inventario/Buscar/Reportes,
// se muestran los nombres actualizados desde la hoja CUSTODIOS.
async function construirIndiceCustodios(workbook) {
  const custodios = await listarCustodios();
  const indice = new Map();

  custodios.forEach((custodio) => {
    const nombreActual = nombreCompletoCustodio(custodio);
    const nombres = normalizarTexto(custodio.nombres);
    const apellidos = normalizarTexto(custodio.apellidos);
    const full = `${nombres} ${apellidos}`.trim();

    [nombreActual, nombres, apellidos, full]
      .map((valor) => normalizarTexto(valor).toLowerCase())
      .filter(Boolean)
      .forEach((alias) => {
        indice.set(alias, nombreActual);
      });
  });

  return indice;
}

function resolverNombreCustodio(valor, indiceCustodios) {
  const original = normalizarTexto(valor);
  if (!original) return '';
  const resolved = indiceCustodios.get(original.toLowerCase());
  return resolved || original;
}

function normalizarDigitos(valor) {
  return normalizarTexto(valor).replace(/\D/g, '');
}

function normalizarPlaca(valor) {
  return normalizarTexto(valor).toUpperCase();
}

function buscarFilaCustodio(sheet, identificador) {
  const buscadoTexto = normalizarTexto(identificador);
  const buscadoDigitos = normalizarDigitos(identificador);
  let fila = null;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || fila) return;
    const rowId = String(row.number);
    const rowCi = normalizarDigitos(row.getCell(3).value);

    if (rowId === buscadoTexto || (buscadoDigitos && rowCi === buscadoDigitos)) {
      fila = row;
    }
  });

  return fila;
}

function actualizarReferenciasCustodioEnPedidos(sheet, aliasesViejos, nombreNuevo) {
  if (!sheet || !aliasesViejos.length || !nombreNuevo) {
    return false;
  }

  const aliases = new Set(aliasesViejos.map((valor) => valor.toLowerCase()));
  let huboCambios = false;

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);

    [COLUMNA_JEFE_PATRULLA_PEDIDO, COLUMNA_CUSTODIO2_PEDIDO, COLUMNA_CUSTODIO3_PEDIDO].forEach((columna) => {
      const valorActual = normalizarTexto(row.getCell(columna).value);
      if (!valorActual) return;

      if (aliases.has(valorActual.toLowerCase())) {
        row.getCell(columna).value = nombreNuevo;
        huboCambios = true;
      }
    });
  }

  return huboCambios;
}

function mapearFilaVehiculo(row) {
  return {
    id: String(row.number),
    placa: normalizarPlaca(row.getCell(1).value)
  };
}

function buscarFilaVehiculo(sheet, identificador) {
  const buscadoTexto = normalizarTexto(identificador);
  const buscadoPlaca = normalizarPlaca(identificador);
  let fila = null;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || fila) return;
    const rowId = String(row.number);
    const rowPlaca = normalizarPlaca(row.getCell(1).value);
    if (rowId === buscadoTexto || (buscadoPlaca && rowPlaca === buscadoPlaca)) {
      fila = row;
    }
  });

  return fila;
}

function actualizarReferenciasVehiculoEnPedidos(sheet, placaAnterior, placaNueva) {
  if (!sheet || !placaAnterior || !placaNueva) {
    return false;
  }

  const anterior = normalizarPlaca(placaAnterior);
  let huboCambios = false;

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const actual = normalizarPlaca(row.getCell(COLUMNA_PLACA_PEDIDO).value);
    if (actual && actual === anterior) {
      row.getCell(COLUMNA_PLACA_PEDIDO).value = placaNueva;
      huboCambios = true;
    }
  }

  return huboCambios;
}

function actualizarReferenciasVehiculoEnCustodias(sheet, placaAnterior, placaNueva) {
  if (!sheet || !placaAnterior || !placaNueva) {
    return false;
  }

  const anterior = normalizarPlaca(placaAnterior);
  let huboCambios = false;

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const observaciones = normalizarTexto(row.getCell(5).value);
    if (!observaciones) continue;

    const reemplazadas = observaciones.replace(/Placa:\s*([^|]+)/i, (match, placaActual) => {
      if (normalizarPlaca(placaActual) === anterior) {
        huboCambios = true;
        return `Placa: ${placaNueva}`;
      }
      return match;
    });

    if (reemplazadas !== observaciones) {
      row.getCell(5).value = reemplazadas;
    }
  }

  return huboCambios;
}

// Lista custodios para la pantalla de administracion.
async function listarCustodios() {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaCatalogoCustodios(workbook);
  const custodios = [];

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const custodio = mapearFilaCustodio(row);
    const hasData = custodio.nombres || custodio.apellidos || custodio.ci || custodio.telefono;
    if (!hasData) continue;

    custodios.push(custodio);
  }

  return custodios;
}

// Crea un custodio validando CI unica.
async function crearCustodio(data) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaCatalogoCustodios(workbook);

  const nombres = normalizarTexto(data.nombres);
  const apellidos = normalizarTexto(data.apellidos);
  const ci = normalizarDigitos(data.ci);
  const telefono = normalizarDigitos(data.telefono);

  if (!nombres || !apellidos || !ci || !telefono) {
    return 'Nombres, apellidos, CI y telefono son obligatorios';
  }

  if (ci.length !== 10) {
    return 'La CI debe tener exactamente 10 digitos';
  }

  if (telefono.length !== 10) {
    return 'El telefono debe tener exactamente 10 digitos';
  }

  let existe = false;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || existe) return;
    const ciActual = normalizarTexto(row.getCell(3).value);
    if (ciActual === ci) existe = true;
  });

  if (existe) {
    return 'Ya existe un custodio con esa CI';
  }

  sheet.addRow([nombres, apellidos, ci, telefono]);
  await saveWorkbook(workbook);
  return 'Custodio creado correctamente';
}

// Actualiza custodio por identificador interno o CI.
async function actualizarCustodio(identificador, data) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaCatalogoCustodios(workbook);
  const sheetPedidos = asegurarHojaPedidos(workbook);
  const identificadorBuscado = normalizarTexto(identificador);

  const nombres = normalizarTexto(data.nombres);
  const apellidos = normalizarTexto(data.apellidos);
  const ciNueva = normalizarDigitos(data.ci);
  const telefono = normalizarDigitos(data.telefono);

  if (!nombres || !apellidos || !ciNueva || !telefono) {
    return 'Nombres, apellidos, CI y telefono son obligatorios';
  }

  if (ciNueva.length !== 10) {
    return 'La CI debe tener exactamente 10 digitos';
  }

  if (telefono.length !== 10) {
    return 'El telefono debe tener exactamente 10 digitos';
  }

  const targetRow = buscarFilaCustodio(sheet, identificadorBuscado);

  if (!targetRow) {
    return 'Custodio no encontrado';
  }

  const custodioAnterior = mapearFilaCustodio(targetRow);
  const aliasesAnteriores = obtenerAliasesCustodio(custodioAnterior);
  const nombreNuevo = nombreCompletoCustodio({ nombres, apellidos });

  const ciBuscada = normalizarDigitos(targetRow.getCell(3).value);

  if (ciNueva !== ciBuscada) {
    let duplicada = false;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1 || duplicada || row.number === targetRow.number) return;
      const ci = normalizarTexto(row.getCell(3).value);
      if (ci === ciNueva) duplicada = true;
    });
    if (duplicada) {
      return 'La CI nueva ya existe en otro custodio';
    }
  }

  targetRow.getCell(1).value = nombres;
  targetRow.getCell(2).value = apellidos;
  targetRow.getCell(3).value = ciNueva;
  targetRow.getCell(4).value = telefono;

  actualizarReferenciasCustodioEnPedidos(sheetPedidos, aliasesAnteriores, nombreNuevo);

  await saveWorkbook(workbook);
  return 'Custodio actualizado correctamente';
}

// Elimina custodio por identificador interno o CI.
async function eliminarCustodio(identificador) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaCatalogoCustodios(workbook);
  const fila = buscarFilaCustodio(sheet, identificador);
  const rowNumberToDelete = fila ? fila.number : null;

  if (!rowNumberToDelete) {
    return 'Custodio no encontrado';
  }

  sheet.spliceRows(rowNumberToDelete, 1);
  await saveWorkbook(workbook);
  return 'Custodio eliminado correctamente';
}

async function listarVehiculos() {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaVehiculos(workbook);
  const vehiculos = [];

  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const vehiculo = mapearFilaVehiculo(row);
    if (!vehiculo.placa) continue;
    vehiculos.push(vehiculo);
  }

  return vehiculos;
}

async function crearVehiculo(data) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaVehiculos(workbook);
  const placa = normalizarPlaca(data.placa);

  if (!placa) {
    return 'La placa es obligatoria';
  }

  let existe = false;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || existe) return;
    if (normalizarPlaca(row.getCell(1).value) === placa) existe = true;
  });

  if (existe) {
    return 'Ya existe un vehículo con esa placa';
  }

  sheet.addRow([placa]);
  await saveWorkbook(workbook);
  return 'Vehículo creado correctamente';
}

async function actualizarVehiculo(identificador, data) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaVehiculos(workbook);
  const sheetPedidos = asegurarHojaPedidos(workbook);
  const sheetCustodias = workbook.getWorksheet(getNombreHojaCustodias());
  const fila = buscarFilaVehiculo(sheet, identificador);
  const placaNueva = normalizarPlaca(data.placa);

  if (!fila) {
    return 'Vehículo no encontrado';
  }

  if (!placaNueva) {
    return 'La placa es obligatoria';
  }

  const placaAnterior = normalizarPlaca(fila.getCell(1).value);
  if (placaNueva !== placaAnterior) {
    let duplicada = false;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1 || duplicada || row.number === fila.number) return;
      if (normalizarPlaca(row.getCell(1).value) === placaNueva) duplicada = true;
    });
    if (duplicada) {
      return 'La placa nueva ya existe en otro vehículo';
    }
  }

  fila.getCell(1).value = placaNueva;
  actualizarReferenciasVehiculoEnPedidos(sheetPedidos, placaAnterior, placaNueva);
  actualizarReferenciasVehiculoEnCustodias(sheetCustodias, placaAnterior, placaNueva);

  await saveWorkbook(workbook);
  return 'Vehículo actualizado correctamente';
}

async function eliminarVehiculo(identificador) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaVehiculos(workbook);
  const fila = buscarFilaVehiculo(sheet, identificador);

  if (!fila) {
    return 'Vehículo no encontrado';
  }

  sheet.spliceRows(fila.number, 1);
  await saveWorkbook(workbook);
  return 'Vehículo eliminado correctamente';
}

// Registra un pedido nuevo. Se permite cliente repetido por requerimiento.
async function registrarPedido(producto) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaPedidos(workbook);
  const codigoNuevo = normalizarTexto(producto.codigo).toUpperCase();
  const fechaPedido = formatearFecha(producto.fecha) || formatearFecha(new Date());
  const horaPedido = normalizarTexto(producto.hora);
  const estadoPedido = normalizarTexto(producto.estado).toUpperCase() || 'PENDIENTE';
  const guiaPedido = normalizarTexto(producto.guia ?? producto.stockMin);
  sincronizarIdsPedidoPorFila(sheet);

  if (!codigoNuevo || !normalizarTexto(producto.nombre)) {
    return 'Código y nombre son obligatorios';
  }

  sheet.addRow([
    String(sheet.rowCount + 1),
    fechaPedido,
    codigoNuevo,
    normalizarTexto(producto.nombre),
    normalizarTexto(producto.unidad) || 'Unidades',
    normalizarTexto(producto.grupo) || 'General',
    guiaPedido,
    '',
    '',
    '',
    '',
    horaPedido,
    estadoPedido
  ]);

  // Corrige IDs por seguridad para que siempre coincidan con la fila real.
  sincronizarIdsPedidoPorFila(sheet);
  const pedidoId = String(sheet.rowCount);

  await saveWorkbook(workbook);
  return {
    mensaje: `Pedido registrado correctamente. ID Pedido: ${pedidoId}`,
    pedidoId
  };
}

// Actualiza los datos visibles de un pedido existente.
async function actualizarPedido(codigoActual, productoActualizado) {
  const workbook = await getWorkbook();
  const sheetProd = asegurarHojaPedidos(workbook);

  if (!sheetProd) {
    return 'No existe la hoja de pedidos';
  }

  const identificadorBuscado = normalizarTexto(codigoActual);
  const nuevoCodigo = normalizarTexto(productoActualizado.codigo).toUpperCase();
  const nuevoNombre = normalizarTexto(productoActualizado.nombre);
  const nuevaUnidad = normalizarTexto(productoActualizado.unidad) || 'Unidades';
  const nuevoGrupo = normalizarTexto(productoActualizado.grupo) || 'General';
  const nuevaGuia = normalizarTexto(productoActualizado.guia ?? productoActualizado.stockMin);
  const nuevaFecha = formatearFecha(productoActualizado.fecha);
  const nuevaHora = normalizarTexto(productoActualizado.hora);
  const nuevoEstado = normalizarTexto(productoActualizado.estado).toUpperCase() || 'PENDIENTE';

  if (!nuevoCodigo || !nuevoNombre) {
    return 'Código y nombre son obligatorios';
  }

  const filaProducto = buscarFilaPedido(sheetProd, identificadorBuscado);

  if (!filaProducto) {
    return 'Pedido no encontrado';
  }

  filaProducto.getCell(COLUMNA_CLIENTE_PEDIDO).value = nuevoCodigo;
  filaProducto.getCell(COLUMNA_LUGAR_PEDIDO).value = nuevoNombre;
  filaProducto.getCell(COLUMNA_DESTINO_PEDIDO).value = nuevaUnidad;
  filaProducto.getCell(COLUMNA_TIPO_CUSTODIA_PEDIDO).value = nuevoGrupo;
  filaProducto.getCell(COLUMNA_GUIA_PEDIDO).value = nuevaGuia;
  filaProducto.getCell(COLUMNA_FECHA_PEDIDO).value = nuevaFecha || formatearFecha(filaProducto.getCell(COLUMNA_FECHA_PEDIDO).value) || formatearFecha(new Date());
  filaProducto.getCell(COLUMNA_HORA_PEDIDO).value = nuevaHora;
  filaProducto.getCell(COLUMNA_ESTADO_PEDIDO).value = nuevoEstado;

  sincronizarIdsPedidoPorFila(sheetProd);

  await saveWorkbook(workbook);
  return 'Pedido actualizado correctamente';
}

// Elimina un pedido y sus registros de custodia asociados por ID Pedido.
async function eliminarPedido(codigo) {
  const workbook = await getWorkbook();
  const sheetProd = asegurarHojaPedidos(workbook);

  if (!sheetProd) {
    return 'No existe la hoja de pedidos';
  }

  const filaPedido = buscarFilaPedido(sheetProd, codigo);
  if (!filaPedido) {
    return 'Pedido no encontrado';
  }

  const filaEliminar = filaPedido.number;
  const pedidoIdEliminado = normalizarTexto(filaPedido.getCell(COLUMNA_ID_PEDIDO).value);
  sheetProd.spliceRows(filaEliminar, 1);
  sincronizarIdsPedidoPorFila(sheetProd);

  const sheetCustodias = workbook.getWorksheet(getNombreHojaCustodias());
  if (sheetCustodias) {
    // Se recorre de abajo hacia arriba para eliminar filas sin desordenar índices.
    for (let i = sheetCustodias.rowCount; i >= 2; i -= 1) {
      const row = sheetCustodias.getRow(i);
      const pedidoIdCustodia = normalizarTexto(row.getCell(1).value);
      if (pedidoIdCustodia === pedidoIdEliminado) {
        sheetCustodias.spliceRows(i, 1);
      }
    }
  }

  await saveWorkbook(workbook);
  return 'Pedido eliminado correctamente';
}

// Registra una custodia y la asocia al pedido correspondiente.
async function registrarCustodia(custodia) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaCustodias(workbook);
  const sheetProd = asegurarHojaPedidos(workbook);

  const pedidoId = normalizarTexto(custodia.pedidoId || custodia.codigo);
  const tipo = normalizarTexto(custodia.tipo).toUpperCase();
  const cantidad = parseFloat(custodia.cantidad) || 0;

  if (!pedidoId || !tipo || cantidad <= 0) {
    return 'Datos de custodia incompletos';
  }

  if (tipo !== TIPO_REGISTRO_CUSTODIA) {
    return 'Tipo de registro inválido';
  }

  const jefePatrulla = normalizarTexto(custodia.jefePatrulla);
  const custodio2 = normalizarTexto(custodia.custodio2);
  const custodio3 = normalizarTexto(custodia.custodio3);
  const placa = normalizarTexto(custodia.placa);

  if (!pedidoId || !jefePatrulla || !custodio2 || !custodio3 || !placa) {
    return 'Faltan datos de custodia para asociar al ID del pedido';
  }

  sincronizarIdsPedidoPorFila(sheetProd);

  let pedidoRow = null;
  sheetProd.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || pedidoRow) return;
    const idPedido = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);
    if (idPedido === pedidoId) {
      pedidoRow = row;
    }
  });

  if (!pedidoRow) {
    return 'No existe un pedido con ese ID Pedido';
  }

  // Guarda datos de custodia al lado de los datos del pedido.
  pedidoRow.getCell(COLUMNA_JEFE_PATRULLA_PEDIDO).value = jefePatrulla;
  pedidoRow.getCell(COLUMNA_CUSTODIO2_PEDIDO).value = custodio2;
  pedidoRow.getCell(COLUMNA_CUSTODIO3_PEDIDO).value = custodio3;
  pedidoRow.getCell(COLUMNA_PLACA_PEDIDO).value = placa;

  sheet.addRow([
    pedidoId,
    custodia.fecha ? new Date(custodia.fecha) : new Date(),
    tipo,
    cantidad,
    normalizarTexto(custodia.observaciones),
    0
  ]);

  await saveWorkbook(workbook);
  return 'Custodia registrada correctamente';
}

// Obtiene listas operativas para custodias desde hojas CUSTODIOS y vehiculos.
async function obtenerDatosCustodia() {
  const workbook = await getWorkbook();
  const custodiosSheet = asegurarHojaCatalogoCustodios(workbook);
  const vehiculosSheet = asegurarHojaVehiculos(workbook);
  const pedidosSheet = asegurarHojaPedidos(workbook);

  const custodiosSet = new Set();
  const vehiculosSet = new Set();
  const pedidoIdsSet = new Set();

  if (custodiosSheet) {
    const custodios = await listarCustodios();
    custodios.forEach((custodio) => {
      const nombre = nombreCompletoCustodio(custodio);
      if (nombre) custodiosSet.add(nombre);
    });
  }

  if (vehiculosSheet) {
    const vehiculos = await listarVehiculos();
    vehiculos.forEach((vehiculo) => {
      if (vehiculo.placa) vehiculosSet.add(vehiculo.placa);
    });
  }

  if (pedidosSheet) {
    const huboCambios = sincronizarIdsPedidoPorFila(pedidosSheet);

    pedidosSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const pedidoId = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);
      if (pedidoId) pedidoIdsSet.add(pedidoId);
    });

    if (huboCambios) {
      await saveWorkbook(workbook);
    }
  }

  return {
    custodios: Array.from(custodiosSet),
    vehiculos: Array.from(vehiculosSet),
    pedidoIds: Array.from(pedidoIdsSet)
  };
}

// Busca pedidos por codigo, nombre o grupo.
async function buscarPedido(texto) {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaPedidos(workbook);
  const filtro = normalizarTexto(texto).toLowerCase();
  const indiceCustodios = await construirIndiceCustodios(workbook);

  if (!sheet) return [];

  const resultados = [];
  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const codigo = normalizarTexto(row.getCell(COLUMNA_CLIENTE_PEDIDO).value);
    const pedidoId = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);

    // Omitir filas completamente vacías
    if (!codigo && !pedidoId) continue;

    const nombre = normalizarTexto(row.getCell(COLUMNA_LUGAR_PEDIDO).value);
    const unidad = normalizarTexto(row.getCell(COLUMNA_DESTINO_PEDIDO).value) || 'Unidades';
    const grupo = normalizarTexto(row.getCell(COLUMNA_TIPO_CUSTODIA_PEDIDO).value) || 'General';
    const guia = normalizarTexto(row.getCell(COLUMNA_GUIA_PEDIDO).value);
    const fecha = formatearFecha(row.getCell(COLUMNA_FECHA_PEDIDO).value);
    const hora = obtenerHoraPedidoDesdeFila(row);
    const estado = obtenerEstadoPedidoDesdeFila(row);
    const jefePatrulla = resolverNombreCustodio(row.getCell(COLUMNA_JEFE_PATRULLA_PEDIDO).value, indiceCustodios);
    const custodio2 = resolverNombreCustodio(row.getCell(COLUMNA_CUSTODIO2_PEDIDO).value, indiceCustodios);
    const custodio3 = resolverNombreCustodio(row.getCell(COLUMNA_CUSTODIO3_PEDIDO).value, indiceCustodios);
    const placaVehiculo = normalizarTexto(row.getCell(COLUMNA_PLACA_PEDIDO).value);

    // Filtramos por codigo, nombre, grupo, guia, pedidoId y custodios ya resueltos.
    const coincide = !filtro
      || codigo.toLowerCase().includes(filtro)
      || nombre.toLowerCase().includes(filtro)
      || grupo.toLowerCase().includes(filtro)
      || guia.toLowerCase().includes(filtro)
      || pedidoId.toLowerCase().includes(filtro)
      || jefePatrulla.toLowerCase().includes(filtro)
      || custodio2.toLowerCase().includes(filtro)
      || custodio3.toLowerCase().includes(filtro);

    if (!coincide) continue;

    resultados.push({
      codigo,
      nombre,
      unidad,
      grupo,
      guia,
      fecha,
      hora,
      estado,
      jefePatrulla,
      custodio2,
      custodio3,
      placaVehiculo,
      pedidoId
    });
  }

  return resultados;
}

// Devuelve listas de unidades y grupos para formularios.
async function obtenerListas() {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaPedidos(workbook);

  // Tipos de custodia: deben mantenerse igual a los de pedidos.html y getListasPorDefecto().
  const defaultUnidades = ['Fluvial', 'Interna', 'Externa', 'Hombre en Cabina'];
  const defaultGrupos = ['Fluvial', 'Interna', 'Externa', 'Hombre en Cabina'];

  if (!sheet) {
    return { unidades: defaultUnidades, grupos: defaultGrupos };
  }

  // Conserva opciones base y agrega las personalizadas que ya existan.
  const unidadesSet = new Set(defaultUnidades);
  const gruposSet = new Set(defaultGrupos);

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    unidadesSet.add(normalizarTexto(row.getCell(COLUMNA_DESTINO_PEDIDO).value) || 'Unidades');
    gruposSet.add(normalizarTexto(row.getCell(COLUMNA_TIPO_CUSTODIA_PEDIDO).value) || 'General');
  });

  const unidades = Array.from(unidadesSet).filter(Boolean);
  const grupos = Array.from(gruposSet).filter(Boolean);

  return { unidades, grupos };
}

// Obtiene los pedidos consolidados para inventario, dashboard y reportes.
async function obtenerPedidos() {
  const workbook = await getWorkbook();
  const sheet = asegurarHojaPedidos(workbook);
  const indiceCustodios = await construirIndiceCustodios(workbook);

  if (!sheet) return [];

  const pedidos = [];
  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const row = sheet.getRow(i);
    const codigo = normalizarTexto(row.getCell(COLUMNA_CLIENTE_PEDIDO).value);
    const pedidoId = normalizarTexto(row.getCell(COLUMNA_ID_PEDIDO).value);
    if (!codigo && !pedidoId) continue;

    const nombre = normalizarTexto(row.getCell(COLUMNA_LUGAR_PEDIDO).value);
    const unidad = normalizarTexto(row.getCell(COLUMNA_DESTINO_PEDIDO).value) || 'Unidades';
    const grupo = normalizarTexto(row.getCell(COLUMNA_TIPO_CUSTODIA_PEDIDO).value) || 'General';
    const guia = normalizarTexto(row.getCell(COLUMNA_GUIA_PEDIDO).value);
    const fecha = formatearFecha(row.getCell(COLUMNA_FECHA_PEDIDO).value);
    const hora = obtenerHoraPedidoDesdeFila(row);
    const estado = obtenerEstadoPedidoDesdeFila(row);
    const jefePatrulla = resolverNombreCustodio(row.getCell(COLUMNA_JEFE_PATRULLA_PEDIDO).value, indiceCustodios);
    const custodio2 = resolverNombreCustodio(row.getCell(COLUMNA_CUSTODIO2_PEDIDO).value, indiceCustodios);
    const custodio3 = resolverNombreCustodio(row.getCell(COLUMNA_CUSTODIO3_PEDIDO).value, indiceCustodios);
    const placaVehiculo = normalizarTexto(row.getCell(COLUMNA_PLACA_PEDIDO).value);

    pedidos.push({
      codigo,
      nombre,
      unidad,
      grupo,
      guia,
      fecha,
      hora,
      estado,
      jefePatrulla,
      custodio2,
      custodio3,
      placaVehiculo,
      pedidoId
    });
  }

  return pedidos;
}

// Verifica y crea las hojas y columnas necesarias en el Excel si no existen.
// Se puede invocar desde la interfaz (boton Configuracion BD) para reparar la BD.
async function inicializarExcel() {
  const workbook = await getWorkbook();
  const resumen = [];
  const hojaCustodias = getNombreHojaCustodias();

  // --- Hoja de Pedidos ---
  const existiaPedidos = Boolean(workbook.getWorksheet(HOJA_PEDIDOS));
  const existiaPedidosLegacy = Boolean(workbook.getWorksheet(HOJA_PEDIDOS_LEGACY));
  const sheetPedidos = asegurarHojaPedidos(workbook);
  if (!existiaPedidos && !existiaPedidosLegacy) {
    resumen.push(`Hoja '${HOJA_PEDIDOS}' creada con sus 13 columnas`);
  } else if (!existiaPedidos && existiaPedidosLegacy) {
    resumen.push(`Hoja '${HOJA_PEDIDOS_LEGACY}' migrada y renombrada a '${HOJA_PEDIDOS}'`);
  } else {
    resumen.push(`Hoja '${HOJA_PEDIDOS}' verificada (${sheetPedidos.rowCount - 1} registros)`);
  }

  // --- Hoja de Custodias ---
  const existiaCustodias = Boolean(workbook.getWorksheet(hojaCustodias));
  const sheetCustodias = asegurarHojaCustodias(workbook);
  if (!existiaCustodias) {
    resumen.push(`Hoja '${hojaCustodias}' creada con sus columnas`);
  } else {
    resumen.push(`Hoja '${hojaCustodias}' verificada (${sheetCustodias.rowCount - 1} registros)`);
  }

  // --- Hoja CUSTODIOS (lista de custodios) ---
  if (!workbook.getWorksheet('CUSTODIOS')) {
    const sh = workbook.addWorksheet('CUSTODIOS');
    sh.addRow(['Nombre Custodio']);
    resumen.push("Hoja 'CUSTODIOS' creada");
  } else {
    resumen.push("Hoja 'CUSTODIOS' ya existe");
  }

  // --- Hoja vehiculos (lista de vehículos) ---
  if (!workbook.getWorksheet('vehiculos') && !workbook.getWorksheet('VEHICULOS')) {
    const sh = workbook.addWorksheet('vehiculos');
    sh.addRow(['Placa Vehículo']);
    resumen.push("Hoja 'vehiculos' creada");
  } else {
    resumen.push("Hoja 'vehiculos' ya existe");
  }

  await saveWorkbook(workbook);

  return 'Configuracion BD completada:\n' + resumen.map(r => '• ' + r).join('\n');
}

module.exports = {
  registrarPedido,
  actualizarPedido,
  eliminarPedido,
  registrarCustodia,
  listarCustodios,
  crearCustodio,
  actualizarCustodio,
  eliminarCustodio,
  listarVehiculos,
  crearVehiculo,
  actualizarVehiculo,
  eliminarVehiculo,
  obtenerDatosCustodia,
  buscarPedido,
  obtenerListas,
  obtenerPedidos,
  inicializarExcel
};