/**
 * AutoPartes Inventario - Backend API con Google Sheets
 */

const SHEET_PRODUCTS = 'Productos';
const SHEET_MOVEMENTS = 'Movimientos';
const SHEET_CLIENTS = 'Clientes';
const SHEET_SUPPLIERS = 'Proveedores';
const SHEET_ORDERS = 'Pedidos';

const LOW_STOCK_THRESHOLD = 5;

const ORDERS_HEADERS = [
  'ID', 'FECHA_PEDIDO', 'NUM_PEDIDO', 'PROVEEDOR', 'CODIGO', 'ALLIANCE', 
  'DESCRIPCION', 'MODELO', 'PRECIO_COMPRA', 'CANTIDAD_PEDIDO', 
  'NUM_FACTURA', 'CANTIDAD_EMBARCADA', 'CANTIDAD_PENDIENTE', 'ESTATUS'
];

function doGet(e) {
  const action = e.parameter ? e.parameter.action : '';
  let result = {};

  try {
    if (action === 'getProducts') result = getProducts();
    else if (action === 'getDashboardMetrics') result = getDashboardMetrics();
    else if (action === 'getOrders') result = getOrders();
    else result = { error: 'Acción GET no válida' };
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = {};
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;

    if (action === 'saveBatchOrders') result = saveBatchOrders(payload);
    else if (action === 'deleteOrder') result = deleteOrder(payload);
    else result = { error: 'Acción POST no válida' };
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  return sheet;
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    let cleanHeader = String(header).trim().toUpperCase();
    let value = row[index];
    
    // Mapeo flexible para columna ESTATUS / ESTATUS LOGÍSTICO
    if (cleanHeader === 'ESTATUS LOGÍSTICO' || cleanHeader === 'ESTATUS LOGISTICO' || cleanHeader === 'ESTATUS') {
      obj['ESTATUS'] = value !== undefined && value !== null ? String(value).trim() : 'Solicitado';
    } else {
      obj[cleanHeader] = value !== undefined && value !== null ? String(value).trim() : '';
    }
  });
  return obj;
}

function getOrders() {
  try {
    const sheet = getSheet(SHEET_ORDERS, ORDERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => rowToObject(data[0], row)).reverse();
  } catch (error) {
    throw new Error('No se pudieron cargar los pedidos: ' + error.message);
  }
}

function saveBatchOrders(payload) {
  try {
    if (!payload || !payload.items || payload.items.length === 0) {
      throw new Error('Debe agregar al menos un producto al pedido.');
    }

    const sheet = getSheet(SHEET_ORDERS, ORDERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const idIndex = headers.indexOf('ID');

    if (payload.batchId && idIndex !== -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idIndex]).trim() === String(payload.batchId).trim()) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    const newBatchId = payload.batchId || ('ORD-' + Date.now() + '-' + Math.floor(Math.random() * 1000));
    const newRows = [];

    payload.items.forEach(item => {
      const cantPedido = parseInt(item.cantidadPedido, 10) || 0;
      const cantEmbarcada = parseInt(item.cantidadEmbarcada, 10) || 0;
      const cantPendiente = Math.max(0, cantPedido - cantEmbarcada);

      const orderObj = {
        'ID': newBatchId,
        'FECHA_PEDIDO': payload.fechaPedido,
        'NUM_PEDIDO': payload.numPedido,
        'PROVEEDOR': payload.proveedor,
        'CODIGO': item.codigo,
        'ALLIANCE': item.alliance,
        'DESCRIPCION': item.descripcion,
        'MODELO': item.modelo,
        'PRECIO_COMPRA': parseFloat(item.precioCompra) || 0,
        'CANTIDAD_PEDIDO': cantPedido,
        'NUM_FACTURA': item.numFactura || '',
        'CANTIDAD_EMBARCADA': cantEmbarcada,
        'CANTIDAD_PENDIENTE': cantPendiente,
        'ESTATUS': item.estatus || 'Solicitado'
      };

      const row = headers.map(header => {
        if (header === 'ESTATUS' || header === 'ESTATUS LOGÍSTICO') return orderObj.ESTATUS;
        return orderObj[header] !== undefined ? orderObj[header] : '';
      });
      newRows.push(row);
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    }

    return { success: true, message: `Se guardó correctamente el pedido #${payload.numPedido}` };
  } catch (error) {
    throw new Error('Error al guardar el pedido: ' + error.message);
  }
}
