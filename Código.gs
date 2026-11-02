/**
 * AutoPartes Inventario - Backend API con Google Sheets
 * Archivo: Código.gs
 */

const SHEET_PRODUCTS = 'Productos';
const SHEET_MOVEMENTS = 'Movimientos';
const SHEET_CLIENTS = 'Clientes';
const SHEET_SUPPLIERS = 'Proveedores';
const SHEET_ORDERS = 'Pedidos';

const LOW_STOCK_THRESHOLD = 5;

const PRODUCTS_HEADERS = [
  'CODIGO', 'ALLIANCE', 'DESCRIPCION', 'MODELO', 'PRECIO COMPRA ($)',
  'GRUPO', 'PRECIO VENTA ($)', 'ORIGEN', 'UBICACIÓN', 'STOCK', 'IMAGEN'
];

const MOVEMENTS_HEADERS = [
  'FECHA', 'TIPO', 'CODIGO', 'DESCRIPCION', 'CANTIDAD',
  'STOCK ANTERIOR', 'STOCK NUEVO', 'MOTIVO', 'DOCUMENTO', 'USUARIO'
];

const CLIENTS_HEADERS = [
  'NOMBRE', 'EMAIL', 'DIRECCION', 'TELEFONO1', 'TELEFONO2', 'TELEFONO 3', 'TELEFONO 4'
];

const SUPPLIERS_HEADERS = [
  'NOMBRE', 'DIRECCION', 'TELEFONO', 'CORREO', 'UBICACION'
];

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
    else if (action === 'getClients') result = getClients();
    else if (action === 'getSuppliers') result = getSuppliers();
    else if (action === 'getOrders') result = getOrders();
    else if (action === 'getMovements') result = getMovements();
    else result = { error: 'Acción GET no válida: ' + action };
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

    if (action === 'saveProduct') result = saveProduct(payload);
    else if (action === 'registerBatchMovements') result = registerBatchMovements(payload);
    else if (action === 'saveClient') result = saveClient(payload);
    else if (action === 'saveSupplier') result = saveSupplier(payload);
    else if (action === 'saveBatchOrders') result = saveBatchOrders(payload);
    else if (action === 'deleteProduct') result = deleteProduct(payload);
    else if (action === 'deleteClient') result = deleteClient(payload);
    else if (action === 'deleteSupplier') result = deleteSupplier(payload);
    else if (action === 'deleteOrder') result = deleteOrder(payload);
    else result = { error: 'Acción POST no válida: ' + action };
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
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    let cleanHeader = String(header).trim().toUpperCase();
    let value = row[index];
    
    if (cleanHeader === 'ESTATUS LOGÍSTICO' || cleanHeader === 'ESTATUS LOGISTICO' || cleanHeader === 'ESTATUS') {
      obj['ESTATUS'] = value !== undefined && value !== null && String(value).trim() !== '' ? String(value).trim() : 'Solicitado';
    } else if (cleanHeader === 'CANTIDAD PENDIENTE' || cleanHeader === 'CANTIDAD_PENDIENTE') {
      obj['CANTIDAD_PENDIENTE'] = value !== undefined && value !== null ? parseInt(value, 10) || 0 : 0;
    } else {
      obj[cleanHeader] = value !== undefined && value !== null ? String(value).trim() : '';
    }
  });
  return obj;
}

function extractClientName(docField) {
  if (!docField) return 'Cliente General';
  const match = docField.match(/\(([^)]+)\)/);
  if (match && match[1]) return match[1].trim();
  return 'Cliente General';
}

function getDashboardMetrics() {
  try {
    const products = getProducts();
    const movements = getMovements();

    const priceMap = {};
    let totalProductos = products.length;
    let valorTotalInventario = 0;
    let productosStockBajo = 0;
    let productosStockNormal = 0;

    products.forEach(p => {
      const stock = parseInt(p.STOCK, 10) || 0;
      const precioCompra = parseFloat(p['PRECIO COMPRA ($)']) || 0;
      const precioVenta = parseFloat(p['PRECIO VENTA ($)']) || 0;
      const code = String(p.CODIGO).trim();

      priceMap[code] = precioVenta;
      valorTotalInventario += (stock * precioCompra);
      if (stock <= LOW_STOCK_THRESHOLD) {
        productosStockBajo++;
      } else {
        productosStockNormal++;
      }
    });

    let totalEntradas = 0;
    let totalSalidas = 0;
    const salesMap = {};
    const clientSalesMap = {};

    movements.forEach(m => {
      const cant = parseInt(m.CANTIDAD, 10) || 0;
      if (m.TIPO === 'ENTRADA') {
        totalEntradas += cant;
      } else if (m.TIPO === 'SALIDA') {
        totalSalidas += cant;
        const code = String(m.CODIGO).trim();
        salesMap[code] = (salesMap[code] || 0) + cant;

        const clientName = extractClientName(m.DOCUMENTO);
        const unitPrice = priceMap[code] || 0;
        clientSalesMap[clientName] = (clientSalesMap[clientName] || 0) + (cant * unitPrice);
      }
    });

    const topProducts = products.map(p => {
      const code = String(p.CODIGO).trim();
      return {
        CODIGO: p.CODIGO,
        DESCRIPCION: p.DESCRIPCION,
        TOTAL_VENTAS: salesMap[code] || 0
      };
    })
    .sort((a, b) => b.TOTAL_VENTAS - a.TOTAL_VENTAS)
    .slice(0, 5);

    const topClients = Object.keys(clientSalesMap).map(clientName => {
      return {
        NOMBRE: clientName,
        MONTO_TOTAL: clientSalesMap[clientName] || 0
      };
    })
    .sort((a, b) => b.MONTO_TOTAL - a.MONTO_TOTAL)
    .slice(0, 5);

    return {
      totalProductos: totalProductos || 0,
      valorTotalInventario: valorTotalInventario || 0,
      productosStockBajo: productosStockBajo || 0,
      productosStockNormal: productosStockNormal || 0,
      totalEntradas: totalEntradas || 0,
      totalSalidas: totalSalidas || 0,
      topProducts: topProducts || [],
      topClients: topClients || []
    };
  } catch (error) {
    throw new Error('Error al calcular métricas: ' + error.message);
  }
}

function getProducts() {
  const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(data[0], row));
}

function getMovements() {
  const sheet = getSheet(SHEET_MOVEMENTS, MOVEMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(data[0], row)).reverse();
}

function getClients() {
  const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(data[0], row));
}

function getSuppliers() {
  const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(data[0], row));
}

function getOrders() {
  const sheet = getSheet(SHEET_ORDERS, ORDERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => rowToObject(data[0], row)).reverse();
}

function saveProduct(product) {
  const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const codigoIndex = headers.indexOf('CODIGO');
  
  if (codigoIndex === -1) throw new Error('Columna CODIGO no encontrada');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codigoIndex]).trim().toLowerCase() === String(product.CODIGO).trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }

  const row = headers.map(header => product[header] !== undefined ? product[header] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
    return { success: true, message: 'Producto actualizado correctamente' };
  } else {
    sheet.appendRow(row);
    return { success: true, message: 'Producto creado correctamente' };
  }
}

function registerBatchMovements(payload) {
  if (!payload || !payload.items || payload.items.length === 0) {
    throw new Error('Debe especificar al menos un producto');
  }

  const sheetMov = getSheet(SHEET_MOVEMENTS, MOVEMENTS_HEADERS);
  const sheetProd = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
  const prodData = sheetProd.getDataRange().getValues();
  const prodHeaders = prodData[0].map(h => String(h).trim().toUpperCase());
  const codeIdx = prodHeaders.indexOf('CODIGO');
  const descIdx = prodHeaders.indexOf('DESCRIPCION');
  const stockIdx = prodHeaders.indexOf('STOCK');

  if (codeIdx === -1 || stockIdx === -1) {
    throw new Error('La estructura de la hoja de Productos es incorrecta.');
  }

  const now = new Date();
  const docRef = payload.numFactura 
    ? `${payload.numFactura} (${payload.entidad || 'General'})`
    : (payload.entidad || 'General');

  payload.items.forEach(item => {
    let prodRowIdx = -1;
    let currentStock = 0;
    let desc = '';

    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][codeIdx]).trim().toLowerCase() === String(item.codigo).trim().toLowerCase()) {
        prodRowIdx = i + 1;
        currentStock = parseInt(prodData[i][stockIdx], 10) || 0;
        desc = prodData[i][descIdx] || '';
        break;
      }
    }

    if (prodRowIdx === -1) throw new Error(`El producto con código ${item.codigo} no existe.`);

    let newStock = currentStock;
    const cant = parseInt(item.cantidad, 10) || 0;

    if (payload.tipo === 'ENTRADA') newStock += cant;
    else if (payload.tipo === 'SALIDA') newStock = Math.max(0, currentStock - cant);
    else if (payload.tipo === 'AJUSTE') newStock = cant;

    sheetProd.getRange(prodRowIdx, stockIdx + 1).setValue(newStock);

    sheetMov.appendRow([
      now, payload.tipo, item.codigo, desc, cant, currentStock, newStock, payload.motivo || '', docRef, payload.usuario || 'Admin'
    ]);
  });

  return { success: true, message: 'Movimientos registrados exitosamente' };
}

function saveBatchOrders(payload) {
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
      if (header === 'CANTIDAD_PENDIENTE' || header === 'CANTIDAD PENDIENTE') return orderObj.CANTIDAD_PENDIENTE;
      return orderObj[header] !== undefined ? orderObj[header] : '';
    });
    newRows.push(row);
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { success: true, message: `Se guardó correctamente el pedido #${payload.numPedido}` };
}

function saveClient(client) {
  const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const nameIndex = headers.indexOf('NOMBRE');

  let rowIndex = -1;
  if (nameIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameIndex]).trim().toLowerCase() === String(client.NOMBRE).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const row = headers.map(header => client[header] !== undefined ? client[header] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
    return { success: true, message: 'Cliente actualizado correctamente' };
  } else {
    sheet.appendRow(row);
    return { success: true, message: 'Cliente registrado correctamente' };
  }
}

function saveSupplier(supplier) {
  const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const nameIndex = headers.indexOf('NOMBRE');

  let rowIndex = -1;
  if (nameIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameIndex]).trim().toLowerCase() === String(supplier.NOMBRE).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const row = headers.map(header => supplier[header] !== undefined ? supplier[header] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
    return { success: true, message: 'Proveedor actualizado correctamente' };
  } else {
    sheet.appendRow(row);
    return { success: true, message: 'Proveedor registrado correctamente' };
  }
}

function deleteProduct(codigo) {
  const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const codigoIndex = headers.indexOf('CODIGO');

  if (codigoIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codigoIndex]).trim().toLowerCase() === String(codigo).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Producto eliminado correctamente' };
      }
    }
  }
  throw new Error('Producto no encontrado');
}

function deleteOrder(id) {
  const sheet = getSheet(SHEET_ORDERS, ORDERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const idIndex = headers.indexOf('ID');
  let count = 0;
  
  if (idIndex !== -1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIndex]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
  }
  if (count > 0) return { success: true, message: 'Pedido eliminado correctamente' };
  throw new Error('Pedido no encontrado');
}

function deleteClient(nombre) {
  const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const nameIndex = headers.indexOf('NOMBRE');

  if (nameIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameIndex]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Cliente eliminado correctamente' };
      }
    }
  }
  throw new Error('Cliente no encontrado');
}

function deleteSupplier(nombre) {
  const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const nameIndex = headers.indexOf('NOMBRE');

  if (nameIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameIndex]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Proveedor eliminado correctamente' };
      }
    }
  }
  throw new Error('Proveedor no encontrado');
}
