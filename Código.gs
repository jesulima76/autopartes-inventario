/**
 * AutoPartes Inventario - Backend API con Google Sheets
 * Base de datos: Hoja de cálculo de Google
 */

const SHEET_PRODUCTS = 'Productos';
const SHEET_MOVEMENTS = 'Movimientos';
const SHEET_CLIENTS = 'Clientes';
const SHEET_SUPPLIERS = 'Proveedores';
const SHEET_ORDERS = 'Pedidos';

const LOW_STOCK_THRESHOLD = 5;

// Columnas EXACTAS
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

// --- API HTTP ENDPOINTS FOR VERCEL ---

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

    if (action === 'saveProduct') result = saveProduct(payload);
    else if (action === 'registerBatchMovements') result = registerBatchMovements(payload);
    else if (action === 'saveClient') result = saveClient(payload);
    else if (action === 'saveSupplier') result = saveSupplier(payload);
    else if (action === 'saveBatchOrders') result = saveBatchOrders(payload);
    else if (action === 'deleteProduct') result = deleteProduct(payload);
    else if (action === 'deleteClient') result = deleteClient(payload);
    else if (action === 'deleteSupplier') result = deleteSupplier(payload);
    else if (action === 'deleteOrder') result = deleteOrder(payload);
    else result = { error: 'Acción POST no válida' };
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- HELPER FUNCTIONS ---

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
    let value = row[index];
    if (header === 'PRECIO COMPRA ($)' || header === 'PRECIO VENTA ($)' || header === 'PRECIO_COMPRA') {
      value = parseFloat(value) || 0;
    } else if (header === 'STOCK' || header === 'CANTIDAD' || header === 'STOCK ANTERIOR' || header === 'STOCK NUEVO' || header === 'CANTIDAD_PEDIDO' || header === 'CANTIDAD_EMBARCADA' || header === 'CANTIDAD_PENDIENTE') {
      value = parseInt(value, 10) || 0;
    } else if (value instanceof Date) {
      value = value.toISOString();
    }
    obj[header] = value !== undefined && value !== null ? value : '';
  });
  return obj;
}

function extractClientName(docField) {
  if (!docField) return 'Cliente General';
  const match = docField.match(/\(([^)]+)\)/);
  if (match && match[1]) {
    return match[1].trim();
  }
  return 'Cliente General';
}

// --- DASHBOARD METRICS ---

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
        const totalAmount = cant * unitPrice;
        clientSalesMap[clientName] = (clientSalesMap[clientName] || 0) + totalAmount;
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
    console.error('Error en getDashboardMetrics:', error);
    throw new Error('Error al calcular métricas del Dashboard: ' + error.message);
  }
}

// --- PRODUCTOS Y MOVIMIENTOS ---

function getProducts() {
  try {
    const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => rowToObject(data[0], row));
  } catch (error) {
    throw new Error('No se pudieron cargar los productos: ' + error.message);
  }
}

function getMovements() {
  try {
    const sheet = getSheet(SHEET_MOVEMENTS, MOVEMENTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => rowToObject(data[0], row)).reverse();
  } catch (error) {
    throw new Error('No se pudieron cargar los movimientos: ' + error.message);
  }
}

function saveProduct(product) {
  try {
    const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const codigoIndex = headers.indexOf('CODIGO');
    
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codigoIndex]) === String(product.CODIGO)) {
        existingRowIndex = i + 1;
        break;
      }
    }

    const rowData = headers.map(header => product[header] !== undefined ? product[header] : '');

    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: 'Producto guardado correctamente' };
  } catch (error) {
    throw new Error('No se pudo guardar el producto: ' + error.message);
  }
}

function deleteProduct(codigo) {
  try {
    const sheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const codigoIndex = data[0].indexOf('CODIGO');
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codigoIndex]) === String(codigo)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Producto eliminado' };
      }
    }
    throw new Error('Producto no encontrado');
  } catch (error) {
    throw new Error('No se pudo eliminar el producto: ' + error.message);
  }
}

function registerBatchMovements(data) {
  try {
    if (!data || !data.items || data.items.length === 0) {
      throw new Error('Debe proporcionar al menos un producto.');
    }

    const productsSheet = getSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const movementsSheet = getSheet(SHEET_MOVEMENTS, MOVEMENTS_HEADERS);
    
    const productData = productsSheet.getDataRange().getValues();
    const productHeaders = productData[0];
    const codigoIndex = productHeaders.indexOf('CODIGO');
    const stockIndex = productHeaders.indexOf('STOCK');
    const descIndex = productHeaders.indexOf('DESCRIPCION');
    
    const productMap = new Map();
    for (let i = 1; i < productData.length; i++) {
      const code = String(productData[i][codigoIndex]).trim();
      productMap.set(code, {
        rowIndex: i + 1,
        stock: parseInt(productData[i][stockIndex], 10) || 0,
        descripcion: productData[i][descIndex] || ''
      });
    }

    data.items.forEach(item => {
      const codeStr = String(item.codigo).trim();
      const prod = productMap.get(codeStr);
      if (!prod) throw new Error(`El código de producto "${item.codigo}" no existe en el inventario.`);
      
      const cantidad = parseInt(item.cantidad, 10);
      if (isNaN(cantidad) || cantidad <= 0) throw new Error(`Cantidad inválida para el código "${item.codigo}".`);

      if (data.tipo === 'SALIDA' && prod.stock < cantidad) {
        throw new Error(`Stock insuficiente para "${prod.descripcion}" (Stock actual: ${prod.stock}, Requerido: ${cantidad}).`);
      }
    });

    const now = new Date();
    const movementRows = [];

    let docParts = [];
    if (data.numFactura) docParts.push(`FACTURA #${data.numFactura}`);
    if (data.documentoReferencia) docParts.push(data.documentoReferencia);
    if (data.entidad) docParts.push(`(${data.entidad})`);

    const docRef = docParts.join(' ') || '-';

    data.items.forEach(item => {
      const codeStr = String(item.codigo).trim();
      const prod = productMap.get(codeStr);
      const cantidad = parseInt(item.cantidad, 10);
      let stockAnterior = prod.stock;
      let newStock = prod.stock;

      if (data.tipo === 'ENTRADA') newStock += cantidad;
      else if (data.tipo === 'SALIDA') newStock -= cantidad;
      else if (data.tipo === 'AJUSTE') newStock = cantidad;

      prod.stock = newStock;
      productsSheet.getRange(prod.rowIndex, stockIndex + 1).setValue(newStock);

      const row = MOVEMENTS_HEADERS.map(header => {
        switch(header) {
          case 'FECHA': return now;
          case 'TIPO': return data.tipo;
          case 'CODIGO': return codeStr;
          case 'DESCRIPCION': return prod.descripcion;
          case 'CANTIDAD': return cantidad;
          case 'STOCK ANTERIOR': return stockAnterior;
          case 'STOCK NUEVO': return newStock;
          case 'MOTIVO': return data.motivo || '';
          case 'DOCUMENTO': return docRef;
          case 'USUARIO': return data.usuario || 'Sistema';
          default: return '';
        }
      });
      movementRows.push(row);
    });

    if (movementRows.length > 0) {
      movementsSheet.getRange(movementsSheet.getLastRow() + 1, 1, movementRows.length, MOVEMENTS_HEADERS.length).setValues(movementRows);
    }

    return { success: true, message: `Se registraron ${data.items.length} producto(s) correctamente.` };
  } catch (error) {
    console.error('Error en registerBatchMovements:', error);
    throw new Error(error.message || 'Error al procesar la lista de movimientos');
  }
}

// --- CLIENTES ---

function getClients() {
  try {
    const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => rowToObject(data[0], row));
  } catch (error) {
    throw new Error('No se pudieron cargar los clientes: ' + error.message);
  }
}

function saveClient(client) {
  try {
    const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nombreIndex = headers.indexOf('NOMBRE');
    
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nombreIndex]).trim().toLowerCase() === String(client.NOMBRE).trim().toLowerCase()) {
        existingRowIndex = i + 1;
        break;
      }
    }

    const rowData = CLIENTS_HEADERS.map(header => client[header] !== undefined ? client[header] : '');

    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: 'Cliente guardado correctamente' };
  } catch (error) {
    throw new Error('No se pudo guardar el cliente: ' + error.message);
  }
}

function deleteClient(nombre) {
  try {
    const sheet = getSheet(SHEET_CLIENTS, CLIENTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const nombreIndex = data[0].indexOf('NOMBRE');
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nombreIndex]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Cliente eliminado' };
      }
    }
    throw new Error('Cliente no encontrado');
  } catch (error) {
    throw new Error('No se pudo eliminar el cliente: ' + error.message);
  }
}

// --- PROVEEDORES ---

function getSuppliers() {
  try {
    const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => rowToObject(data[0], row));
  } catch (error) {
    throw new Error('No se pudieron cargar los proveedores: ' + error.message);
  }
}

function saveSupplier(supplier) {
  try {
    const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nombreIndex = headers.indexOf('NOMBRE');
    
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nombreIndex]).trim().toLowerCase() === String(supplier.NOMBRE).trim().toLowerCase()) {
        existingRowIndex = i + 1;
        break;
      }
    }

    const rowData = SUPPLIERS_HEADERS.map(header => supplier[header] !== undefined ? supplier[header] : '');

    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: 'Proveedor guardado correctamente' };
  } catch (error) {
    throw new Error('No se pudo guardar el proveedor: ' + error.message);
  }
}

function deleteSupplier(nombre) {
  try {
    const sheet = getSheet(SHEET_SUPPLIERS, SUPPLIERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const nombreIndex = data[0].indexOf('NOMBRE');
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nombreIndex]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Proveedor eliminado' };
      }
    }
    throw new Error('Proveedor no encontrado');
  } catch (error) {
    throw new Error('No se pudo eliminar el proveedor: ' + error.message);
  }
}

// --- MÓDULO DE PEDIDOS INTERNACIONALES MULTI-PRODUCTO ---

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
    const idIndex = data[0].indexOf('ID');

    if (payload.batchId) {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idIndex]) === String(payload.batchId)) {
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

      const row = ORDERS_HEADERS.map(header => orderObj[header] !== undefined ? orderObj[header] : '');
      newRows.push(row);
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, ORDERS_HEADERS.length).setValues(newRows);
    }

    return { success: true, message: `Se registraron ${newRows.length} ítem(s) para el pedido #${payload.numPedido}` };
  } catch (error) {
    console.error('Error en saveBatchOrders:', error);
    throw new Error('No se pudo guardar el pedido internacional: ' + error.message);
  }
}

function deleteOrder(id) {
  try {
    const sheet = getSheet(SHEET_ORDERS, ORDERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const idIndex = data[0].indexOf('ID');
    let count = 0;
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIndex]) === String(id)) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }

    if (count > 0) {
      return { success: true, message: 'Pedido eliminado correctamente' };
    }
    throw new Error('Pedido no encontrado');
  } catch (error) {
    throw new Error('No se pudo eliminar el pedido: ' + error.message);
  }
}