// Configuración
var SPREADSHEET_ID = '1D2RsD_g-ltoCZjodiNjKHhDTgfRq8gyziR0ya30GV70';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sistema Gestión Tienda Abarrotes');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Función para obtener datos de cualquier hoja
function getSheetData(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error('No se encontró la hoja: ' + sheetName);
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var headers = data[0];
    var jsonData = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        // Convertir a string para evitar problemas
        var value = data[i][j];
        if (value === null || value === undefined) value = '';
        row[headers[j]] = value.toString();
      }
      jsonData.push(row);
    }
    return jsonData;
  } catch (error) {
    console.error('Error en getSheetData: ' + error.toString());
    return [];
  }
}

// Login de usuario - CORREGIDO
function loginUsuario(email, password) {
  try {
    console.log('Intentando login con:', email, password);
    
    var usuarios = getSheetData('USUARIOS');
    console.log('Usuarios encontrados:', usuarios.length);
    
    var usuario = usuarios.find(function(user) {
      var emailMatch = user.Email && user.Email.trim() === email.trim();
      var passwordMatch = user.Contraseña && user.Contraseña.trim() === password.trim();
      return emailMatch && passwordMatch;
    });
    
    console.log('Usuario encontrado:', usuario);
    return usuario || null;
  } catch (error) {
    console.error('Error en login: ' + error.toString());
    return null;
  }
}

// Obtener estadísticas del dashboard
function getEstadisticas() {
  try {
    var productos = getSheetData('PRODUCTOS');
    var ventas = getSheetData('VENTAS');
    
    // Calcular ventas de hoy
    var hoy = new Date();
    var ventasHoy = ventas.filter(function(venta) {
      if (!venta.Fecha_Venta) return false;
      try {
        var fechaVenta = new Date(venta.Fecha_Venta);
        return fechaVenta.toDateString() === hoy.toDateString();
      } catch (e) {
        return false;
      }
    });
    
    var totalVentasHoy = ventasHoy.reduce(function(sum, venta) {
      return sum + (parseFloat(venta.Total_Venta) || 0);
    }, 0);
    
    // Productos con stock bajo
    var productosBajoStock = productos.filter(function(producto) {
      var stockActual = parseInt(producto.Stock_Actual) || 0;
      var stockMinimo = parseInt(producto.Stock_Minimo) || 5; // Default 5 si no existe
      return stockActual <= stockMinimo;
    });
    
    return {
      totalProductos: productos.length,
      totalVentasHoy: totalVentasHoy,
      alertasStock: productosBajoStock.length,
      productosBajoStock: productosBajoStock.length
    };
  } catch (error) {
    console.error('Error en getEstadisticas:', error);
    return {
      totalProductos: 0,
      totalVentasHoy: 0,
      alertasStock: 0,
      productosBajoStock: 0
    };
  }
}

// Obtener producto por código de barras - CORREGIDO
function getProductoByCodigo(codigo) {
  try {
    console.log('Buscando producto con código:', codigo);
    
    var productos = getSheetData('PRODUCTOS');
    console.log('Productos totales:', productos.length);
    
    var producto = productos.find(function(prod) {
      var codigoMatch = prod.Codigo_Barras && prod.Codigo_Barras.toString().trim() === codigo.toString().trim();
      var idMatch = prod.ID_Producto && prod.ID_Producto.toString().trim() === codigo.toString().trim();
      return codigoMatch || idMatch;
    });
    
    console.log('Producto encontrado:', producto);
    return producto || null;
  } catch (error) {
    console.error('Error buscando producto: ' + error.toString());
    return null;
  }
}



// Obtener todos los productos
function getProductos() {
  return getSheetData('PRODUCTOS');
}

// Obtener alertas de stock
function getAlertasStock() {
  try {
    var productos = getSheetData('PRODUCTOS');
    var alertas = productos.filter(function(producto) {
      var stockActual = parseInt(producto.Stock_Actual) || 0;
      var stockMinimo = parseInt(producto.Stock_Minimo) || 5;
      return stockActual <= stockMinimo;
    }).map(function(producto) {
      return {
        Nombre_Producto: producto.Nombre_Producto || 'Sin nombre',
        Stock_Actual: producto.Stock_Actual || '0',
        Stock_Minimo: producto.Stock_Minimo || '5',
        Diferencia: (parseInt(producto.Stock_Actual) || 0) - (parseInt(producto.Stock_Minimo) || 5)
      };
    });
    
    return alertas;
  } catch (error) {
    console.error('Error en getAlertasStock:', error);
    return [];
  }
}

// Registrar una nueva venta - COMPLETAMENTE FUNCIONAL
function registrarVenta(ventaData, carrito) {
  try {
    console.log('Registrando venta:', ventaData);
    console.log('Carrito:', carrito);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ventasSheet = ss.getSheetByName('VENTAS');
    var detalleSheet = ss.getSheetByName('DETALLE_VENTAS');
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    
    // Generar ID de venta
    var lastVenta = ventasSheet.getLastRow();
    var newVentaId = 'VTA' + String(lastVenta).padStart(3, '0');
    
    // Agregar venta principal
    ventasSheet.appendRow([
      newVentaId,
      new Date(),
      Utilities.formatDate(new Date(), 'GMT-5', 'HH:mm'),
      ventaData.idCajero,
      parseFloat(ventaData.total),
      ventaData.metodoPago,
      'Completada',
      'T' + String(lastVenta).padStart(3, '0')
    ]);
    
    // Agregar detalles
    carrito.forEach(function(item) {
      detalleSheet.appendRow([
        'DET' + Utilities.getUuid().substring(0, 8),
        newVentaId,
        item.idProducto,
        parseInt(item.cantidad),
        parseFloat(item.precio),
        parseFloat(item.subtotal)
      ]);
    });
    
    // Actualizar stock para cada producto
    carrito.forEach(function(item) {
      actualizarStockProducto(item.idProducto, -parseInt(item.cantidad), 'Venta: ' + newVentaId, ventaData.idCajero);
    });
    
    console.log('Venta registrada exitosamente:', newVentaId);
    return newVentaId;
    
  } catch (error) {
    console.error('Error registrando venta:', error);
    throw new Error('No se pudo registrar la venta: ' + error.toString());
  }
}

// Función auxiliar para actualizar stock
function actualizarStockProducto(productoId, cantidad, motivo, usuarioId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    var movimientosSheet = ss.getSheetByName('INVENTARIO_MOVIMIENTOS');
    
    var productosData = productosSheet.getDataRange().getValues();
    
    // Buscar producto
    for (var i = 1; i < productosData.length; i++) {
      if (productosData[i][0] === productoId) {
        var stockAnterior = parseInt(productosData[i][5]) || 0;
        var stockNuevo = stockAnterior + cantidad;
        
        // Actualizar stock
        productosSheet.getRange(i + 1, 6).setValue(stockNuevo);
        
        // Registrar movimiento
        movimientosSheet.appendRow([
          'MOV' + Utilities.getUuid().substring(0, 8),
          new Date(),
          productoId,
          cantidad < 0 ? 'Venta' : 'Compra',
          cantidad,
          stockAnterior,
          stockNuevo,
          usuarioId,
          motivo
        ]);
        
        break;
      }
    }
  } catch (error) {
    console.error('Error actualizando stock:', error);
  }
}

// Obtener movimientos de inventario
function getMovimientosInventario() {
  return getSheetData('INVENTARIO_MOVIMIENTOS');
}

// Función para diagnosticar el sistema
function diagnosticarSistema() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var sheetNames = sheets.map(function(sheet) {
      return sheet.getName();
    });
    
    var datos = {};
    sheetNames.forEach(function(sheetName) {
      datos[sheetName] = getSheetData(sheetName).length;
    });
    
    return {
      status: 'OK',
      sheets: sheetNames,
      datos: datos,
      spreadsheetId: SPREADSHEET_ID
    };
  } catch (error) {
    return {
      status: 'ERROR',
      error: error.toString(),
      spreadsheetId: SPREADSHEET_ID
    };
  }
}


// === GESTIÓN DE PRODUCTOS - HU007 ===

// HU007.1 - Registrar entrada de mercancía
function registrarEntradaMercancia(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    var movimientosSheet = ss.getSheetByName('INVENTARIO_MOVIMIENTOS');
    
    // Buscar si el producto ya existe
    var productosData = productosSheet.getDataRange().getValues();
    var productoExistente = null;
    var filaProducto = -1;
    
    for (var i = 1; i < productosData.length; i++) {
      if (productosData[i][0] === productoData.idProducto || 
          productosData[i][1] === productoData.codigoBarras) {
        productoExistente = productosData[i];
        filaProducto = i + 1;
        break;
      }
    }
    
    if (productoExistente) {
      // Actualizar producto existente
      var stockAnterior = parseInt(productoExistente[5]) || 0;
      var stockNuevo = stockAnterior + parseInt(productoData.cantidad);
      
      productosSheet.getRange(filaProducto, 6).setValue(stockNuevo);
      
      // Registrar movimiento
      registrarMovimientoInventario(
        productoData.idProducto,
        parseInt(productoData.cantidad),
        stockAnterior,
        stockNuevo,
        productoData.idUsuario,
        'Entrada de mercancía: ' + productoData.motivo
      );
      
      return { 
        success: true, 
        tipo: 'actualizado',
        stockAnterior: stockAnterior,
        stockNuevo: stockNuevo
      };
      
    } else {
      // Crear nuevo producto
      var lastRow = productosSheet.getLastRow();
      var newId = 'PRO' + String(lastRow).padStart(3, '0');
      
      productosSheet.appendRow([
        newId,
        productoData.codigoBarras,
        productoData.nombre,
        productoData.categoria || 'General',
        parseFloat(productoData.precio),
        parseInt(productoData.cantidad),
        parseInt(productoData.stockMinimo) || 5,
        productoData.proveedor || '',
        new Date(),
        'Activo'
      ]);
      
      // Registrar movimiento
      registrarMovimientoInventario(
        newId,
        parseInt(productoData.cantidad),
        0,
        parseInt(productoData.cantidad),
        productoData.idUsuario,
        'Nuevo producto: ' + productoData.motivo
      );
      
      return { 
        success: true, 
        tipo: 'nuevo',
        productoId: newId
      };
    }
    
  } catch (error) {
    console.error('Error registrando entrada:', error);
    return { success: false, error: error.toString() };
  }
}

// === GESTIÓN DE PRODUCTOS - NUEVAS FUNCIONES ===

// HU007.1 - Registrar producto por código de barras (NUEVA FUNCIÓN)
function registrarProductoPorCodigo(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    
    // Verificar si el código de barras ya existe
    var productos = getSheetData('PRODUCTOS');
    var productoExistente = productos.find(function(p) {
      return p.Codigo_Barras === productoData.codigoBarras;
    });
    
    if (productoExistente) {
      return { 
        success: false, 
        error: 'El código de barras ya existe para el producto: ' + productoExistente.Nombre_Producto
      };
    }
    
    // Generar nuevo ID
    var lastRow = productosSheet.getLastRow();
    var newId = 'PRO' + String(lastRow).padStart(3, '0');
    
    // Registrar nuevo producto
    productosSheet.appendRow([
      newId,
      productoData.codigoBarras,
      productoData.nombre,
      productoData.categoria || 'General',
      parseFloat(productoData.precio),
      parseInt(productoData.stockInicial) || 0,
      parseInt(productoData.stockMinimo) || 5,
      productoData.proveedor || '',
      new Date(),
      'Activo'
    ]);
    
    // Registrar movimiento de inventario si hay stock inicial
    if (parseInt(productoData.stockInicial) > 0) {
      registrarMovimientoInventario(
        newId,
        parseInt(productoData.stockInicial),
        0,
        parseInt(productoData.stockInicial),
        productoData.idUsuario,
        'Registro inicial de producto'
      );
    }
    
    return { 
      success: true, 
      productoId: newId,
      mensaje: 'Producto registrado exitosamente'
    };
    
  } catch (error) {
    console.error('Error registrando producto:', error);
    return { success: false, error: error.toString() };
  }
}

// Función para buscar producto por código de barras (para verificar existencia)
function verificarCodigoBarras(codigoBarras) {
  try {
    var productos = getSheetData('PRODUCTOS');
    var productoExistente = productos.find(function(p) {
      return p.Codigo_Barras === codigoBarras;
    });
    
    return productoExistente || null;
  } catch (error) {
    return null;
  }
}

// Buscar productos por término
function buscarProductos(termino) {
  try {
    var productos = getSheetData('PRODUCTOS');
    if (!termino) return productos;
    
    return productos.filter(function(producto) {
      var busqueda = termino.toLowerCase();
      return (
        (producto.Nombre_Producto && producto.Nombre_Producto.toLowerCase().includes(busqueda)) ||
        (producto.Codigo_Barras && producto.Codigo_Barras.toLowerCase().includes(busqueda)) ||
        (producto.ID_Producto && producto.ID_Producto.toLowerCase().includes(busqueda)) ||
        (producto.Categoria && producto.Categoria.toLowerCase().includes(busqueda))
      );
    });
  } catch (error) {
    return [];
  }
}

// HU007.2 - Ajustar stock manualmente
function ajustarStockManual(ajusteData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    var productosData = productosSheet.getDataRange().getValues();
    
    // Buscar producto
    for (var i = 1; i < productosData.length; i++) {
      if (productosData[i][0] === ajusteData.idProducto) {
        var stockAnterior = parseInt(productosData[i][5]) || 0;
        var stockNuevo = parseInt(ajusteData.nuevoStock);
        
        // Actualizar stock
        productosSheet.getRange(i + 1, 6).setValue(stockNuevo);
        
        // Registrar movimiento
        registrarMovimientoInventario(
          ajusteData.idProducto,
          stockNuevo - stockAnterior,
          stockAnterior,
          stockNuevo,
          ajusteData.idUsuario,
          'Ajuste manual: ' + ajusteData.motivo
        );
        
        return { 
          success: true,
          stockAnterior: stockAnterior,
          stockNuevo: stockNuevo
        };
      }
    }
    
    return { success: false, error: 'Producto no encontrado' };
    
  } catch (error) {
    console.error('Error ajustando stock:', error);
    return { success: false, error: error.toString() };
  }
}

// Función auxiliar para registrar movimientos
function registrarMovimientoInventario(productoId, cantidad, stockAnterior, stockNuevo, usuarioId, motivo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var movimientosSheet = ss.getSheetByName('INVENTARIO_MOVIMIENTOS');
    
    var movimientoId = 'MOV' + new Date().getTime();
    var tipoMovimiento = cantidad >= 0 ? 'Entrada' : 'Salida';
    
    movimientosSheet.appendRow([
      movimientoId,
      new Date(),
      productoId,
      tipoMovimiento,
      cantidad,
      stockAnterior,
      stockNuevo,
      usuarioId,
      motivo
    ]);
    
  } catch (error) {
    console.error('Error registrando movimiento:', error);
  }
}

// === GESTIÓN DE PRODUCTOS - ACTUALIZAR CATEGORÍAS ===

// Obtener categorías de productos actualizadas
function getCategoriasProductos() {
  return [
    'Despensa',
    'Lácteos y huevos', 
    'Carnes y embutidos',
    'Bebidas',
    'Panadería y galletas',
    'Limpieza y hogar',
    'Cuidado personal',
    'Snacks y golosinas',
    'Productos congelados',
    'Mascotas'
  ];
}





// Buscar productos por término
function buscarProductos(termino) {
  try {
    var productos = getSheetData('PRODUCTOS');
    if (!termino) return productos;
    
    return productos.filter(function(producto) {
      var busqueda = termino.toLowerCase();
      return (
        (producto.Nombre_Producto && producto.Nombre_Producto.toLowerCase().includes(busqueda)) ||
        (producto.Codigo_Barras && producto.Codigo_Barras.toLowerCase().includes(busqueda)) ||
        (producto.ID_Producto && producto.ID_Producto.toLowerCase().includes(busqueda)) ||
        (producto.Categoria && producto.Categoria.toLowerCase().includes(busqueda))
      );
    });
  } catch (error) {
    return [];
  }
}