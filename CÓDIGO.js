// ================= Configuración =================
var SPREADSHEET_ID = '1Rm9IoAmkgYFu4Zs1S5i2M4ssJfKuSufrEJIVJzJiWSs';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sistema Gestión Tienda Abarrotes');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ================= Utilidades base =================

// Función genérica para leer una hoja y devolver JSON
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

// ================= Autenticación =================

// Login de usuario - ACTUALIZADO para nueva BD
function loginUsuario(email, password) {
  try {
    console.log('Intentando login con:', email, password);
    
    var usuarios = getSheetData('USUARIO'); // Cambiado de USUARIOS a USUARIO
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

// ================ Estadísticas básicas (compat) ================
function getEstadisticas() {
  try {
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var ventas = getSheetData('VENTA'); // Cambiado de VENTAS a VENTA
    
    // Calcular ventas de hoy
    var hoy = new Date();
    var ventasHoy = ventas.filter(function(venta) {
      if (!venta.FechaHora_Venta) return false; // Cambiado de Fecha_Venta a FechaHora_Venta
      try {
        var fechaVenta = new Date(venta.FechaHora_Venta);
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
      var stockMinimo = parseInt(producto.Stock_Minimo) || 5;
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

// ================ Productos / Inventario ================
function getProductoByCodigo(codigo) {
  try {
    console.log('Buscando producto con código:', codigo);
    
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
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

function getProductos() {
  return getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
}

function getAlertasStock() {
  try {
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
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

// Registrar una nueva venta - ACTUALIZADO para nueva BD
function registrarVenta(ventaData, carrito) {
  try {
    console.log('Registrando venta:', ventaData);
    console.log('Carrito:', carrito);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ventasSheet = ss.getSheetByName('VENTA'); // Cambiado de VENTAS a VENTA
    var detalleSheet = ss.getSheetByName('DETALLE_VENTA'); // Cambiado de DETALLE_VENTAS a DETALLE_VENTA
    
    // Generar ID de venta
    var lastVenta = ventasSheet.getLastRow();
    var newVentaId = 'VTA' + String(lastVenta).padStart(3, '0');
    
    // Agregar venta principal
    ventasSheet.appendRow([
      newVentaId,
      new Date(),
      ventaData.idCajero,
      'TUR001', // ID_Turno temporal - puedes implementar lógica de turnos después
      parseFloat(ventaData.total),
      'Completada', // Estado_Venta
      'T' + String(lastVenta).padStart(3, '0') // Numero_Ticket
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
    
    // Registrar método de pago
    var ventaMetodoSheet = ss.getSheetByName('VENTA_METODO_PAGO');
    ventaMetodoSheet.appendRow([
      'VMP' + Utilities.getUuid().substring(0, 8),
      newVentaId,
      obtenerIdMetodoPago(ventaData.metodoPago),
      parseFloat(ventaData.total)
    ]);
    
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

// Función auxiliar para obtener ID de método de pago
function obtenerIdMetodoPago(nombreMetodo) {
  var metodos = getSheetData('METODO_PAGO');
  var metodo = metodos.find(function(m) {
    return m.Nombre_Metodo === nombreMetodo;
  });
  return metodo ? metodo.ID_Metodo : 'MP001'; // Default si no encuentra
}

function actualizarStockProducto(productoId, cantidad, motivo, usuarioId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var movimientosSheet = ss.getSheetByName('MOVIMIENTO_INVENTARIO'); // Cambiado de INVENTARIO_MOVIMIENTOS a MOVIMIENTO_INVENTARIO
    
    var productosData = productosSheet.getDataRange().getValues();
    
    // Buscar producto
    for (var i = 1; i < productosData.length; i++) {
      if (productosData[i][0] === productoId) {
        var stockAnterior = parseInt(productosData[i][5]) || 0; // Stock_Actual está en columna 5
        var stockNuevo = stockAnterior + cantidad;
        
        // Actualizar stock
        productosSheet.getRange(i + 1, 6).setValue(stockNuevo); // Stock_Actual está en columna 6
        
        // Determinar tipo de movimiento
        var tipoMovimiento = cantidad >= 0 ? 'ENTRADA' : 'SALIDA';
        var idTipo = tipoMovimiento === 'ENTRADA' ? 'TIPO001' : 'TIPO002';
        var idMotivo = 'MOT001'; // Motivo por defecto
        
        // Registrar movimiento en nueva estructura
        movimientosSheet.appendRow([
          'MOV' + Utilities.getUuid().substring(0, 8),
          new Date(),
          productoId,
          idTipo,
          idMotivo,
          cantidad,
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

function getMovimientosInventario() {
  return getSheetData('MOVIMIENTO_INVENTARIO'); // Cambiado de INVENTARIO_MOVIMIENTOS a MOVIMIENTO_INVENTARIO
}

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

// ============== GESTIÓN DE PRODUCTOS - HU007 ==============
function registrarEntradaMercancia(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    
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
      var stockAnterior = parseInt(productoExistente[5]) || 0;
      var stockNuevo = stockAnterior + parseInt(productoData.cantidad);
      productosSheet.getRange(filaProducto, 6).setValue(stockNuevo);
      registrarMovimientoInventario(
        productoData.idProducto,
        parseInt(productoData.cantidad),
        stockAnterior,
        stockNuevo,
        productoData.idUsuario,
        'Entrada de mercancía: ' + productoData.motivo
      );
      return { success: true, tipo: 'actualizado', stockAnterior: stockAnterior, stockNuevo: stockNuevo };
    } else {
      var lastRow = productosSheet.getLastRow();
      var newId = 'PRO' + String(lastRow).padStart(3, '0');
      productosSheet.appendRow([
        newId,
        productoData.codigoBarras,
        productoData.nombre,
        'CAT001', // ID_Categoria temporal
        parseFloat(productoData.precio),
        parseInt(productoData.cantidad),
        parseInt(productoData.stockMinimo) || 5,
        'PROV001', // ID_Proveedor temporal
        new Date(),
        'Activo'
      ]);
      registrarMovimientoInventario(
        newId,
        parseInt(productoData.cantidad),
        0,
        parseInt(productoData.cantidad),
        productoData.idUsuario,
        'Nuevo producto: ' + productoData.motivo
      );
      return { success: true, tipo: 'nuevo', productoId: newId };
    }
  } catch (error) {
    console.error('Error registrando entrada:', error);
    return { success: false, error: error.toString() };
  }
}

// === NUEVAS FUNCIONES DE PRODUCTOS ===
function registrarProductoPorCodigo(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var productoExistente = productos.find(function(p) {
      return p.Codigo_Barras === productoData.codigoBarras;
    });
    if (productoExistente) {
      return { 
        success: false, 
        error: 'El código de barras ya existe para el producto: ' + productoExistente.Nombre_Producto
      };
    }
    var lastRow = productosSheet.getLastRow();
    var newId = 'PRO' + String(lastRow).padStart(3, '0');
    productosSheet.appendRow([
      newId,
      productoData.codigoBarras,
      productoData.nombre,
      'CAT001', // ID_Categoria temporal - puedes implementar selección después
      parseFloat(productoData.precio),
      parseInt(productoData.stockInicial) || 0,
      parseInt(productoData.stockMinimo) || 5,
      'PROV001', // ID_Proveedor temporal
      new Date(),
      'Activo'
    ]);
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
    return { success: true, productoId: newId, mensaje: 'Producto registrado exitosamente' };
  } catch (error) {
    console.error('Error registrando producto:', error);
    return { success: false, error: error.toString() };
  }
}

function verificarCodigoBarras(codigoBarras) {
  try {
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var productoExistente = productos.find(function(p) {
      return p.Codigo_Barras === codigoBarras;
    });
    return productoExistente || null;
  } catch (error) {
    return null;
  }
}

function buscarProductos(termino) {
  try {
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    if (!termino) return productos;
    return productos.filter(function(producto) {
      var busqueda = termino.toLowerCase();
      return (
        (producto.Nombre_Producto && producto.Nombre_Producto.toLowerCase().includes(busqueda)) ||
        (producto.Codigo_Barras && producto.Codigo_Barras.toLowerCase().includes(busqueda)) ||
        (producto.ID_Producto && producto.ID_Producto.toLowerCase().includes(busqueda))
      );
    });
  } catch (error) {
    return [];
  }
}

function ajustarStockManual(ajusteData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var productosData = productosSheet.getDataRange().getValues();
    for (var i = 1; i < productosData.length; i++) {
      if (productosData[i][0] === ajusteData.idProducto) {
        var stockAnterior = parseInt(productosData[i][5]) || 0;
        var stockNuevo = parseInt(ajusteData.nuevoStock);
        productosSheet.getRange(i + 1, 6).setValue(stockNuevo);
        registrarMovimientoInventario(
          ajusteData.idProducto,
          stockNuevo - stockAnterior,
          stockAnterior,
          stockNuevo,
          ajusteData.idUsuario,
          'Ajuste manual: ' + ajusteData.motivo
        );
        return { success: true, stockAnterior: stockAnterior, stockNuevo: stockNuevo };
      }
    }
    return { success: false, error: 'Producto no encontrado' };
  } catch (error) {
    console.error('Error ajustando stock:', error);
    return { success: false, error: error.toString() };
  }
}

// ============== FUNCIÓN PARA REGISTRAR MOVIMIENTOS ==============
function registrarMovimientoInventario(productoId, cantidad, stockAnterior, stockNuevo, usuarioId, motivo) {
  try {
    console.log('📝 Registrando movimiento:', productoId, cantidad);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var movimientosSheet = ss.getSheetByName('MOVIMIENTO_INVENTARIO'); // Cambiado de INVENTARIO_MOVIMIENTOS a MOVIMIENTO_INVENTARIO
    
    // Generar ID único para el movimiento
    var movimientoId = 'MOV' + Utilities.getUuid().substring(0, 8);
    
    // Determinar tipo de movimiento
    var tipoMovimiento = cantidad >= 0 ? 'ENTRADA' : 'SALIDA';
    var idTipo = tipoMovimiento === 'ENTRADA' ? 'TIPO001' : 'TIPO002';
    var idMotivo = 'MOT001'; // Motivo por defecto
    
    // Registrar en la hoja de movimientos (nueva estructura)
    movimientosSheet.appendRow([
      movimientoId,
      new Date(),
      productoId,
      idTipo,
      idMotivo,
      cantidad,
      usuarioId,
      motivo // Referencia
    ]);
    
    console.log('✅ Movimiento registrado:', movimientoId);
    
  } catch (error) {
    console.error('❌ Error registrando movimiento:', error);
    // NO lanzar error para no interrumpir el flujo principal
  }
}

// === CATEGORÍAS (estático/fallback) ===
function getCategoriasProductos() {
  // Obtener categorías reales de la base de datos
  try {
    var categorias = getSheetData('CATEGORIA');
    if (categorias && categorias.length > 0) {
      return categorias.map(function(cat) {
        return cat.Nombre_Categoria;
      });
    }
  } catch (error) {
    console.error('Error obteniendo categorías:', error);
  }
  
  // Fallback a categorías por defecto
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

// ==================== DASHBOARD NUEVO ====================

// Parseo robusto de fechas (Date, texto o serial de Sheets)
function parseFecha(str) {
  if (!str) return null;
  try {
    if (Object.prototype.toString.call(str) === '[object Date]') return str;
    var d = new Date(str);
    if (!isNaN(d.getTime())) return d;
    if (!isNaN(parseFloat(str))) {
      var base = new Date(1899, 11, 30);
      base.setDate(base.getDate() + parseInt(str, 10));
      return base;
    }
  } catch(e) {}
  return null;
}

function _matchFiltroVenta(venta, filtros) {
  var fechaOk = true, categoriaOk = true, proveedorOk = true, clienteOk = true;

  // Rango de fechas
  if (filtros && (filtros.desde || filtros.hasta)) {
    var fv = parseFecha(venta.FechaHora_Venta); // Cambiado de Fecha_Venta a FechaHora_Venta
    if (!fv) return false;
    if (filtros.desde) {
      var fd = parseFecha(filtros.desde);
      if (fd && fv < fd) fechaOk = false;
    }
    if (filtros.hasta) {
      var fh = parseFecha(filtros.hasta);
      if (fh) {
        fh.setHours(23, 59, 59, 999);
        if (fv > fh) fechaOk = false;
      }
    }
  }

  if ((filtros && (filtros.categoria || filtros.proveedor)) || (filtros && filtros.cliente)) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var detSheet = ss.getSheetByName('DETALLE_VENTA'); // Cambiado de DETALLE_VENTAS a DETALLE_VENTA
    var prodSheet = ss.getSheetByName('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var detData = detSheet ? detSheet.getDataRange().getValues() : [];
    var prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];

    var prodMap = {};
    if (prodData.length > 1) {
      var headersP = prodData[0];
      for (var i = 1; i < prodData.length; i++) {
        var row = prodData[i];
        var obj = {};
        for (var j = 0; j < headersP.length; j++) obj[headersP[j]] = row[j];
        prodMap[obj.ID_Producto] = obj;
      }
    }

    var headersD = detData[0] || [];
    var idx = {
      ID_Detalle: headersD.indexOf('ID_Detalle'),
      ID_Venta: headersD.indexOf('ID_Venta'),
      ID_Producto: headersD.indexOf('ID_Producto'),
      Cantidad: headersD.indexOf('Cantidad'),
      Precio_Unitario: headersD.indexOf('Precio_Unitario') // Cambiado de Precio a Precio_Unitario
    };

    var tieneCategoria = false, tieneProveedor = false;
    if (detData.length > 1) {
      for (var k = 1; k < detData.length; k++) {
        if (detData[k][idx.ID_Venta] === venta.ID_Venta) {
          var pid = detData[k][idx.ID_Producto];
          var p = prodMap[pid];
          if (filtros.categoria) {
            // Para obtener categoría necesitarías unir con la tabla CATEGORIA
            // Por ahora asumimos que está en el producto
            if (p && p.ID_Categoria) {
              // Aquí podrías buscar el nombre de la categoría por ID
              tieneCategoria = true; // Simplificado por ahora
            }
          } else {
            tieneCategoria = true;
          }
          if (filtros.proveedor) {
            if (p && p.ID_Proveedor) {
              // Aquí podrías buscar el nombre del proveedor por ID
              tieneProveedor = true; // Simplificado por ahora
            }
          } else {
            tieneProveedor = true;
          }
        }
      }
    } else {
      if (!filtros.categoria) tieneCategoria = true;
      if (!filtros.proveedor) tieneProveedor = true;
    }

    categoriaOk = tieneCategoria;
    proveedorOk = tieneProveedor;

    if (filtros.cliente) {
      var c = (venta.ID_Cliente || '').toString().toLowerCase();
      clienteOk = c.indexOf(String(filtros.cliente).toLowerCase()) !== -1;
    }
  }

  return fechaOk && categoriaOk && proveedorOk && clienteOk;
}

function getDashboardData(filtros) {
  try {
    var productos = getSheetData('PRODUCTO'); // Cambiado de PRODUCTOS a PRODUCTO
    var ventas = getSheetData('VENTA'); // Cambiado de VENTAS a VENTA
    var detalles = getSheetData('DETALLE_VENTA'); // Cambiado de DETALLE_VENTAS a DETALLE_VENTA
    var usuarios = getSheetData('USUARIO'); // Cambiado de USUARIOS a USUARIO

    filtros = filtros || {};

    // ===== KPIs =====
    var hoy = new Date();
    var hoyStr = hoy.toDateString();

    var ventasDelDia = ventas.filter(function(v) {
      var fv = parseFecha(v.FechaHora_Venta); // Cambiado de Fecha_Venta a FechaHora_Venta
      return fv && fv.toDateString() === hoyStr && _matchFiltroVenta(v, filtros);
    });
    var montoDia = ventasDelDia.reduce(function(sum, v){
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);

    var y = hoy.getFullYear(), m = hoy.getMonth();
    var inicioMes = new Date(y, m, 1, 0, 0, 0);
    var finMes = new Date(y, m + 1, 0, 23, 59, 59);
    var ventasMes = ventas.filter(function(v){
      var fv = parseFecha(v.FechaHora_Venta); // Cambiado de Fecha_Venta a FechaHora_Venta
      if (!fv) return false;
      if (fv < inicioMes || fv > finMes) return false;
      return _matchFiltroVenta(v, filtros);
    });
    var ingresosMes = ventasMes.reduce(function(sum, v){
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);

    var stockTotal = productos.length;
    var porAgotarse = productos.filter(function(p){
      var sa = parseInt(p.Stock_Actual) || 0;
      var sm = parseInt(p.Stock_Minimo) || 5;
      return sa <= sm;
    }).length;

    var clientes = usuarios.filter(function(u){
      // En la nueva estructura, podrías verificar en HISTORIAL_ROL
      return true; // Simplificado por ahora
    }).length;

    var pendientes = ventas.filter(function(v){
      return _matchFiltroVenta(v, filtros) && String(v.Estado_Venta || '').toLowerCase().indexOf('pend') !== -1; // Cambiado de Estado a Estado_Venta
    }).length;

    var entregados = ventas.filter(function(v){
      var e = String(v.Estado_Venta || '').toLowerCase(); // Cambiado de Estado a Estado_Venta
      return _matchFiltroVenta(v, filtros) && (e.indexOf('complet') !== -1 || e.indexOf('entreg') !== -1);
    }).length;

    // ===== Gráfico: ventas por categoría =====
    var prodIndex = {};
    productos.forEach(function(p){ prodIndex[p.ID_Producto] = p; });

    var ventasPorCat = {};
    detalles.forEach(function(d){
      var v = ventas.find(function(vv){ return vv.ID_Venta === d.ID_Venta; });
      if (!v || !_matchFiltroVenta(v, filtros)) return;
      var p = prodIndex[d.ID_Producto];
      // Para categoría necesitarías unir con tabla CATEGORIA
      var cat = 'General'; // Simplificado por ahora
      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0; // Cambiado de Precio a Precio_Unitario
      var subtotal = (parseFloat(d.Subtotal) || (cant * precio));
      ventasPorCat[cat] = (ventasPorCat[cat] || 0) + subtotal;
    });
    var categoriasData = Object.keys(ventasPorCat).map(function(cat){
      return { categoria: cat, total: ventasPorCat[cat] };
    }).sort(function(a,b){ return b.total - a.total; });

    // ===== Gráfico: tendencia (7 o 30 días) =====
    var dias = filtros.frecuencia === 'mensual' ? 30 : 7;
    var hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var serie = [];
    for (var i= dias - 1; i >= 0; i--) {
      var d0 = new Date(hoy0); d0.setDate(d0.getDate() - i);
      var etiqueta = Utilities.formatDate(d0, 'GMT-5', 'dd/MM');
      var suma = ventas.filter(function(vt){
        var fv = parseFecha(vt.FechaHora_Venta); // Cambiado de Fecha_Venta a FechaHora_Venta
        return fv && fv.toDateString() === d0.toDateString() && _matchFiltroVenta(vt, filtros);
      }).reduce(function(ac, vt){ return ac + (parseFloat(vt.Total_Venta) || 0); }, 0);
      serie.push({ label: etiqueta, total: suma });
    }

    // ===== Gráfico: top productos (pie) =====
    var totalPorProducto = {};
    detalles.forEach(function(d){
      var v = ventas.find(function(vv){ return vv.ID_Venta === d.ID_Venta; });
      if (!v || !_matchFiltroVenta(v, filtros)) return;
      var p = prodIndex[d.ID_Producto];
      var nombre = p && p.Nombre_Producto ? p.Nombre_Producto : (d.ID_Producto || 'Desconocido');
      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0; // Cambiado de Precio a Precio_Unitario
      var subtotal = (parseFloat(d.Subtotal) || (cant * precio));
      totalPorProducto[nombre] = (totalPorProducto[nombre] || 0) + subtotal;
    });
    var topProdPie = Object.keys(totalPorProducto).map(function(n){
      return { nombre: n, total: totalPorProducto[n] };
    }).sort(function(a,b){ return b.total - a.total; }).slice(0, 8);

    // ===== Tablas =====
    var unidadesPorProd = {}, gananciaPorProd = {};
    detalles.forEach(function(d){
      var v = ventas.find(function(vv){ return vv.ID_Venta === d.ID_Venta; });
      if (!v || !_matchFiltroVenta(v, filtros)) return;
      var p = prodIndex[d.ID_Producto];
      var key = p && p.Nombre_Producto ? p.Nombre_Producto : (d.ID_Producto || 'Desconocido');
      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0; // Cambiado de Precio a Precio_Unitario
      var subtotal = (parseFloat(d.Subtotal) || (cant * precio));
      unidadesPorProd[key] = (unidadesPorProd[key] || 0) + cant;
      gananciaPorProd[key] = (gananciaPorProd[key] || 0) + subtotal;
    });

    var tablaMasVendidos = Object.keys(unidadesPorProd).map(function(n){
      var p = Object.values(prodIndex).find(function(pp){ return pp && pp.Nombre_Producto === n; });
      return {
        nombre: n,
        categoria: 'General', // Simplificado
        unidades: unidadesPorProd[n],
        ganancia: gananciaPorProd[n] || 0
      };
    }).sort(function(a,b){ return b.unidades - a.unidades; }).slice(0, 10);

    var tablaBajoStock = productos
      .map(function(p){
        var sa = parseInt(p.Stock_Actual) || 0;
        var sm = parseInt(p.Stock_Minimo) || 5;
        return {
          nombre: p.Nombre_Producto || 'Sin nombre',
          categoria: 'General', // Simplificado
          stock: sa,
          minimo: sm
        };
      })
      .filter(function(x){ return x.stock <= x.minimo; })
      .sort(function(a,b){ return a.stock - b.stock; })
      .slice(0, 10);

    var ultimosPedidos = ventas
      .filter(function(v){ return _matchFiltroVenta(v, filtros); })
      .map(function(v){
        return {
          fecha: v.FechaHora_Venta, // Cambiado de Fecha_Venta a FechaHora_Venta
          cliente: v.ID_Usuario || '—', // Usando ID_Usuario como cliente
          monto: parseFloat(v.Total_Venta) || 0,
          estado: v.Estado_Venta || '—', // Cambiado de Estado a Estado_Venta
          id: v.ID_Venta || '—'
        };
      })
      .sort(function(a,b){
        var da = parseFecha(a.fecha), db = parseFecha(b.fecha);
        return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
      })
      .slice(0, 10);

    return {
      kpis: {
        ventasDia: montoDia,
        ingresosMes: ingresosMes,
        productosEnStock: stockTotal,
        porAgotarse: porAgotarse,
        clientes: clientes,
        pedidosPendientes: pendientes,
        pedidosEntregados: entregados
      },
      charts: {
        porCategoria: categoriasData,
        tendencia: serie,
        topProductos: topProdPie
      },
      tablas: {
        masVendidos: tablaMasVendidos,
        bajoStock: tablaBajoStock,
        ultimosPedidos: ultimosPedidos
      }
    };

  } catch (error) {
    console.error('Error getDashboardData:', error);
    return {
      kpis: { ventasDia:0, ingresosMes:0, productosEnStock:0, porAgotarse:0, clientes:0, pedidosPendientes:0, pedidosEntregados:0 },
      charts: { porCategoria:[], tendencia:[], topProductos:[] },
      tablas: { masVendidos:[], bajoStock:[], ultimosPedidos:[] },
      error: error.toString()
    };
  }
}
