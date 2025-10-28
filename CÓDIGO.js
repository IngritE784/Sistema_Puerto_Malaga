// ================= Configuración =================
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

// ================ Estadísticas básicas (compat) ================
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

function getProductos() {
  return getSheetData('PRODUCTOS');
}

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

function getMovimientosInventario() {
  return getSheetData('INVENTARIO_MOVIMIENTOS');
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
    var productosSheet = ss.getSheetByName('PRODUCTOS');
    
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
        productoData.categoria || 'General',
        parseFloat(productoData.precio),
        parseInt(productoData.cantidad),
        parseInt(productoData.stockMinimo) || 5,
        productoData.proveedor || '',
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
    var productosSheet = ss.getSheetByName('PRODUCTOS');
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
    var lastRow = productosSheet.getLastRow();
    var newId = 'PRO' + String(lastRow).padStart(3, '0');
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
    var productos = getSheetData('PRODUCTOS');
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

function ajustarStockManual(ajusteData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTOS');
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
    var movimientosSheet = ss.getSheetByName('INVENTARIO_MOVIMIENTOS');
    
    // Generar ID único para el movimiento
    var movimientoId = 'MOV' + Utilities.getUuid().substring(0, 8);
    
    // Determinar tipo de movimiento
    var tipoMovimiento = cantidad >= 0 ? 'Entrada' : 'Salida';
    
    // Registrar en la hoja de movimientos
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
    
    console.log('✅ Movimiento registrado:', movimientoId);
    
  } catch (error) {
    console.error('❌ Error registrando movimiento:', error);
    // NO lanzar error para no interrumpir el flujo principal
  }
}

// === CATEGORÍAS (estático/fallback) ===
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
    var fv = parseFecha(venta.Fecha_Venta);
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
    var detSheet = ss.getSheetByName('DETALLE_VENTAS');
    var prodSheet = ss.getSheetByName('PRODUCTOS');
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
      Precio: headersD.indexOf('Precio')
    };

    var tieneCategoria = false, tieneProveedor = false;
    if (detData.length > 1) {
      for (var k = 1; k < detData.length; k++) {
        if (detData[k][idx.ID_Venta] === venta.ID_Venta) {
          var pid = detData[k][idx.ID_Producto];
          var p = prodMap[pid];
          if (filtros.categoria) {
            if (p && p.Categoria && String(p.Categoria).toLowerCase() === String(filtros.categoria).toLowerCase()) {
              tieneCategoria = true;
            }
          } else {
            tieneCategoria = true;
          }
          if (filtros.proveedor) {
            if (p && p.Proveedor && String(p.Proveedor).toLowerCase() === String(filtros.proveedor).toLowerCase()) {
              tieneProveedor = true;
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
      var c = (venta.ID_Cliente || venta.Cliente || '').toString().toLowerCase();
      clienteOk = c.indexOf(String(filtros.cliente).toLowerCase()) !== -1;
    }
  }

  return fechaOk && categoriaOk && proveedorOk && clienteOk;
}

function getDashboardData(filtros) {
  try {
    var productos = getSheetData('PRODUCTOS');
    var ventas = getSheetData('VENTAS');
    var detalles = getSheetData('DETALLE_VENTAS');
    var usuarios = getSheetData('USUARIOS');

    filtros = filtros || {};

    // ===== KPIs =====
    var hoy = new Date();
    var hoyStr = hoy.toDateString();

    var ventasDelDia = ventas.filter(function(v) {
      var fv = parseFecha(v.Fecha_Venta);
      return fv && fv.toDateString() === hoyStr && _matchFiltroVenta(v, filtros);
    });
    var montoDia = ventasDelDia.reduce(function(sum, v){
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);

    var y = hoy.getFullYear(), m = hoy.getMonth();
    var inicioMes = new Date(y, m, 1, 0, 0, 0);
    var finMes = new Date(y, m + 1, 0, 23, 59, 59);
    var ventasMes = ventas.filter(function(v){
      var fv = parseFecha(v.Fecha_Venta);
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
      return (u.Rol || '').toString().toLowerCase() === 'cliente';
    }).length;

    var pendientes = ventas.filter(function(v){
      return _matchFiltroVenta(v, filtros) && String(v.Estado || '').toLowerCase().indexOf('pend') !== -1;
    }).length;

    var entregados = ventas.filter(function(v){
      var e = String(v.Estado || '').toLowerCase();
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
      var cat = p && p.Categoria ? p.Categoria : 'Sin categoría';
      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio) || 0;
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
        var fv = parseFecha(vt.Fecha_Venta);
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
      var precio = parseFloat(d.Precio) || 0;
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
      var precio = parseFloat(d.Precio) || 0;
      var subtotal = (parseFloat(d.Subtotal) || (cant * precio));
      unidadesPorProd[key] = (unidadesPorProd[key] || 0) + cant;
      gananciaPorProd[key] = (gananciaPorProd[key] || 0) + subtotal;
    });

    var tablaMasVendidos = Object.keys(unidadesPorProd).map(function(n){
      var p = Object.values(prodIndex).find(function(pp){ return pp && pp.Nombre_Producto === n; });
      return {
        nombre: n,
        categoria: p && p.Categoria ? p.Categoria : '—',
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
          categoria: p.Categoria || '—',
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
          fecha: v.Fecha_Venta,
          cliente: v.Cliente || v.ID_Cliente || '—',
          monto: parseFloat(v.Total_Venta) || 0,
          estado: v.Estado || '—',
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
