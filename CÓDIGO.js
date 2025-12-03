// ================= Configuración =================
var SPREADSHEET_ID = '1VTF5ChP8eavortE2O8qzm3P3jZe5yRB5jfbKhhiyXs0';


// ================= Render HTML ===================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sistema Gestión Tienda Abarrotes');

    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');

}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ================= Utilidades base =================


// Lee una hoja por nombre y devuelve un array de objetos {header: valor}
function getSheetData(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];


    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];


    var headers = data[0];
    var jsonData = [];


    for (var i = 1; i < data.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = data[i][j];
        if (value === null || value === undefined) value = '';
        rowObj[header] = value.toString();
      }
      jsonData.push(rowObj);
    }
    return jsonData;
  } catch (error) {
    console.error('Error en getSheetData(' + sheetName + '): ' + error.toString());
    return [];
  }
}


// ==================== FECHAS ======================


// Parseo robusto de fechas (Date, texto o serial de Sheets)
function parseFecha(str) {
  if (!str) return null;
  try {
    // Si ya es Date
    if (Object.prototype.toString.call(str) === '[object Date]') {
      if (!isNaN(str.getTime())) return str;
    }


    // Si viene como string
    var d = new Date(str);
    if (!isNaN(d.getTime())) return d;


    // Si viene como número tipo serial de Excel/Sheets
    var num = parseFloat(str);
    if (!isNaN(num)) {
      var base = new Date(1899, 11, 30); // base Excel
      base.setDate(base.getDate() + parseInt(num, 10));
      return base;
    }
  } catch (e) {
    return null;
  }
  return null;
}


// ================= Autenticación =================


// Login de usuario: USUARIO + HISTORIAL_ROL (para obtener Rol activo)
function loginUsuario(email, password) {
  try {
    var usuarios = getSheetData('USUARIO');
    var roles = getSheetData('HISTORIAL_ROL');


    var usuario = usuarios.find(function (u) {
      var emailMatch = u.Email && u.Email.trim().toLowerCase() === email.trim().toLowerCase();
      var passMatch = u.Contraseña && u.Contraseña.trim() === password.trim();
      return emailMatch && passMatch;
    });


    if (!usuario) return null;


    // Buscar rol activo en HISTORIAL_ROL
    var rolesUsuario = roles.filter(function (r) {
      return r.ID_Usuario === usuario.ID_Usuario && (r.Estado || '').toLowerCase() === 'activo';
    });


    var rol = 'Sin rol';
    if (rolesUsuario.length > 0) {
      // Si hay varios, tomamos el más reciente por Fecha_Inicio
      rolesUsuario.sort(function (a, b) {
        var fa = parseFecha(a.Fecha_Inicio);
        var fb = parseFecha(b.Fecha_Inicio);
        return (fb ? fb.getTime() : 0) - (fa ? fa.getTime() : 0);
      });
      rol = rolesUsuario[0].Rol || 'Sin rol';
    }


    usuario.Rol = rol;
    return usuario;
  } catch (error) {
    console.error('Error en loginUsuario: ' + error.toString());
    return null;
  }
}


// ================ Estadísticas simples (compat) ================
function getEstadisticas() {
  try {
    var productos = getSheetData('PRODUCTO');
    var ventas = getSheetData('VENTA');


    var hoy = new Date();
    var hoyStr = hoy.toDateString();


    var ventasHoy = ventas.filter(function (venta) {
      if (!venta.FechaHora_Venta) return false;
      var f = parseFecha(venta.FechaHora_Venta);
      return f && f.toDateString() === hoyStr;
    });


    var totalVentasHoy = ventasHoy.reduce(function (sum, v) {
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);


    var productosBajoStock = productos.filter(function (p) {
      var sa = parseInt(p.Stock_Actual) || 0;
      var sm = parseInt(p.Stock_Minimo) || 5;
      return sa <= sm;
    });


    return {
      totalProductos: productos.length,
      totalVentasHoy: totalVentasHoy,
      alertasStock: productosBajoStock.length,
      productosBajoStock: productosBajoStock.length
    };
  } catch (e) {
    console.error('Error en getEstadisticas:', e);
    return {
      totalProductos: 0,
      totalVentasHoy: 0,
      alertasStock: 0,
      productosBajoStock: 0
    };
  }
}


// ================ Productos / Inventario ================


// Producto por código de barras o ID_Producto----------------------------------------------------------------------
function getProductoByCodigo(codigo) {
  try {
    console.log('Buscando producto con código:', codigo);
    
    var productos = getSheetData('PRODUCTO');
    console.log('Productos totales:', productos.length);
    
    var producto = productos.find(function(prod) {
      var codigoMatch = prod.Codigo_Barras && prod.Codigo_Barras.toString().trim() === codigo.toString().trim();
      var idMatch = prod.ID_Producto && prod.ID_Producto.toString().trim() === codigo.toString().trim();
      return codigoMatch || idMatch;
    });


    // Si está inactivo, se trata como no encontrado para ventas
    if (producto && String(producto.Estado || '').toLowerCase() === 'inactivo') {
      console.log('Producto encontrado pero INACTIVO:', producto.ID_Producto);
      return null;
    }
    
    console.log('Producto encontrado:', producto);
    return producto || null;
  } catch (error) {
    console.error('Error buscando producto: ' + error.toString());
    return null;
  }
}




// Devuelve productos enriquecidos con nombre de categoría y proveedor-----------------------------
function getProductos() {
  try {
    var productos = getSheetData('PRODUCTO');
    var categorias = getSheetData('CATEGORIA');
    var proveedores = getSheetData('PROVEEDOR');


    var catMap = {};
    categorias.forEach(function (c) {
      catMap[c.ID_Categoria] = c.Nombre_Categoria;
    });


    var provMap = {};
    proveedores.forEach(function (p) {
      p = p || {};
      provMap[p.ID_Proveedor] = p.Nombre_Proveedor;
    });


    return productos.map(function (p) {
      p.Categoria = catMap[p.ID_Categoria] || '';
      p.Proveedor = provMap[p.ID_Proveedor] || '';
      return p;
    });
  } catch (error) {
    console.error('Error en getProductos:', error);
    return [];
  }
}


// Alertas stock calculadas a partir de PRODUCTO
function getAlertasStock() {
  try {
    var productos = getSheetData('PRODUCTO');
    var alertas = productos
      .filter(function (p) {
        var sa = parseInt(p.Stock_Actual) || 0;
        var sm = parseInt(p.Stock_Minimo) || 5;
        return sa <= sm;
      })
      .map(function (p) {
        var sa = parseInt(p.Stock_Actual) || 0;
        var sm = parseInt(p.Stock_Minimo) || 5;
        return {
          Nombre_Producto: p.Nombre_Producto || 'Sin nombre',
          Stock_Actual: sa,
          Stock_Minimo: sm,
          Diferencia: sa - sm
        };
      });
    return alertas;
  } catch (e) {
    console.error('Error en getAlertasStock:', e);
    return [];
  }
}


// ====== Métodos auxiliares (categoría / proveedor / método pago) ======


function obtenerIdCategoriaPorNombre(nombreCategoria) {
  if (!nombreCategoria) return 'CAT001';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('CATEGORIA');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxId = headers.indexOf('ID_Categoria');
  var idxNombre = headers.indexOf('Nombre_Categoria');


  // Buscar existente
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[idxNombre]) continue;
    if (
      row[idxNombre].toString().trim().toLowerCase() ===
      nombreCategoria.toString().trim().toLowerCase()
    ) {
      return row[idxId];
    }
  }


  // Crear nueva
  var newId = 'CAT' + Utilities.formatString('%03d', Math.max(1, data.length));
  sheet.appendRow([newId, nombreCategoria, '', 'Activo']);
  return newId;
}


function obtenerIdProveedorPorNombre(nombreProveedor) {
  if (!nombreProveedor) return 'PROV001';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('PROVEEDOR');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxId = headers.indexOf('ID_Proveedor');
  var idxNombre = headers.indexOf('Nombre_Proveedor');


  // Buscar existente
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[idxNombre]) continue;
    if (
      row[idxNombre].toString().trim().toLowerCase() ===
      nombreProveedor.toString().trim().toLowerCase()
    ) {
      return row[idxId];
    }
  }


  // Crear nuevo proveedor
  var newId = 'PROV' + Utilities.formatString('%03d', Math.max(1, data.length));
  sheet.appendRow([newId, nombreProveedor, '', '', '', 'Activo']);
  return newId;
}


function obtenerIdMetodoPago(nombreMetodo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('METODO_PAGO');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('ID_Metodo');
    var idxNombre = headers.indexOf('Nombre_Metodo');


    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[idxNombre]) continue;
      if (row[idxNombre].toString().trim() === nombreMetodo.toString().trim()) {
        return row[idxId];
      }
    }
  } catch (e) {
    console.error('Error en obtenerIdMetodoPago:', e);
  }
  // Si no hay nada configurado, devolvemos un ID por defecto
  return 'MP001';
}


// Registrar movimiento en MOVIMIENTO_INVENTARIO
function registrarMovimientoInventario(productoId, cantidad, usuarioId, motivo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MOVIMIENTO_INVENTARIO');


    var movimientoId = 'MOV' + Utilities.getUuid().substring(0, 8);


    var tipoMovimiento = cantidad >= 0 ? 'ENTRADA' : 'SALIDA';
    var idTipo = tipoMovimiento === 'ENTRADA' ? 'TIPO001' : 'TIPO002';
    var idMotivo = 'MOT001'; // motivo genérico


    sheet.appendRow([
      movimientoId,
      new Date(),
      productoId,
      idTipo,
      idMotivo,
      cantidad,
      usuarioId,
      motivo || ''
    ]);
  } catch (e) {
    console.error('Error en registrarMovimientoInventario:', e);
  }
}


// Actualizar stock de un producto y registrar movimiento
function actualizarStockProducto(productoId, deltaCantidad, motivo, usuarioId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('PRODUCTO');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];


    var idxId = headers.indexOf('ID_Producto');
    var idxStock = headers.indexOf('Stock_Actual');


    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[idxId] === productoId) {
        var stockAnterior = parseInt(row[idxStock]) || 0;
        var stockNuevo = stockAnterior + (parseInt(deltaCantidad) || 0);
        if (stockNuevo < 0) stockNuevo = 0;
        sheet.getRange(i + 1, idxStock + 1).setValue(stockNuevo);


        registrarMovimientoInventario(
          productoId,
          deltaCantidad,
          usuarioId,
          motivo || ('Ajuste stock ' + productoId)
        );
        break;
      }
    }
  } catch (e) {
    console.error('Error en actualizarStockProducto:', e);
  }
}


// Devolver movimientos de inventario (para la vista)
function getMovimientosInventario() {
  try {
    // 1. Obtener los datos base
    var movimientos = getSheetData('MOVIMIENTO_INVENTARIO');
    var productos = getSheetData('PRODUCTO');
    var usuarios = getSheetData('USUARIO'); // <-- Necesario para el nombre
    var tiposMov = getSheetData('TIPO_MOVIMIENTO');
    var motivos = getSheetData('MOTIVO_MOVIMIENTO');

    // 2. Crear mapas de búsqueda (ID -> Nombre)
    var prodMap = {};
    productos.forEach(function(p) {
      prodMap[p.ID_Producto] = p.Nombre_Producto || null; // Usa null si está vacío
    });

    var userMap = {};
    usuarios.forEach(function(u) {
      userMap[u.ID_Usuario] = u.Nombre || u.Email || null; // Usa null si está vacío
    });

    var tipoMap = {};
    tiposMov.forEach(function(t) {
      tipoMap[t.ID_Tipo] = t.Nombre_Tipo || null;
    });

    var motivoMap = {};
    motivos.forEach(function(m) {
      motivoMap[m.ID_Motivo] = m.Nombre_Motivo || null;
    });

    // 3. Enriquecer los datos de movimientos
    var movimientosEnriquecidos = movimientos.map(function(m) {
      var mov = JSON.parse(JSON.stringify(m)); 
      
      // Asigna el nombre o deja 'undefined' (que el HTML ahora manejará)
      mov.Nombre_Producto = prodMap[m.ID_Producto];
      mov.Nombre_Usuario = userMap[m.ID_Usuario]; 
      mov.Nombre_Tipo = tipoMap[m.ID_Tipo_Movimiento];
      mov.Nombre_Motivo = motivoMap[m.ID_Motivo];
      
      return mov;
    });
    
    // 4. Devolver la lista ordenada
    return movimientosEnriquecidos.sort(function(a, b) {
        var da = parseFecha(a.Fecha_Movimiento);
        var db = parseFecha(b.Fecha_Movimiento);
        return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });

  } catch (e) {
    console.error('Error en getMovimientosInventario (enriquecido):', e);
    return [];
  }
}


// ============== GESTIÓN DE PRODUCTOS ==============


// Entrada de mercancía (sumar stock a producto existente o crear uno nuevo)
function registrarEntradaMercancia(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO');
    var data = productosSheet.getDataRange().getValues();
    var headers = data[0];


    var idxId = headers.indexOf('ID_Producto');
    var idxCodigo = headers.indexOf('Codigo_Barras');
    var idxPrecio = headers.indexOf('Precio');
    var idxStock = headers.indexOf('Stock_Actual');
    var idxStockMin = headers.indexOf('Stock_Minimo');
    var idxIdCat = headers.indexOf('ID_Categoria');
    var idxIdProv = headers.indexOf('ID_Proveedor');
    var idxEstado = headers.indexOf('Estado');


    var productoExistente = null;
    var filaProducto = -1;


    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[idxId] === productoData.idProducto || row[idxCodigo] === productoData.codigoBarras) {
        productoExistente = row;
        filaProducto = i + 1;
        break;
      }
    }


    var cantidad = parseInt(productoData.cantidad) || 0;
    var usuarioId = productoData.idUsuario;


    if (productoExistente) {
      var stockAnterior = parseInt(productoExistente[idxStock]) || 0;
      var stockNuevo = stockAnterior + cantidad;
      productosSheet.getRange(filaProducto, idxStock + 1).setValue(stockNuevo);


      registrarMovimientoInventario(
        productoData.idProducto,
        cantidad,
        usuarioId,
        'Entrada de mercancía: ' + (productoData.motivo || '')
      );


      return {
        success: true,
        tipo: 'actualizado',
        stockAnterior: stockAnterior,
        stockNuevo: stockNuevo
      };
    } else {
      var lastRow = productosSheet.getLastRow();
      var newId = 'PRO' + Utilities.formatString('%03d', lastRow);


      var idCategoria = obtenerIdCategoriaPorNombre(productoData.categoria || 'General');
      var idProveedor = obtenerIdProveedorPorNombre(productoData.proveedor || '');


      productosSheet.appendRow([
        newId,
        productoData.codigoBarras,
        productoData.nombre,
        idCategoria,
        parseFloat(productoData.precio) || 0,
        cantidad,
        parseInt(productoData.stockMinimo) || 5,
        idProveedor,
        new Date(),
        'Activo'
      ]);


      registrarMovimientoInventario(
        newId,
        cantidad,
        usuarioId,
        'Nuevo producto: ' + (productoData.motivo || '')
      );


      return { success: true, tipo: 'nuevo', productoId: newId };
    }
  } catch (error) {
    console.error('Error registrando entrada:', error);
    return { success: false, error: error.toString() };
  }
}


// Registrar producto nuevo por código de barras
function registrarProductoPorCodigo(productoData) {
  try {
    var productos = getSheetData('PRODUCTO');
    var existente = productos.find(function (p) {
      return p.Codigo_Barras === productoData.codigoBarras;
    });
    if (existente) {
      return {
        success: false,
        error: 'El código de barras ya existe para el producto: ' + existente.Nombre_Producto
      };
    }


    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var productosSheet = ss.getSheetByName('PRODUCTO');
    var lastRow = productosSheet.getLastRow();
    var newId = 'PRO' + Utilities.formatString('%03d', lastRow);

    // MODIFICACION --------------------------------------------------------------------------------------------------------------------------
    // Se asume que productoData.categoria YA es un ID_Categoria que viene del <select> del frontend
    var idCategoria = productoData.categoria || 'CAT001'; 
    var idProveedor = 'PROV001'; // Se deja un proveedor por defecto
    
    var stockInicial = parseInt(productoData.stockInicial) || 0;
    var stockMinimo = parseInt(productoData.stockMinimo) || 5;


    productosSheet.appendRow([
      newId,
      productoData.codigoBarras,
      productoData.nombre,
      idCategoria,
      parseFloat(productoData.precio),
      stockInicial,
      stockMinimo,
      idProveedor, 
      new Date(),
      'Activo'
    ]);




    if (stockInicial > 0) {
      registrarMovimientoInventario(
        newId,
        stockInicial,
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


// Verificar si existe código de barras
function verificarCodigoBarras(codigoBarras) {
  try {
    var productos = getSheetData('PRODUCTO');
    var p = productos.find(function (prod) {
      return prod.Codigo_Barras === codigoBarras;
    });
    return p || null;
  } catch (e) {
    console.error('Error en verificarCodigoBarras:', e);
    return null;
  }
}


// Búsqueda de productos (para listado)
function buscarProductos(termino) {
  try {
    var productos = getProductos(); // ya viene enriquecido
    if (!termino) return productos;
    var busq = termino.toLowerCase();
    return productos.filter(function (p) {
      return (
        (p.Nombre_Producto && p.Nombre_Producto.toLowerCase().includes(busq)) ||
        (p.Codigo_Barras && p.Codigo_Barras.toLowerCase().includes(busq)) ||
        (p.ID_Producto && p.ID_Producto.toLowerCase().includes(busq)) ||
        (p.Categoria && p.Categoria.toLowerCase().includes(busq)) ||
        (p.Proveedor && p.Proveedor.toLowerCase().includes(busq))
      );
    });
  } catch (e) {
    console.error('Error en buscarProductos:', e);
    return [];
  }
}


// Ajuste manual de stock
function ajustarStockManual(ajusteData) {
  try {
    var nuevoStock = parseInt(ajusteData.nuevoStock);
    if (isNaN(nuevoStock) || nuevoStock < 0) {
      return { success: false, error: 'Stock inválido' };
    }


    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('PRODUCTO');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];


    var idxId = headers.indexOf('ID_Producto');
    var idxStock = headers.indexOf('Stock_Actual');


    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[idxId] === ajusteData.idProducto) {
        var stockAnterior = parseInt(row[idxStock]) || 0;
        sheet.getRange(i + 1, idxStock + 1).setValue(nuevoStock);


        var delta = nuevoStock - stockAnterior;
        registrarMovimientoInventario(
          ajusteData.idProducto,
          delta,
          ajusteData.idUsuario,
          'Ajuste manual: ' + (ajusteData.motivo || '')
        );


        return {
          success: true,
          stockAnterior: stockAnterior,
          stockNuevo: nuevoStock
        };
      }
    }
    return { success: false, error: 'Producto no encontrado' };
  } catch (error) {
    console.error('Error ajustando stock:', error);
    return { success: false, error: error.toString() };
  }
}
// agregado-------------------------------------------------------------------------


function editarProducto(productoData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('PRODUCTO');
    var data = sheet.getDataRange().getValues();


    var fila = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === productoData.idProducto) { // ID_Producto en col 1
        fila = i + 1; // índice de fila en hoja
        break;
      }
    }


    if (fila === -1) {
      return { success: false, error: 'Producto no encontrado' };
    }


    // Validar código de barras único (si viene informado)
    var nuevoCodigo = (productoData.codigoBarras || '').toString().trim();
    if (nuevoCodigo) {
      for (var j = 1; j < data.length; j++) {
        if (j === fila - 1) continue; // saltar la misma fila
        var codigoExistente = (data[j][1] || '').toString().trim(); // col 2: Codigo_Barras
        if (codigoExistente && codigoExistente === nuevoCodigo) {
          return { success: false, error: 'El código de barras ya está usado por otro producto' };
        }
      }
      sheet.getRange(fila, 2).setValue(nuevoCodigo); // col 2
    }


    // Actualizar nombre, precio, stock mínimo y estado
    sheet.getRange(fila, 3).setValue(productoData.nombre); // Nombre_Producto (col 3)
    sheet.getRange(fila, 5).setValue(parseFloat(productoData.precio) || 0); // Precio (col 5)
    sheet.getRange(fila, 7).setValue(parseInt(productoData.stockMinimo) || 0); // Stock_Minimo (col 7)


    if (productoData.estado) {
      sheet.getRange(fila, 10).setValue(productoData.estado); // Estado (col 10)
    }


    return { success: true };
  } catch (error) {
    console.error('Error en editarProducto:', error);
    return { success: false, error: error.toString() };
  }
}






function desactivarProducto(idProducto) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('PRODUCTO');
    var data = sheet.getDataRange().getValues();


    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idProducto) { // ID_Producto
        sheet.getRange(i + 1, 10).setValue('Inactivo'); // Estado en col 10
        return { success: true };
      }
    }


    return { success: false, error: 'Producto no encontrado' };
  } catch (error) {
    console.error('Error en desactivarProducto:', error);
    return { success: false, error: error.toString() };
  }
}




// Categorías: devuelve nombres; si falla, fallback a lista fija 
function getCategoriasProductos() {
  try {
    var categorias = getSheetData('CATEGORIA');
    if (categorias && categorias.length > 0) {
      // Devuelve objetos: {id, nombre}
      return categorias.map(function(cat) {
        return {
          id: cat.ID_Categoria,
          nombre: cat.Nombre_Categoria
        };
      });
    }
  } catch (error) {
    console.error('Error obteniendo categorías:', error);
  }
  
  // Fallback si no hay hoja o está vacía
  return [
    { id: 'CATGEN', nombre: 'General' },
    { id: 'CATDESP', nombre: 'Despensa' },
    { id: 'CATLACT', nombre: 'Lácteos y huevos' },
    { id: 'CATCARN', nombre: 'Carnes y embutidos' },
    { id: 'CATBEB',  nombre: 'Bebidas' }
  ];
}


// NUEVA FUNCION AÑADIDA------------------------------------------------------------------------------------------------------
function crearCategoria(nombreCategoria) {
  try {
    if (!nombreCategoria) {
      throw new Error('Nombre de categoría vacío');
    }


    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('CATEGORIA');
    if (!sheet) {
      throw new Error('No se encontró la hoja CATEGORIA');
    }


    // Leemos headers para respetar el orden de columnas
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];


    var idxId      = headers.indexOf('ID_Categoria');
    var idxNombre  = headers.indexOf('Nombre_Categoria');
    var idxEstado  = headers.indexOf('Estado');
    var idxFecha   = headers.indexOf('Fecha_Creacion');


    var lastRow = sheet.getLastRow();
    var newId = 'CAT' + String(lastRow).padStart(3, '0');


    var nuevaFila = new Array(headers.length).fill('');


    if (idxId      >= 0) nuevaFila[idxId]     = newId;
    if (idxNombre  >= 0) nuevaFila[idxNombre] = nombreCategoria;
    if (idxEstado  >= 0) nuevaFila[idxEstado] = 'Activo';
    if (idxFecha   >= 0) nuevaFila[idxFecha]  = new Date();


    sheet.appendRow(nuevaFila);


    return {
      success: true,
      idCategoria: newId,
      nombre: nombreCategoria
    };


  } catch (error) {
    console.error('Error crearCategoria:', error);
    return { success: false, error: error.toString() };
  }
}






// ============== REGISTRO DE VENTAS ==============


function registrarVenta(ventaData, carrito) {
  try {
    if (!carrito || carrito.length === 0) {
      throw new Error('El carrito está vacío');
    }


    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ventasSheet = ss.getSheetByName('VENTA');
    var detalleSheet = ss.getSheetByName('DETALLE_VENTA');
    var ventaMetodoSheet = ss.getSheetByName('VENTA_METODO_PAGO');


    var lastRowVentas = ventasSheet.getLastRow();
    var newVentaId = 'VTA' + Utilities.formatString('%03d', lastRowVentas);


    // Numero_Ticket simple
    var numeroTicket = 'T' + Utilities.formatString('%03d', lastRowVentas);


    // Registrar venta
    ventasSheet.appendRow([
      newVentaId,             // ID_Venta
      new Date(),             // FechaHora_Venta
      ventaData.idCajero,     // ID_Usuario
      'TUR001',               // ID_Turno (por ahora fijo)
      parseFloat(ventaData.total) || 0, // Total_Venta
      'Completada',           // Estado_Venta
      numeroTicket,           // Numero_Ticket
      ''                      // Columna 1 (vacío por ahora)
    ]);


    // Detalles de venta
    carrito.forEach(function (item) {
      detalleSheet.appendRow([
        'DET' + Utilities.getUuid().substring(0, 8), // ID_Detalle
        newVentaId,                     // ID_Venta
        item.idProducto,                // ID_Producto
        parseInt(item.cantidad) || 0,   // Cantidad
        parseFloat(item.precio) || 0,   // Precio_Unitario
        parseFloat(item.subtotal) || 0  // Subtotal
      ]);


      // Actualizar stock
      actualizarStockProducto(
        item.idProducto,
        -1 * (parseInt(item.cantidad) || 0),
        'Venta: ' + newVentaId,
        ventaData.idCajero
      );
    });


    // Registrar método de pago
    var idMetodo = obtenerIdMetodoPago(ventaData.metodoPago || 'Efectivo');
    ventaMetodoSheet.appendRow([
      'VMP' + Utilities.getUuid().substring(0, 8),
      newVentaId,
      idMetodo,
      parseFloat(ventaData.total) || 0
    ]);


    return newVentaId;
  } catch (error) {
    console.error('Error registrando venta:', error);
    throw new Error('No se pudo registrar la venta: ' + error.toString());
  }
}

// ============== DASHBOARD NUEVO ==============


/**
 * Función auxiliar REFACTORIZADA para verificar si una venta cumple con los filtros.
 * Esta versión es EFICIENTE: no accede al Spreadsheet.
 * Recibe los datos (detalles y productos) como parámetros.
 */
function _matchFiltroVenta(venta, filtros, detallesVenta, prodMap) {
  var fechaOk = true, categoriaOk = true, proveedorOk = true;


  // 1. Filtro de Rango de fechas
  if (filtros && (filtros.desde || filtros.hasta)) {
    var fv = parseFecha(venta.FechaHora_Venta);
    if (!fv) return false; // Si la venta no tiene fecha válida, no pasa el filtro
    
    if (filtros.desde) {
      var fd = parseFecha(filtros.desde);
      if (fd && fv < fd) fechaOk = false;
    }
    if (filtros.hasta) {
      var fh = parseFecha(filtros.hasta);
      if (fh) {
        fh.setHours(23, 59, 59, 999); // Asegurar que incluya todo el día "hasta"
        if (fv > fh) fechaOk = false;
      }
    }
  }
  
  // Si no pasa el filtro de fecha, no seguir
  if (!fechaOk) return false;


  // 2. Filtros de Categoría o Proveedor
  // Si no hay filtro de categoría NI de proveedor, la venta pasa.
  if (!filtros.categoria && !filtros.proveedor) {
    return true; 
  }


  // Si hay filtros, asumimos que no cumple hasta encontrar un producto que sí
  categoriaOk = !filtros.categoria; // Si no hay filtro de cat, es true
  proveedorOk = !filtros.proveedor; // Si no hay filtro de prov, es true


  // Buscar los detalles (productos) que pertenecen a ESTA venta
  var detallesDeEstaVenta = detallesVenta.filter(function(d) {
    return d.ID_Venta === venta.ID_Venta;
  });


  // Si la venta no tiene detalles, no puede cumplir filtros de producto
  if (detallesDeEstaVenta.length === 0) {
      // Si el filtro de cat o prov existe, la venta no pasa
      if (filtros.categoria || filtros.proveedor) return false;
  }


  // Revisar cada producto de la venta
  for (var i = 0; i < detallesDeEstaVenta.length; i++) {
    var d = detallesDeEstaVenta[i];
    var p = prodMap[d.ID_Producto]; // Obtener info del producto desde el mapa
    
    if (p) {
      // Chequear filtro categoría
      if (filtros.categoria && p.ID_Categoria === filtros.categoria) {
        categoriaOk = true;
      }
      // Chequear filtro proveedor
      if (filtros.proveedor && p.ID_Proveedor === filtros.proveedor) {
        proveedorOk = true;
      }
    }
    
    // Si ya cumplió ambos, no seguir iterando
    if (categoriaOk && proveedorOk) break;
  }


  return fechaOk && categoriaOk && proveedorOk;
}




function getDashboardData(filtros) {
  try {
    filtros = filtros || {};


    var productos = getSheetData('PRODUCTO');
    var ventas = getSheetData('VENTA');
    var detalles = getSheetData('DETALLE_VENTA'); // <_ Se pasa a _matchFiltroVenta
    var categorias = getSheetData('CATEGORIA');
    var usuarios = getSheetData('USUARIO');
    var metodosPago = getSheetData('METODO_PAGO');
    var ventasMetodos = getSheetData('VENTA_METODO_PAGO');


    // Mapas auxiliares
    var prodMap = {}; // <_ Se pasa a _matchFiltroVenta
    productos.forEach(function (p) {
      prodMap[p.ID_Producto] = p;
    });


    var catMap = {};
    categorias.forEach(function (c) {
      catMap[c.ID_Categoria] = c.Nombre_Categoria;
    });

    var metodoMap = {};
    metodosPago.forEach(function(mp) {
      metodoMap[mp.ID_Metodo] = mp.Nombre_Metodo;
    });


    var hoy = new Date();
    var hoyStr = hoy.toDateString();


    // === KPIs ===


    // Ventas del día
    var ventasDiaArr = ventas.filter(function (v) {
      var fv = parseFecha(v.FechaHora_Venta);
      // Filtro optimizado
      return fv && fv.toDateString() === hoyStr && _matchFiltroVenta(v, filtros, detalles, prodMap);
    });
    var ventasDia = ventasDiaArr.reduce(function (sum, v) {
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);


    // Ingresos del mes actual
    var y = hoy.getFullYear();
    var m = hoy.getMonth();
    var inicioMes = new Date(y, m, 1, 0, 0, 0);
    var finMes = new Date(y, m + 1, 0, 23, 59, 59);


    var ventasMes = ventas.filter(function (v) {
      var fv = parseFecha(v.FechaHora_Venta);
      if (!fv) return false;
      if (fv < inicioMes || fv > finMes) return false;
      // Filtro optimizado
      return _matchFiltroVenta(v, filtros, detalles, prodMap);
    });


    var ingresosMes = ventasMes.reduce(function (sum, v) {
      return sum + (parseFloat(v.Total_Venta) || 0);
    }, 0);


    // Productos en stock y por agotarse
    var productosEnStock = productos.length;
    var porAgotarse = productos.filter(function (p) {
      var sa = parseInt(p.Stock_Actual) || 0;
      var sm = parseInt(p.Stock_Minimo) || 5;
      return sa <= sm;
    }).length;


    // Clientes registrados (usuarios con rol "Cliente" en HISTORIAL_ROL)
    var roles = getSheetData('HISTORIAL_ROL');
    var clientesIds = {};
    roles.forEach(function (r) {
      if (
        (r.Rol || '').toLowerCase().indexOf('client') !== -1 &&
        (r.Estado || '').toLowerCase() === 'activo'
      ) {
        clientesIds[r.ID_Usuario] = true;
      }
    });
    var clientesRegistrados = Object.keys(clientesIds).length;


    // Pedidos pendientes / entregados (por Estado_Venta)
    var pendientes = ventas.filter(function (v) {
      // Filtro optimizado
      if (!_matchFiltroVenta(v, filtros, detalles, prodMap)) return false;
      var est = (v.Estado_Venta || '').toLowerCase();
      return est.indexOf('pend') !== -1;
    }).length;


    var entregados = ventas.filter(function (v) {
      // Filtro optimizado
      if (!_matchFiltroVenta(v, filtros, detalles, prodMap)) return false;
      var est = (v.Estado_Venta || '').toLowerCase();
      return est.indexOf('complet') !== -1 || est.indexOf('entreg') !== -1;
    }).length;


    // === Gráfico: ventas por categoría ===
    var ventasPorCat = {};
    detalles.forEach(function (d) {
      var v = ventas.find(function (vv) {
        return vv.ID_Venta === d.ID_Venta;
      });
      // Filtro optimizado
      if (!v || !_matchFiltroVenta(v, filtros, detalles, prodMap)) return;


      var p = prodMap[d.ID_Producto];
      var idCat = p ? p.ID_Categoria : null;
      var nombreCat = idCat ? catMap[idCat] || 'Sin categoría' : 'Sin categoría';


      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0;
      var subtotal = parseFloat(d.Subtotal) || cant * precio;


      ventasPorCat[nombreCat] = (ventasPorCat[nombreCat] || 0) + subtotal;
    });


    var categoriasData = Object.keys(ventasPorCat)
      .map(function (c) {
        return { categoria: c, total: ventasPorCat[c] };
      })
      .sort(function (a, b) {
        return b.total - a.total;
      });


    // === Gráfico: tendencia (7 o 30 días) ===
    var dias = filtros.frecuencia === 'mensual' ? 30 : 7;
    var hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var serie = [];
    for (var i = dias - 1; i >= 0; i--) {
      var d0 = new Date(hoy0);
      d0.setDate(d0.getDate() - i);
      var label =
        Utilities.formatDate(d0, Session.getScriptTimeZone(), 'dd/MM');


      var totalDia = ventas
        .filter(function (v) {
          var fv = parseFecha(v.FechaHora_Venta);
          // Filtro optimizado
          return fv && fv.toDateString() === d0.toDateString() && _matchFiltroVenta(v, filtros, detalles, prodMap);
        })
        .reduce(function (sum, v) {
          return sum + (parseFloat(v.Total_Venta) || 0);
        }, 0);


      serie.push({ label: label, total: totalDia });
    }


    // === Gráfico: top productos (pie) ===
    var totalPorProducto = {};
    detalles.forEach(function (d) {
      var v = ventas.find(function (vv) {
        return vv.ID_Venta === d.ID_Venta;
      });
      // Filtro optimizado
      if (!v || !_matchFiltroVenta(v, filtros, detalles, prodMap)) return;


      var p = prodMap[d.ID_Producto];
      var nombre = p && p.Nombre_Producto ? p.Nombre_Producto : (d.ID_Producto || 'Desconocido');


      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0;
      var subtotal = parseFloat(d.Subtotal) || cant * precio;


      totalPorProducto[nombre] = (totalPorProducto[nombre] || 0) + subtotal;
    });


    var topProdPie = Object.keys(totalPorProducto)
      .map(function (n) {
        return { nombre: n, total: totalPorProducto[n] };
      })
      .sort(function (a, b) {
        return b.total - a.total;
      })
      .slice(0, 8);


    // === Tablas ===

    // === Resumen Métodos de Pago ===
    var resumenPagos = {};
    // Re-filtramos ventas (esto es ineficiente, pero sigue tu lógica actual)
    var ventasFiltradasGeneral = ventas.filter(function(v) { 
      // Filtro optimizado
      return _matchFiltroVenta(v, filtros, detalles, prodMap); 
    });
    var ventasFiltradasIDs = {};
    ventasFiltradasGeneral.forEach(function(v) { ventasFiltradasIDs[v.ID_Venta] = true; });

    ventasMetodos.forEach(function(vm) {
      if (ventasFiltradasIDs[vm.ID_Venta]) {
        var nombreMetodo = metodoMap[vm.ID_Metodo] || 'Desconocido';
        // Tu hoja VENTA_METODO_PAGO puede tener 'Monto' o 'Monto_Pagado'
        var monto = parseFloat(vm.Monto_Pagado) || parseFloat(vm.Monto) || 0; 
        resumenPagos[nombreMetodo] = (resumenPagos[nombreMetodo] || 0) + monto;
      }
    });

    var tablaResumenPagos = Object.keys(resumenPagos).map(function(nombre) {
      return { metodo: nombre, total: resumenPagos[nombre] };
    });


    // Más vendidos
    var unidadesPorProd = {};
    var gananciaPorProd = {};


    detalles.forEach(function (d) {
      var v = ventas.find(function (vv) {
        return vv.ID_Venta === d.ID_Venta;
      });
      // Filtro optimizado
      if (!v || !_matchFiltroVenta(v, filtros, detalles, prodMap)) return;


      var p = prodMap[d.ID_Producto];
      var nombre = p && p.Nombre_Producto ? p.Nombre_Producto : (d.ID_Producto || 'Desconocido');
      var idCat = p ? p.ID_Categoria : null;
      var nombreCat = idCat ? catMap[idCat] || 'Sin categoría' : 'Sin categoría';


      var cant = parseFloat(d.Cantidad) || 0;
      var precio = parseFloat(d.Precio_Unitario) || 0;
      var subtotal = parseFloat(d.Subtotal) || cant * precio;


      unidadesPorProd[nombre] = (unidadesPorProd[nombre] || 0) + cant;
      gananciaPorProd[nombre] = (gananciaPorProd[nombre] || 0) + subtotal;


      // Guardamos categoría en el propio objeto producto para luego
      p = p || {};
      p._CategoriaDashboard = nombreCat;
      prodMap[d.ID_Producto] = p;
    });


    var tablaMasVendidos = Object.keys(unidadesPorProd)
      .map(function (nombre) {
        // buscar producto para recuperar categoría
        var cat = 'General';
        for (var id in prodMap) {
          var pp = prodMap[id];
          if (pp && pp.Nombre_Producto === nombre && pp._CategoriaDashboard) {
            cat = pp._CategoriaDashboard;
            break;
          }
        }
        return {
          nombre: nombre,
          categoria: cat,
          unidades: unidadesPorProd[nombre],
          ganancia: gananciaPorProd[nombre] || 0
        };
      })
      .sort(function (a, b) {
        return b.unidades - a.unidades;
      })
      .slice(0, 10);


    // Bajo stock
    var tablaBajoStock = productos
      .map(function (p) {
        var sa = parseInt(p.Stock_Actual) || 0;
        var sm = parseInt(p.Stock_Minimo) || 5;
        var nombreCat = p.ID_Categoria ? catMap[p.ID_Categoria] || 'Sin categoría' : 'Sin categoría';
        return {
          nombre: p.Nombre_Producto || 'Sin nombre',
          categoria: nombreCat,
          stock: sa,
          minimo: sm
        };
      })
      .filter(function (x) {
        return x.stock <= x.minimo;
      })
      .sort(function (a, b) {
        return a.stock - b.stock;
      })
      .slice(0, 10);


    // Últimos pedidos (ventas recientes)
  var ultimosPedidos = ventas
  .filter(function(v) { 
    // Filtro optimizado
    return _matchFiltroVenta(v, filtros, detalles, prodMap); 
  })
  .map(function(v) {
    return {
      fecha: v.FechaHora_Venta,
      monto: parseFloat(v.Total_Venta) || 0,
      estado: v.Estado_Venta || '—',
      id: v.ID_Venta || '—'
    };
  })
  .sort(function(a, b) {
    var da = parseFecha(a.fecha), db = parseFecha(b.fecha);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  })
  .slice(0, 10);






    return {
      kpis: {
        ventasDia: ventasDia,
        ingresosMes: ingresosMes,
        productosEnStock: productosEnStock,
        porAgotarse: porAgotarse,
        clientes: clientesRegistrados,
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
        ultimosPedidos: ultimosPedidos,
        resumenPagos: tablaResumenPagos
      }
    };
  } catch (error) {
    console.error('Error en getDashboardData:', error);
    return {
      kpis: {
        ventasDia: 0,
        ingresosMes: 0,
        productosEnStock: 0,
        porAgotarse: 0,
        clientes: 0,
        pedidosPendientes: 0,
        pedidosEntregados: 0
      },
      charts: { porCategoria: [], tendencia: [], topProductos: [] },
      tablas: { masVendidos: [], bajoStock: [], ultimosPedidos: [], resumenPagos: [] },
      error: error.toString()
    };
  }
}


//FUNCION AGREGADA-------------------------------------------------------------------------------------------------


function generarReporteDashboard(filtros) {
  try {
    filtros = filtros || {};
    var datos = getDashboardData(filtros);  // Reutilizamos tu dashboard


    var hoy = new Date();
    var zona = 'America/Lima';
    var nombreArchivoBase = 'Reporte_Tienda_' + Utilities.formatDate(hoy, zona, 'yyyyMMdd_HHmm');


    // 1) Crear documento de texto con el contenido del reporte
    var doc = DocumentApp.create(nombreArchivoBase);
    var body = doc.getBody();


    body.appendParagraph('Reporte de tienda - Multiservicios Sr. Puerto Malaga')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);


    body.appendParagraph(
      'Fecha de generación: ' +
      Utilities.formatDate(hoy, zona, 'dd/MM/yyyy HH:mm')
    );
    body.appendParagraph('');


    // Rango de filtros, si se usó
    var desdeStr = filtros.desde ? Utilities.formatDate(new Date(filtros.desde), zona, 'dd/MM/yyyy') : '';
    var hastaStr = filtros.hasta ? Utilities.formatDate(new Date(filtros.hasta), zona, 'dd/MM/yyyy') : '';
    if (desdeStr || hastaStr) {
      body.appendParagraph('Rango filtrado: ' + (desdeStr || '—') + ' - ' + (hastaStr || '—'));
      body.appendParagraph('');
    }


    // ===== 1. KPIs =====
    body.appendParagraph('1. Resumen de indicadores')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);


    var kpisTable = body.appendTable([
      ['Indicador',           'Valor'],
      ['Ventas del día',      'S/ ' + Number(datos.kpis.ventasDia || 0).toFixed(2)],
      ['Ingresos del mes',    'S/ ' + Number(datos.kpis.ingresosMes || 0).toFixed(2)],
      ['Productos en stock',  String(datos.kpis.productosEnStock || 0)],
      ['Por agotarse',        String(datos.kpis.porAgotarse || 0)]
    ]);
    kpisTable.setBorderWidth(0.5);


    body.appendParagraph('');


    // ===== 2. Productos más vendidos =====
    body.appendParagraph('2. Productos más vendidos')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);


    var masVendidos = datos.tablas.masVendidos || [];
    if (masVendidos.length > 0) {
      var tv = [['Producto', 'Categoría', 'Unidades', 'Ganancia (S/)']];
      masVendidos.forEach(function(r) {
        tv.push([
          r.nombre || '',
          r.categoria || '',
          String(r.unidades || 0),
          Number(r.ganancia || 0).toFixed(2)
        ]);
      });
      var t2 = body.appendTable(tv);
      t2.setBorderWidth(0.5);
    } else {
      body.appendParagraph('No hay datos de ventas en el periodo seleccionado.');
    }


    body.appendParagraph('');


    // ===== 3. Productos con bajo stock =====
    body.appendParagraph('3. Productos con bajo stock')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);


    var bajoStock = datos.tablas.bajoStock || [];
    if (bajoStock.length > 0) {
      var tb = [['Producto', 'Categoría', 'Stock', 'Mínimo']];
      bajoStock.forEach(function(r) {
        tb.push([
          r.nombre || '',
          r.categoria || '',
          String(r.stock || 0),
          String(r.minimo || 0)
        ]);
      });
      var t3 = body.appendTable(tb);
      t3.setBorderWidth(0.5);
    } else {
      body.appendParagraph('No hay productos con stock bajo en el periodo seleccionado.');
    }


    doc.saveAndClose();


    // 2) Convertir a PDF
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(nombreArchivoBase + '.pdf');


    // 3) Registrar en la hoja REPORTE
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var repSheet = ss.getSheetByName('REPORTE');
    if (!repSheet) {
      throw new Error('No se encontró la hoja REPORTE');
    }


    var headers = repSheet.getRange(1, 1, 1, repSheet.getLastColumn()).getValues()[0];
    var nuevaFila = new Array(headers.length).fill('');
    var lastRow = repSheet.getLastRow();
    var idReporte = 'REP' + String(lastRow).padStart(3, '0');
    var usuario = Session.getActiveUser().getEmail() || '';


    function setCampo(nombreColumna, valor) {
      var idx = headers.indexOf(nombreColumna);
      if (idx >= 0) nuevaFila[idx] = valor;
    }


    setCampo('ID_Reporte',       idReporte);
    setCampo('Tipo_Reporte',     'Dashboard');
    setCampo('Fecha_Generacion', hoy);
    setCampo('Fecha_Desde',      filtros.desde || '');
    setCampo('Fecha_Hasta',      filtros.hasta || '');
    setCampo('Generado_Por',     usuario);
    setCampo('Nivel_Detalle',    'Resumido');
    setCampo('Email_Destino',    '');
    setCampo('Drive_File_Id',    pdfFile.getId());
    setCampo('Nombre_Archivo',   pdfFile.getName());
    setCampo('Origen',           'WebApp');


    repSheet.appendRow(nuevaFila);
    
    // 4. Borrar el Google Doc temporal
    DriveApp.getFileById(doc.getId()).setTrashed(true);


    return {
      success: true,
      idReporte: idReporte,
      fileId: pdfFile.getId(),
      nombreArchivo: pdfFile.getName(),
      url: pdfFile.getUrl()
    };


  } catch (error) {
    console.error('Error generarReporteDashboard:', error);
    return { success: false, error: error.toString() };
  }
}






// ============== REPORTES (PDF + HISTORIAL) ==============


// Generar PDF con datos principales del dashboard y registrar en HISTORIAL_REPORTE
function generarReportePDF(opciones) {
  opciones = opciones || {};
  var filtrosDashboard = {
    desde: opciones.desde || '',
    hasta: opciones.hasta || '',
    categoria: '',
    proveedor: '',
    cliente: '',
    frecuencia: opciones.frecuencia || 'mensual'
  };


  var dash = getDashboardData(filtrosDashboard);


  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var file = DriveApp.getFileById(SPREADSHEET_ID);
  var parentFolder = file.getParents().hasNext()
    ? file.getParents().next()
    : DriveApp.getRootFolder();


  var fechaAhora = new Date();
  var tz = Session.getScriptTimeZone();
  var fechaStr = Utilities.formatDate(fechaAhora, tz, 'dd/MM/yyyy HH:mm');


  var doc = DocumentApp.create('Reporte_Tienda_Temporal_' + fechaAhora.getTime());
  var body = doc.getBody();


  body.appendParagraph('REPORTE DE TIENDA').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Multiservicios Sr. Puerto Malaga').setHeading(
    DocumentApp.ParagraphHeading.HEADING2
  );
  body.appendParagraph('Generado: ' + fechaStr);
  body.appendParagraph('');


  // Rango de fechas del reporte
  if (opciones.desde || opciones.hasta) {
    body.appendParagraph(
      'Rango de análisis: ' +
        (opciones.desde || '—') +
        '  a  ' +
        (opciones.hasta || '—')
    );
    body.appendParagraph('');
  }


  // KPIs principales
  body.appendParagraph('Resumen de indicadores').setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );


  var tablaKpi = body.appendTable();
  tablaKpi.appendTableRow()
    .appendTableCell('Indicador')
    .appendTableCell('Valor');


  tablaKpi.appendTableRow()
    .appendTableCell('Ventas del día')
    .appendTableCell('S/ ' + (dash.kpis.ventasDia || 0).toFixed(2));
  tablaKpi.appendTableRow()
    .appendTableCell('Ingresos del mes')
    .appendTableCell('S/ ' + (dash.kpis.ingresosMes || 0).toFixed(2));
  tablaKpi.appendTableRow()
    .appendTableCell('Productos en stock')
    .appendTableCell(String(dash.kpis.productosEnStock || 0));
  tablaKpi.appendTableRow()
    .appendTableCell('Productos por agotarse')
    .appendTableCell(String(dash.kpis.porAgotarse || 0));
  tablaKpi.appendTableRow()
    .appendTableCell('Clientes registrados (rol Cliente)')
    .appendTableCell(String(dash.kpis.clientes || 0));
  tablaKpi.appendTableRow()
    .appendTableCell('Pedidos pendientes')
    .appendTableCell(String(dash.kpis.pedidosPendientes || 0));
  tablaKpi.appendTableRow()
    .appendTableCell('Pedidos entregados/completados')
    .appendTableCell(String(dash.kpis.pedidosEntregados || 0));


  body.appendParagraph('');


  // Top productos
  body.appendParagraph('Top productos más vendidos (por monto generado)').setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );


  var topProductos = dash.tablas.masVendidos || [];
  if (topProductos.length > 0) {
    var tablaTop = body.appendTable();
    tablaTop.appendTableRow()
      .appendTableCell('Producto')
      .appendTableCell('Categoría')
      .appendTableCell('Unidades')
      .appendTableCell('Monto (S/)');


    topProductos.slice(0, 10).forEach(function (p) {
      tablaTop.appendTableRow()
        .appendTableCell(p.nombre || '')
        .appendTableCell(p.categoria || '')
        .appendTableCell(String(p.unidades || 0))
        .appendTableCell((p.ganancia || 0).toFixed(2));
    });
  } else {
    body.appendParagraph('No hay datos de ventas para el periodo analizado.');
  }


  body.appendParagraph('');


  // Alertas de stock
  body.appendParagraph('Productos con stock bajo').setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var bajoStock = dash.tablas.bajoStock || [];
  if (bajoStock.length > 0) {
    var tablaStock = body.appendTable();
    tablaStock.appendTableRow()
      .appendTableCell('Producto')
      .appendTableCell('Categoría')
      .appendTableCell('Stock')
      .appendTableCell('Mínimo');


    bajoStock.forEach(function (p) {
      tablaStock.appendTableRow()
        .appendTableCell(p.nombre || '')
        .appendTableCell(p.categoria || '')
        .appendTableCell(String(p.stock || 0))
        .appendTableCell(String(p.minimo || 0));
    });
  } else {
    body.appendParagraph('No hay productos con stock por debajo del mínimo.');
  }


  body.appendParagraph('');
  body.appendParagraph('Fin del reporte.').setItalic(true);


  doc.saveAndClose();


  var pdfBlob = doc.getAs(MimeType.PDF);
  var nombreArchivo =
    'Reporte_Tienda_' +
    Utilities.formatDate(fechaAhora, tz, 'yyyyMMdd_HHmm') +
    '.pdf';
  var pdfFile = parentFolder.createFile(pdfBlob).setName(nombreArchivo);


  // Registrar en HISTORIAL_REPORTE
  var histSheet = ss.getSheetByName('REPORTE');
  var lastRow = histSheet.getLastRow();
  var idReporte = 'REP' + Utilities.formatString('%04d', lastRow);


  histSheet.appendRow([
    idReporte,
    opciones.tipoReporte || 'Dashboard',
    new Date(),
    opciones.desde || '',
    opciones.hasta || '',
    opciones.generadoPor || '',
    opciones.nivelDetalle || 'Resumido',
    opciones.emailDestino || '',
    pdfFile.getId(),
    pdfFile.getName(),
    opciones.origen || 'Manual'
  ]);
  
  // Borrar el Google Doc temporal
  DriveApp.getFileById(doc.getId()).setTrashed(true);


  return {
    success: true,
    idReporte: idReporte,
    fileId: pdfFile.getId(),
    fileUrl: pdfFile.getUrl(),
    nombreArchivo: pdfFile.getName()
  };
}


// Listar historial de reportes (con filtros sencillos)
function getHistorialReportes(filtros) {
  filtros = filtros || {};
  var datos = getSheetData('REPORTE');


  var desde = filtros.desde ? parseFecha(filtros.desde) : null;
  var hasta = filtros.hasta ? parseFecha(filtros.hasta) : null;
  if (hasta) {
    hasta.setHours(23, 59, 59, 999);
  }


  return datos.filter(function (r) {
    var ok = true;


    if (filtros.tipoReporte) {
      var tr = (r.Tipo_Reporte || '').toLowerCase();
      if (tr.indexOf(filtros.tipoReporte.toLowerCase()) === -1) ok = false;
    }


    if (filtros.origen) {
      var org = (r.Origen || '').toLowerCase();
      if (org.indexOf(filtros.origen.toLowerCase()) === -1) ok = false;
    }


    if (filtros.emailDestino) {
      var ed = (r.Email_Destino || '').toLowerCase();
      if (ed.indexOf(filtros.emailDestino.toLowerCase()) === -1) ok = false;
    }


    if (desde || hasta) {
      var fg = parseFecha(r.Fecha_Generacion);
      if (!fg) return false;
      if (desde && fg < desde) ok = false;
      if (hasta && fg > hasta) ok = false;
    }


    return ok;
  });
}


// Configurar envío diario de reporte por correo
function configurarEnvioDiario(emailDestino, hora) {
  if (!emailDestino || !hora) {
    throw new Error('Email y hora son obligatorios');
  }


  var props = PropertiesService.getScriptProperties();
  props.setProperty('REPORTE_EMAIL_DESTINO', emailDestino);
  props.setProperty('REPORTE_HORA', hora);


  // Eliminar triggers previos para evitar duplicados
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarReporteProgramado') {
      ScriptApp.deleteTrigger(t);
    }
  });


  var partes = hora.split(':');
  var h = parseInt(partes[0], 10);
  if (isNaN(h) || h < 0 || h > 23) {
    h = 8; // hora por defecto
  }


  ScriptApp.newTrigger('enviarReporteProgramado')
    .timeBased()
    .everyDays(1)
    .atHour(h)
    .create();


  return { success: true };
}


// Trigger diario: genera reporte y lo envía por email
function enviarReporteProgramado() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('REPORTE_EMAIL_DESTINO');
  if (!email) {
    console.log('No hay email configurado para reportes programados');
    return;
  }


  var tz = Session.getScriptTimeZone();
  var hoy = new Date();
  var fechaStr = Utilities.formatDate(hoy, tz, 'dd/MM/yyyy');


  var resultado = generarReportePDF({
    tipoReporte: 'Diario',
    origen: 'Programado',
    nivelDetalle: 'Resumido',
    emailDestino: email
  });


  var file = DriveApp.getFileById(resultado.fileId);


  MailApp.sendEmail({
    to: email,
    subject: 'Reporte diario - Multiservicios Sr. Puerto Malaga (' + fechaStr + ')',
    htmlBody:
      'Hola,<br><br>' +
      'Adjunto el reporte diario de la tienda correspondiente al <b>' +
      fechaStr +
      '</b>.<br><br>' +
      'Archivo: <b>' +
      file.getName() +
      '</b><br><br>Saludos.', // <_ CORREGIDO
    attachments: [file.getAs(MimeType.PDF)]
  });
}


// Diagnóstico simple del sistema
function diagnosticarSistema() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var nombres = sheets.map(function (s) {
      return s.getName();
    });


    var datos = {};
    nombres.forEach(function (n) {
      datos[n] = getSheetData(n).length;
    });


    return {
      status: 'OK',
      sheets: nombres,
      datos: datos,
      spreadsheetId: SPREADSHEET_ID
    };
  } catch (e) {
    return {
      status: 'ERROR',
      error: e.toString(),
      spreadsheetId: SPREADSHEET_ID
    };
  }
}

//AGRAGADO----------------------------------------------------------------------------------------------------
// Esta función estaba duplicada. Se eliminó la versión simple que solo hacía getSheetData('REPORTE').
// Se mantiene la versión compleja (ver arriba) que permite filtrar.


//AGRAGADO----------------------------------------------------------------------------------------------------
function getConfigEnvioReporte() {
  try { // <_ CORREGIDO
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('REPORTE_CONFIG');
    if (!sheet) {
      return {};
    }


    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {};
    }


    var headers = data[0];
    var row = data[1];
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
    });


    return {
      emailDestino: obj.Email_Destino || '',
      horaEnvio: obj.Hora_Envio || '',
      activo: String(obj.Activo || '') === 'TRUE' || String(obj.Activo).toLowerCase() === 'si'
    }; // <_ CORREGIDO


  } catch (e) {
    console.error('Error getConfigEnvioReporte:', e);
    return {};
  }
}




//AGRAGADO----------------------------------------------------------------------------------------------------
function guardarConfigEnvioReporte(config) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('REPORTE_CONFIG');


    if (!sheet) {
      sheet = ss.insertSheet('REPORTE_CONFIG');
      sheet.appendRow(['ID_Config', 'Email_Destino', 'Hora_Envio', 'Activo', 'Ultimo_Envio']);
    } // <_ CORREGIDO


    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rowIndex = data.length >= 2 ? 2 : sheet.getLastRow() + 1;


    var idConfig = 'CFG001';


    var fila = new Array(headers.length).fill('');
    headers.forEach(function(h, i) {
      if (h === 'ID_Config') {
          fila[i] = idConfig;
      } else if (h === 'Email_Destino') {
          fila[i] = config.emailDestino;
      } else if (h === 'Hora_Envio') {
          fila[i] = config.horaEnvio;
      } else if (h === 'Activo') {
          fila[i] = config.activo ? 'SI' : 'NO';
      } else if (h === 'Ultimo_Envio' && data.length >= 2) {
          // Ultimo_Envio se mantiene
          fila[i] = data[1][i]; // conservar valor anterior
      }
    });


    if (data.length >= 2) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([fila]);
    } else {
      sheet.appendRow(fila);
    }


    // Gestionar el trigger diario
    actualizarTriggerEnvioDiario(config.horaEnvio, config.activo);


    return { success: true };
  } catch (e) {
    console.error('Error guardarConfigEnvioReporte:', e);
    return { success: false, error: e.toString() };
  }
}




//AGRAGADO----------------------------------------------------------------------------------------------------
function actualizarTriggerEnvioDiario(horaStr, activo) {
  // <_ CORREGIDO
  // Eliminar triggers anteriores de esta función
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarReporteDiario') {
      ScriptApp.deleteTrigger(t);
    }
  });


  if (!activo) return;
  if (!horaStr) return;


  var partes = horaStr.split(':');
  var h = parseInt(partes[0], 10);
  var m = partes.length > 1 ? parseInt(partes[1], 10) : 0;


  if (isNaN(h) || h < 0 || h > 23) h = 9;
  if (isNaN(m) || m < 0 || m > 59) m = 0;


  ScriptApp.newTrigger('enviarReporteDiario')
    .timeBased()
    .atHour(h)
    .nearMinute(m)
    .everyDays(1)
    .create();
}



/**
 * ESTA ES LA FUNCIÓN CORRECTA PARA ENVIAR EL REPORTE DIARIO.
 * Reemplaza la que tienes.
 */
function enviarReporteDiario() {
  try {
    Logger.log("Iniciando envío de reporte diario...");
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('REPORTE_CONFIG');
    if (!sheet) {
        Logger.log("Error: No se encontró la hoja REPORTE_CONFIG.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('Error: REPORTE_CONFIG está vacío. No se puede enviar el reporte.');
      return; // Sale si no hay configuración
    }

    var headers = data[0];
    var row = data[1];
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });

    // DESPUÉS:
var valorActivo = obj.Activo;
var activo = (valorActivo === true || String(valorActivo).toUpperCase() === 'VERDADERO' || String(valorActivo).toUpperCase() === 'TRUE' || String(valorActivo).toLowerCase() === 'si');
    if (!activo) {
      Logger.log('Envío diario está inactivo. No se envía reporte.');
      return; // Sale si no está activo
    }

    var email = obj.Email_Destino;
    if (!email) {
      Logger.log('Error: No hay email configurado en REPORTE_CONFIG.');
      return; // Sale si no hay email
    }

    Logger.log('Configuración cargada. Email: ' + email + '. Generando reporte...');

    // Generar reporte (esta función ya guarda en la hoja REPORTE)
    var res = generarReporteDashboard({}); // Usamos filtros vacíos para el reporte diario
    if (!res || !res.success) {
      Logger.log('Error al generar el PDF del dashboard.');
      return;
    }
    
    Logger.log('Reporte PDF generado. ID de archivo: ' + res.fileId);

    var file = DriveApp.getFileById(res.fileId);
    var asunto = 'Reporte diario - Multiservicios Sr. Puerto Malaga';
    var cuerpo =
      'Adjunto encontrarás el reporte diario de la tienda.\n\n' +
      'Archivo: ' + res.nombreArchivo + '\n\n' +
      'Este correo fue generado automáticamente.';

    MailApp.sendEmail({
      to: email,
      subject: asunto,
      body: cuerpo,
      attachments: [file.getAs('application/pdf')]
    });

    Logger.log('Correo con reporte enviado exitosamente a: ' + email);

    // Actualizar Ultimo_Envio
    var idxUlt = headers.indexOf('Ultimo_Envio');
    if (idxUlt >= 0) {
      sheet.getRange(2, idxUlt + 1).setValue(new Date());
    }

  } catch (e) {
    Logger.log('Error fatal en enviarReporteDiario: ' + e.toString());
    console.error('Error enviarReporteDiario:', e);
  }
}



//NUEVA FUNCION CIERRE DE CAJA-------------------------------------------------------------------------------------------------

function generarCierreCajaPDF(idUsuario) {
  try {
    var zona = 'America/Lima'; // Usar la misma zona
    var hoy = new Date();
    
    // Filtros: SÓLO HOY
    var inicioDia = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 0, 0, 0);
    var finDia = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);

    var filtros = {
      desde: inicioDia,
      hasta: finDia
    };
    
    // 1. Obtener los datos (¡Ahora incluirá resumenPagos!)
    var datos = getDashboardData(filtros); // <_ CORREGIDO
    
    var nombreArchivoBase = 'Cierre_Caja_' + Utilities.formatDate(hoy, zona, 'yyyyMMdd_HHmm');
    
    // 2. Crear el Documento
    var doc = DocumentApp.create(nombreArchivoBase);
    var body = doc.getBody();
    
    body.appendParagraph('Cierre de Caja - Multiservicios Sr. Puerto Malaga')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Fecha de Cierre: ' + Utilities.formatDate(hoy, zona, 'dd/MM/yyyy HH:mm'));
    body.appendParagraph('Generado por: ' + (idUsuario || Session.getActiveUser().getEmail()));
    body.appendParagraph('');

    // 3. Resumen de Ventas (KPIs)
    body.appendParagraph('1. Resumen de Ventas del Día')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    var totalVentasDia = (datos.kpis.ventasDia || 0).toFixed(2);
    body.appendTable([
      ['Ventas Totales del Día', 'S/ ' + totalVentasDia],
      ['Pedidos Completados', String(datos.kpis.pedidosEntregados || 0)]
    ]).setBorderWidth(0.5);
    body.appendParagraph('');

    // 4. Desglose por Método de Pago (¡NUEVO!)
    body.appendParagraph('2. Desglose por Método de Pago')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    var resumenPagos = datos.tablas.resumenPagos || [];
    if (resumenPagos.length > 0) {
      var tPagos = [['Método de Pago', 'Total Recaudado (S/)']];
      var totalRecaudado = 0;
      resumenPagos.forEach(function(p) {
        tPagos.push([
          p.metodo || 'Desconocido',
          Number(p.total || 0).toFixed(2)
        ]);
        totalRecaudado += (p.total || 0);
      }); // <_ CORREGIDO
      // Fila de total
      tPagos.push(['TOTAL', totalRecaudado.toFixed(2)]);
      body.appendTable(tPagos).setBorderWidth(0.5); // <_ CORREGIDO
    } else {
      body.appendParagraph('No se registraron pagos en el día.');
    }
    body.appendParagraph('');

    // 5. Productos más vendidos
    body.appendParagraph('3. Productos Más Vendidos del Día')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var masVendidos = datos.tablas.masVendidos || [];
    if (masVendidos.length > 0) {
      var tv = [['Producto', 'Unidades', 'Ganancia (S/)']];
      masVendidos.forEach(function(r) {
        tv.push([
          r.nombre || '',
          String(r.unidades || 0),
          Number(r.ganancia || 0).toFixed(2)
        ]);
      });
      body.appendTable(tv).setBorderWidth(0.5);
    } else {
      body.appendParagraph('No hubo ventas de productos en el día.');
    }
    body.appendParagraph('');
    
    doc.saveAndClose();
    
    // 6. Convertir a PDF
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(nombreArchivoBase + '.pdf');

    // 7. Registrar en REPORTE
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var repSheet = ss.getSheetByName('REPORTE'); // <_ CORREGIDO
    var headers = repSheet.getRange(1, 1, 1, repSheet.getLastColumn()).getValues()[0];
    var nuevaFila = new Array(headers.length).fill('');
    var lastRow = repSheet.getLastRow();
    var idReporte = 'REP' + String(lastRow).padStart(3, '0');

    function setCampo(nombreColumna, valor) { // <_ CORREGIDO
      var idx = headers.indexOf(nombreColumna);
      if (idx >= 0) nuevaFila[idx] = valor;
    }

    setCampo('ID_Reporte',       idReporte);
    setCampo('Tipo_Reporte',    'Cierre de Caja');
    setCampo('Fecha_Generacion', hoy);
    setCampo('Fecha_Desde',      inicioDia);
    setCampo('Fecha_Hasta',      finDia); // <_ CORREGIDO
    setCampo('Generado_Por',     idUsuario || Session.getActiveUser().getEmail());
    setCampo('Nivel_Detalle',    'Resumido');
    setCampo('Drive_File_Id',    pdfFile.getId());
    setCampo('Nombre_Archivo',   pdfFile.getName());
    setCampo('Origen',           'WebApp (Cierre)');

    repSheet.appendRow(nuevaFila);
    
    // 8. Borrar el Google Doc temporal
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return {
      success: true,
      idReporte: idReporte,
      fileId: pdfFile.getId(),
      nombreArchivo: pdfFile.getName(),
      url: pdfFile.getUrl()
    };

  } catch (error) {
    console.error('Error generarCierreCajaPDF:', error);
    return { success: false, error: error.toString() }; // <_ CORREGIDO
  }
}


/**
 * ESTA FUNCIÓN ESTÁ DISEÑADA PARA IMPRIMIR MENSAJES EN EL LOG.
 */
function PRUEBA_DE_LOG_Y_CUOTA() {

  Logger.log("--- INICIANDO PRUEBA ---");

  try {
    var cuota = MailApp.getRemainingDailyQuota();
    Logger.log("Correos restantes para hoy: " + cuota);

    if (cuota > 0) {
      Logger.log("Tienes cuota. Intentando enviar email...");
      MailApp.sendEmail(
        "gpar781+testfinal@gmail.com",
        "Prueba Final (Log)",
        "Email de prueba."
      );
      Logger.log("Comando de envío de email ejecutado.");
    } else {
      Logger.log("ERROR: No tienes cuota (Correos restantes: 0). Debes esperar 24 horas.");
    }

  } catch (e) {
    Logger.log("ERROR GRAVE AL INTENTAR ENVIAR: " + e.message);
  }

  Logger.log("--- PRUEBA TERMINADA ---");
}
