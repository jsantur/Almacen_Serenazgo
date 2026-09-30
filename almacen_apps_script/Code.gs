const APP = {
  NAME: 'Almacén de Prendas y Equipamiento',
  DB_PROP: 'ALMACEN_DB_ID',
  SHEETS: {
    PRODUCTS: 'PRODUCTOS',
    MOVEMENTS: 'MOVIMIENTOS',
    CONFIG: 'CONFIG'
  },
  PRODUCT_HEADERS: ['ID','CODIGO','CATEGORIA','NATURALEZA','TIPO','MODELO','MARCA','COLOR','TALLA','PRESENTACION','GENERO','UBICACION','UNIDAD','CONDICION','PUNTO_REORDEN','STOCK_MAXIMO','DESCRIPCION','ACTIVO','CREADO_EN'],
  MOVEMENT_HEADERS: ['ID','MOVEMENT_ID','FECHA','CODIGO','TIPO_MOVIMIENTO','CANTIDAD','CONDICION','MOTIVO','ORIGEN_DESTINO','BENEFICIARIO','DNI_BENEFICIARIO','AREA_DESTINO','CARGO_BENEFICIARIO','DOCUMENTO_ENTREGA','USUARIO','OBSERVACION','CREADO_EN'],
  CONFIG_HEADERS: ['CLAVE','VALOR']
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP.NAME);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Ejecutar UNA VEZ desde el editor de Apps Script.
 * Crea la base de datos Google Sheets y deja datos de demostración.
 */
function setupSistema() {
  const ss = _getDb_();
  _ensureSheet_(ss, APP.SHEETS.PRODUCTS, APP.PRODUCT_HEADERS);
  _ensureSheet_(ss, APP.SHEETS.MOVEMENTS, APP.MOVEMENT_HEADERS);
  _ensureSheet_(ss, APP.SHEETS.CONFIG, APP.CONFIG_HEADERS);

  _writeDefaultConfig_(ss);
  _seedDemoData_(ss, true);
  SpreadsheetApp.flush();

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    message: 'Sistema re-inicializado correctamente con los datos ficticios de demostración (demo.html).'
  };
}

function getSystemInfo() {
  try {
    const ss = _getDb_();
    return {
      configured: true,
      name: APP.NAME,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      timezone: Session.getScriptTimeZone()
    };
  } catch (e) {
    return { configured: false, name: APP.NAME };
  }
}

function getAppData() {
  const ss = _getDb_();
  const products = _readSheet_(ss, APP.SHEETS.PRODUCTS);
  const movements = _readSheet_(ss, APP.SHEETS.MOVEMENTS);
  const inventory = _buildInventory_(products, movements);
  const activeProducts = products.filter(p => String(p.ACTIVO).toUpperCase() !== 'NO');
  const now = new Date();
  const cutoff = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const recentMoves = movements.filter(m => new Date(m.CREADO_EN || m.FECHA) >= cutoff);
  const low = inventory.filter(i => i.stock <= i.stockMin);
  const out = inventory.filter(i => i.stock <= 0);

  const natureCounts = activeProducts.reduce((a,p)=>{a[p.NATURALEZA]=(a[p.NATURALEZA]||0)+1;return a;},{});

  return {
    info: getSystemInfo(),
    products: products.map(_publicProduct_),
    inventory: inventory.sort((a,b) => a.tipo.localeCompare(b.tipo, 'es') || a.codigo.localeCompare(b.codigo, 'es')),
    movements: movements.sort(_movementSort_).slice(0, 150).map(_publicMovement_),
    metrics: {
      skus: activeProducts.length,
      totalStock: inventory.reduce((s, i) => s + i.stock, 0),
      lowStock: low.length,
      outOfStock: out.length,
      movements30d: recentMoves.length,
      inventariables: natureCounts['INVENTARIABLE'] || 0,
      consumibles: natureCounts['CONSUMIBLE'] || 0
    },
    lists: {
      tipos: _unique_(activeProducts.map(p => p.TIPO)),
      categorias: _unique_(activeProducts.map(p => p.CATEGORIA)),
      naturalezas: _unique_(activeProducts.map(p => p.NATURALEZA)),
      generos: _unique_(activeProducts.map(p => p.GENERO)),
      unidades: _unique_(activeProducts.map(p => p.UNIDAD)),
      condiciones: _unique_(activeProducts.map(p => p.CONDICION)),
      tallas: _unique_(activeProducts.map(p => p.TALLA)),
      presentaciones: _unique_(activeProducts.map(p => p.PRESENTACION)),
      motivos: _getConfigList_(ss, 'MOTIVOS')
    }
  };
}

function createProduct(form) {
  _validateProduct_(form);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _getDb_();
    const sh = ss.getSheetByName(APP.SHEETS.PRODUCTS);
    const rows = _readSheet_(ss, APP.SHEETS.PRODUCTS);
    const code = String(form.codigo).trim().toUpperCase();
    if (rows.some(r => String(r.CODIGO).toUpperCase() === code)) {
      throw new Error('Ya existe un producto con el código ' + code + '.');
    }
    const data = {
      'ID': _id_('PRD'),
      'CODIGO': code,
      'CATEGORIA': String(form.categoria || '').trim(),
      'NATURALEZA': String(form.naturaleza || 'INVENTARIABLE').trim(),
      'TIPO': String(form.tipo || '').trim(),
      'MODELO': String(form.modelo || '').trim(),
      'MARCA': String(form.marca || '').trim(),
      'COLOR': String(form.color || '').trim(),
      'TALLA': String(form.talla || 'NO APLICA').trim(),
      'PRESENTACION': String(form.presentacion || 'NO APLICA').trim(),
      'GENERO': String(form.genero || 'NO APLICA').trim(),
      'UBICACION': String(form.ubicacion || '').trim(),
      'UNIDAD': String(form.unidad || 'UNIDAD').trim(),
      'CONDICION': String(form.condicion || 'NUEVO').trim(),
      'PUNTO_REORDEN': Number(form.puntoReorden || 0),
      'STOCK_MAXIMO': Number(form.stockMaximo || 0),
      'DESCRIPCION': String(form.descripcion || '').trim(),
      'ACTIVO': 'SI',
      'CREADO_EN': new Date()
    };
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const row = new Array(headers.length).fill('');
    headers.forEach((h, i) => {
      const key = String(h).trim().toUpperCase();
      if (data[key] !== undefined) {
        row[i] = data[key];
      }
    });
    sh.appendRow(row);

    const initial = Number(form.stockInicial || 0);
    if (initial > 0) {
      _appendMovement_(ss, {
        fecha: form.fecha || _today_(),
        codigo: code,
        tipoMovimiento: 'SALDO_INICIAL',
        cantidad: initial,
        condicion: form.condicion || 'NUEVO',
        motivo: 'Carga inicial',
        origenDestino: 'Almacén',
        observacion: 'Stock inicial al crear el SKU'
      });
    }
    SpreadsheetApp.flush();
    return { ok: true, message: 'Producto registrado correctamente.' };
  } finally {
    lock.releaseLock();
  }
}

function deactivateProduct(codigo) {
  return toggleProductStatus(codigo);
}

function toggleProductStatus(codigo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _getDb_();
    const sh = ss.getSheetByName(APP.SHEETS.PRODUCTS);
    const colMap = _resolveColumns_(sh);
    const values = sh.getDataRange().getValues();
    const code = String(codigo || '').trim().toUpperCase();
    
    const codeIndex = colMap['CODIGO'] - 1;
    const activoIndex = colMap['ACTIVO'] - 1;
    
    if (codeIndex < 0 || activoIndex < 0) throw new Error('Schema de PRODUCTOS inválido.');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][codeIndex]).toUpperCase() === code) {
        const currentStatus = String(values[r][activoIndex]).toUpperCase();
        const newStatus = currentStatus === 'NO' ? 'SI' : 'NO';
        _updateRowBySchema_(sh, r + 1, { 'ACTIVO': newStatus }, colMap);
        SpreadsheetApp.flush();
        return { ok: true, message: newStatus === 'SI' ? 'Producto activado correctamente.' : 'Producto desactivado correctamente.' };
      }
    }
    throw new Error('Producto no encontrado.');
  } finally {
    lock.releaseLock();
  }
}

function editProduct(form) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _getDb_();
    const sh = ss.getSheetByName(APP.SHEETS.PRODUCTS);
    const colMap = _resolveColumns_(sh);
    const values = sh.getDataRange().getValues();
    const originalCode = String(form.originalCodigo || '').trim().toUpperCase();
    if (!originalCode) throw new Error('Código original no proporcionado.');

    const newSku = String(form.codigo || '').trim().toUpperCase();
    const codeIndex = colMap['CODIGO'] - 1;

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][codeIndex]).toUpperCase() === originalCode) {
        
        _updateRowBySchema_(sh, r + 1, {
          'CODIGO': newSku,
          'CATEGORIA': String(form.categoria || '').trim(),
          'NATURALEZA': String(form.naturaleza || 'INVENTARIABLE').trim(),
          'TIPO': String(form.tipo || '').trim(),
          'MODELO': String(form.modelo || '').trim(),
          'MARCA': String(form.marca || '').trim(),
          'COLOR': String(form.color || '').trim(),
          'TALLA': String(form.talla || 'NO APLICA').trim(),
          'PRESENTACION': String(form.presentacion || 'NO APLICA').trim(),
          'GENERO': String(form.genero || 'NO APLICA').trim(),
          'UBICACION': String(form.ubicacion || '').trim(),
          'UNIDAD': String(form.unidad || 'UNIDAD').trim(),
          'CONDICION': String(form.condicion || 'NUEVO').trim(),
          'PUNTO_REORDEN': Number(form.puntoReorden || 0),
          'STOCK_MAXIMO': Number(form.stockMaximo || 0),
          'DESCRIPCION': String(form.descripcion || '').trim()
        }, colMap);

        // -- MANEJAR EL STOCK NUEVO AQUÍ --
        const newStockStr = form.stockNuevo;
        let adjustmentMessage = '';
        
        if (newStockStr !== undefined && newStockStr !== '') {
           const newStock = Number(newStockStr);
           if (Number.isFinite(newStock) && newStock >= 0) {
              const products = _readSheet_(ss, APP.SHEETS.PRODUCTS);
              const movements = _readSheet_(ss, APP.SHEETS.MOVEMENTS);
              const inventory = _buildInventory_(products, movements);
              const currentInv = inventory.find(i => i.codigo === originalCode);
              const oldStock = currentInv ? currentInv.stock : 0;
              
              const delta = newStock - oldStock;
              if (delta !== 0) {
                 const tipo = delta > 0 ? 'AJUSTE_POSITIVO' : 'AJUSTE_NEGATIVO';
                 _appendMovement_(ss, {
                    movementId: _id_('AJUSTE'),
                    fecha: _today_(),
                    codigo: newSku,
                    tipoMovimiento: tipo,
                    cantidad: Math.abs(delta),
                    condicion: String(form.condicion || 'NUEVO').trim(),
                    motivo: 'Corrección de stock desde edición del producto',
                    origenDestino: 'Almacén',
                    observacion: 'Stock anterior: ' + oldStock + ' | Stock corregido: ' + newStock
                 });
                 adjustmentMessage = ' y stock corregido mediante ajuste de Kardex';
              }
           } else {
              throw new Error('El nuevo stock debe ser un número mayor o igual a 0.');
           }
        }

        // If the SKU changed, cascade update the MOVIMIENTOS sheet
        if (newSku !== originalCode) {
          const mvSh = ss.getSheetByName(APP.SHEETS.MOVEMENTS);
          const mvColMap = _resolveColumns_(mvSh);
          const mvValues = mvSh.getDataRange().getValues();
          const mvCodeIdx = mvColMap['CODIGO'] - 1;
          for (let m = 1; m < mvValues.length; m++) {
            if (String(mvValues[m][mvCodeIdx]).toUpperCase() === originalCode) {
              _updateRowBySchema_(mvSh, m + 1, { 'CODIGO': newSku }, mvColMap);
            }
          }
        }

        SpreadsheetApp.flush();
        return { ok: true, message: 'Producto actualizado correctamente' + adjustmentMessage + '.' + (newSku !== originalCode ? ' SKU actualizado a ' + newSku + '.' : '') };
      }
    }
    throw new Error('Producto no encontrado: ' + originalCode);
  } finally {
    lock.releaseLock();
  }
}

function recordMovement(form) {
  _validateMovement_(form);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = _getDb_();
    const products = _readSheet_(ss, APP.SHEETS.PRODUCTS);
    const product = products.find(p => String(p.CODIGO).toUpperCase() === String(form.codigo).toUpperCase());
    if (!product || String(product.ACTIVO).toUpperCase() === 'NO') {
      throw new Error('El SKU no existe o está inactivo.');
    }

    const movements = _readSheet_(ss, APP.SHEETS.MOVEMENTS);
    const inventory = _buildInventory_(products, movements);
    const current = inventory.find(i => i.codigo === String(form.codigo).toUpperCase());
    const qty = Number(form.cantidad);
    const type = String(form.tipoMovimiento).toUpperCase();

    if (['SALIDA','AJUSTE_NEGATIVO'].includes(type) && current && current.stock < qty) {
      throw new Error('Stock insuficiente. Disponible: ' + current.stock + ' ' + (product.UNIDAD || 'UNIDAD'));
    }

    _appendMovement_(ss, {
      fecha: form.fecha || _today_(),
      codigo: String(form.codigo).trim().toUpperCase(),
      tipoMovimiento: type,
      cantidad: qty,
      condicion: String(form.condicion || '').trim(),
      motivo: String(form.motivo || '').trim(),
      origenDestino: String(form.origenDestino || '').trim(),
      beneficiario: String(form.beneficiario || '').trim(),
      dniBeneficiario: String(form.dniBeneficiario || '').trim(),
      areaDestino: String(form.areaDestino || '').trim(),
      cargoBeneficiario: String(form.cargoBeneficiario || '').trim(),
      documentoEntrega: String(form.documentoEntrega || '').trim(),
      observacion: String(form.observacion || '').trim()
    });

    SpreadsheetApp.flush();
    return { ok: true, message: 'Movimiento registrado correctamente.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra un movimiento con múltiples productos de forma atómica.
 * batchData = { fecha, tipoMovimiento, motivo, beneficiario, dniBeneficiario,
 *               areaDestino, cargoBeneficiario, documentoEntrega,
 *               origenDestino, observacion,
 *               productos: [{ codigo, cantidad, condicion }, ...] }
 * Escribe UNA FILA POR PRODUCTO en la hoja MOVIMIENTOS, todas con el mismo MOVEMENT_ID.
 * Si falta stock para cualquier producto la operación se aborta completa (rollback).
 */
function recordMovementBatch(batchData) {
  if (!batchData || !Array.isArray(batchData.productos) || !batchData.productos.length) {
    throw new Error('Debe incluir al menos un producto en el movimiento.');
  }
  const tipo = String(batchData.tipoMovimiento || '').toUpperCase();
  const validTypes = ['ENTRADA','SALIDA','AJUSTE_POSITIVO','AJUSTE_NEGATIVO','DEVOLUCION'];
  if (!validTypes.includes(tipo)) throw new Error('Tipo de movimiento no válido: ' + tipo);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = _getDb_();
    const products = _readSheet_(ss, APP.SHEETS.PRODUCTS);
    const movements = _readSheet_(ss, APP.SHEETS.MOVEMENTS);
    const inventory = _buildInventory_(products, movements);
    const isSalida = ['SALIDA','AJUSTE_NEGATIVO'].includes(tipo);
    const fecha = batchData.fecha || _today_();
    const movementId = _id_('MOV'); // ID compartido para toda la operación

    // ── 1. VALIDAR TODOS LOS PRODUCTOS ANTES DE ESCRIBIR NADA ──
    const validatedItems = [];
    batchData.productos.forEach(function(item, idx) {
      const code = String(item.codigo || '').trim().toUpperCase();
      if (!code) throw new Error('Fila ' + (idx + 1) + ': El código de producto es obligatorio.');
      const qty = Number(item.cantidad);
      if (!(qty > 0)) throw new Error('Fila ' + (idx + 1) + ': La cantidad debe ser mayor que 0.');
      if (!Number.isInteger(qty)) throw new Error('Fila ' + (idx + 1) + ': La cantidad debe ser un número entero.');

      const product = products.find(function(p) { return String(p.CODIGO).toUpperCase() === code; });
      if (!product || String(product.ACTIVO).toUpperCase() === 'NO') {
        throw new Error('El SKU "' + code + '" no existe o está inactivo.');
      }
      if (isSalida) {
        const inv = inventory.find(function(i) { return i.codigo === code; });
        const available = inv ? inv.stock : 0;
        if (available < qty) {
          throw new Error('Stock insuficiente para "' + (product.TIPO || code) + '" (' + code + '). Disponible: ' + available + '. Solicitado: ' + qty + '.');
        }
      }
      validatedItems.push({
        codigo: code,
        cantidad: qty,
        condicion: String(item.condicion || 'NUEVO').trim()
      });
    });

    // ── 2. ESCRIBIR TODAS LAS FILAS (atómico dentro del lock) ──
    validatedItems.forEach(function(item) {
      _appendMovement_(ss, {
        movementId: movementId,
        fecha: fecha,
        codigo: item.codigo,
        tipoMovimiento: tipo,
        cantidad: item.cantidad,
        condicion: item.condicion,
        motivo: String(batchData.motivo || '').trim(),
        origenDestino: String(batchData.origenDestino || '').trim(),
        beneficiario: String(batchData.beneficiario || '').trim(),
        dniBeneficiario: String(batchData.dniBeneficiario || '').trim(),
        areaDestino: String(batchData.areaDestino || '').trim(),
        cargoBeneficiario: String(batchData.cargoBeneficiario || '').trim(),
        documentoEntrega: String(batchData.documentoEntrega || '').trim(),
        observacion: String(batchData.observacion || '').trim()
      });
    });

    SpreadsheetApp.flush();
    return {
      ok: true,
      movementId: movementId,
      count: validatedItems.length,
      message: 'Movimiento registrado correctamente. ID: ' + movementId + ' · ' + validatedItems.length + ' producto(s).'
    };
  } finally {
    lock.releaseLock();
  }
}

function getKardex(codigo, fromDate, toDate) {
  const ss = _getDb_();
  const products = _readSheet_(ss, APP.SHEETS.PRODUCTS);
  const movements = _readSheet_(ss, APP.SHEETS.MOVEMENTS).sort(_movementSort_);
  const code = String(codigo || '').trim().toUpperCase();
  if (!code) throw new Error('Selecciona un SKU.');

  const product = products.find(p => String(p.CODIGO).toUpperCase() === code);
  if (!product) throw new Error('SKU no encontrado.');

  let balance = 0;
  const rows = [];
  movements.forEach(m => {
    if (String(m.CODIGO).toUpperCase() !== code) return;
    balance += _movementDelta_(m);
    const date = String(m.FECHA || '');
    const afterFrom = !fromDate || date >= fromDate;
    const beforeTo = !toDate || date <= toDate;
    if (afterFrom && beforeTo) {
      rows.push({
        fecha: date,
        tipoMovimiento: m.TIPO_MOVIMIENTO,
        cantidad: Number(m.CANTIDAD || 0),
        entrada: _movementDelta_(m) > 0 ? Number(m.CANTIDAD || 0) : 0,
        salida: _movementDelta_(m) < 0 ? Number(m.CANTIDAD || 0) : 0,
        saldo: balance,
        condicion: m.CONDICION || '',
        motivo: m.MOTIVO || '',
        origenDestino: m.ORIGEN_DESTINO || '',
        beneficiario: m.BENEFICIARIO || '',
        documentoEntrega: m.DOCUMENTO_ENTREGA || '',
        usuario: m.USUARIO || '',
        observacion: m.OBSERVACION || ''
      });
    }
  });

  return {
    product: _publicProduct_(product),
    rows,
    currentStock: balance
  };
}

function exportInventoryCsv() {
  const data = getAppData().inventory;
  const headers = ['Código','Familia','Naturaleza','Producto','Marca','Modelo','Color','Talla','Presentación','Género','Ubicación','Unidad','Stock','Reorden','Máximo','Estado'];
  const lines = [headers.map(_csv_).join(',')];
  data.forEach(i => lines.push([
    i.codigo,i.categoria,i.naturaleza,i.tipo,i.marca,i.modelo,i.color,i.talla,i.presentacion,i.genero,i.ubicacion,i.unidad,i.stock,i.stockMin,i.stockMax,i.estado
  ].map(_csv_).join(',')));
  return lines.join('\n');
}

function exportMovementsCsv(fromDate, toDate) {
  const ss = _getDb_();
  let data = _readSheet_(ss, APP.SHEETS.MOVEMENTS).sort(_movementSort_);
  data = data.filter(m => (!fromDate || String(m.FECHA) >= fromDate) && (!toDate || String(m.FECHA) <= toDate));
  const headers = ['Fecha','Código','Movimiento','Cantidad','Condición','Motivo','Origen/Destino','Beneficiario','Documento','Usuario','Observación'];
  const lines = [headers.map(_csv_).join(',')];
  data.forEach(m => lines.push([
    m.FECHA,m.CODIGO,m.TIPO_MOVIMIENTO,m.CANTIDAD,m.CONDICION,m.MOTIVO,m.ORIGEN_DESTINO,m.BENEFICIARIO,m.DOCUMENTO_ENTREGA,m.USUARIO,m.OBSERVACION
  ].map(_csv_).join(',')));
  return lines.join('\n');
}

function _getDb_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(APP.DB_PROP);
  let ss;

  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      id = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(APP.NAME);
    id = ss.getId();
    props.setProperty(APP.DB_PROP, id);

    _ensureSheet_(ss, APP.SHEETS.PRODUCTS, APP.PRODUCT_HEADERS);
    _ensureSheet_(ss, APP.SHEETS.MOVEMENTS, APP.MOVEMENT_HEADERS);
    _ensureSheet_(ss, APP.SHEETS.CONFIG, APP.CONFIG_HEADERS);

    _writeDefaultConfig_(ss);
    _seedDemoData_(ss, true);
    SpreadsheetApp.flush();
  }

  return ss;
}

function _ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else sh.getRange(1,1,1,headers.length).setValues([headers]); // Update headers
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function _writeDefaultConfig_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.CONFIG);
  const existing = _readSheet_(ss, APP.SHEETS.CONFIG);
  if (existing.length > 0) return;
  const config = [
    ['NOMBRE_EMPRESA', 'Almacén de Prendas y Equipamiento'],
    ['MOTIVOS', 'Compra,Entrega a personal,Devolución,Reposición,Ajuste físico,Regularización,Otro'],
    ['UNIDADES', 'UNIDAD,PAR,CAJA,PAQUETE'],
    ['VERSION', '2.0.0']
  ];
  sh.getRange(2,1,config.length,2).setValues(config);
}

function _seedDemoData_(ss, force) {
  const prSh = ss.getSheetByName(APP.SHEETS.PRODUCTS);
  const mvSh = ss.getSheetByName(APP.SHEETS.MOVEMENTS);

  if (!prSh || !mvSh) return { skipped: true };

  const prValues = prSh.getDataRange().getValues();
  const headers = prValues[0] ? prValues[0].map(String) : [];
  const isV1Header = headers.includes('STOCK_MINIMO') || !headers.includes('NATURALEZA');

  if (!force && prValues.length > 1 && !isV1Header) {
    return { skipped: true };
  }

  prSh.clear();
  mvSh.clear();

  _ensureSheet_(ss, APP.SHEETS.PRODUCTS, APP.PRODUCT_HEADERS);
  _ensureSheet_(ss, APP.SHEETS.MOVEMENTS, APP.MOVEMENT_HEADERS);

  const now = new Date();
  const dateStr = _today_();

  const demoProducts = [
    [_id_('PRD'),'BOR-INS-NEG-42','CALZADO','INVENTARIABLE','BORCEGUIES','Institucional','CAT','Negro','42','NO APLICA','VARÓN','Estante A1','PAR','NUEVO',5,30,'Borceguí de cuero','SI',now],
    [_id_('PRD'),'CAS-INS-AZU-M','INDUMENTARIA','INVENTARIABLE','CASACA','Institucional','Nacional','Azul','M','NO APLICA','UNISEX','Estante B2','UNIDAD','NUEVO',8,50,'Casaca institucional azul','SI',now],
    [_id_('PRD'),'CAS-INS-AZU-L','INDUMENTARIA','INVENTARIABLE','CASACA','Institucional','Nacional','Azul','L','NO APLICA','UNISEX','Estante B2','UNIDAD','NUEVO',8,50,'Casaca institucional azul talla L','SI',now],
    [_id_('PRD'),'POL-INS-BLA-S','INDUMENTARIA','INVENTARIABLE','POLO','Institucional','Textil','Blanco','S','NO APLICA','UNISEX','Estante B3','UNIDAD','NUEVO',10,80,'Polo blanco institucional','SI',now],
    [_id_('PRD'),'POL-INS-BLA-M','INDUMENTARIA','INVENTARIABLE','POLO','Institucional','Textil','Blanco','M','NO APLICA','UNISEX','Estante B3','UNIDAD','NUEVO',10,80,'Polo blanco institucional M','SI',now],
    [_id_('PRD'),'CHA-TAC-VER-L','EPP / PROTECCIÓN','INVENTARIABLE','CHALECO TACTICO','Táctico','Tac-Force','Verde','L','NO APLICA','UNISEX','Estante C1','UNIDAD','NUEVO',3,20,'Chaleco táctico verde olivo','SI',now],
    [_id_('PRD'),'CHA-FAE-NEG-XL','EPP / PROTECCIÓN','INVENTARIABLE','CHALECO DE FAENA','Faena','SecurePro','Negro','XL','NO APLICA','UNISEX','Estante C2','UNIDAD','NUEVO',2,15,'Chaleco de faena negro XL','SI',now],
    [_id_('PRD'),'GUA-PRO-GRI-M','EPP / PROTECCIÓN','INVENTARIABLE','GUANTE','Protección','3M','Gris','M','NO APLICA','UNISEX','Estante C3','PAR','NUEVO',10,60,'Guante de seguridad industrial','SI',now],
    [_id_('PRD'),'DET-LIQ-GEN-1L','QUÍMICOS Y DESINFECCIÓN','CONSUMIBLE','DETERGENTE','Líquido','Sapolio','','NO APLICA','1 L','NO APLICA','Estante D1','UNIDAD','NUEVO',15,100,'Detergente líquido 1 litro','SI',now],
    [_id_('PRD'),'LEJ-DES-GEN-1L','QUÍMICOS Y DESINFECCIÓN','CONSUMIBLE','LEJÍA','Desinfectante','Clorox','','NO APLICA','1 L','NO APLICA','Estante D1','UNIDAD','NUEVO',10,80,'Lejía concentrada 1L','SI',now],
    [_id_('PRD'),'ESCO-LIM-GEN-STD','LIMPIEZA Y ASEO','INVENTARIABLE','ESCOBAS','Limpieza','Hude','','NO APLICA','NO APLICA','NO APLICA','Almacén D2','UNIDAD','NUEVO',5,30,'Escoba de fibra plástica','SI',now],
    [_id_('PRD'),'HOJ-GEN-GEN-A4','PAPELERÍA Y ARCHIVO','CONSUMIBLE','HOJAS A4','','Report','Blanco','NO APLICA','A4 · 500 hojas','NO APLICA','Estante E1','RESMA','NUEVO',20,200,'Resma papel bond A4 75g','SI',now],
    [_id_('PRD'),'LAP-GEN-AZU-STD','ÚTILES DE ESCRITORIO','CONSUMIBLE','LAPICERO','Escritura','Faber-Castell','Azul','NO APLICA','NO APLICA','NO APLICA','Estante E2','UNIDAD','NUEVO',30,200,'Lapicero azul punta fina','SI',now],
    [_id_('PRD'),'COR-INS-NEG-UNI','ACCESORIOS','INVENTARIABLE','CORREA','Institucional','Nacional','Negro','Única','NO APLICA','UNISEX','Estante A2','UNIDAD','NUEVO',5,30,'Correa de cuero institucional negra','SI',now],
    [_id_('PRD'),'ENG-GEN-GEN-STD','EQUIPAMIENTO DE OFICINA','INVENTARIABLE','ENGRAPADOR','Escritorio','Artesco','Negro','NO APLICA','NO APLICA','NO APLICA','Estante E3','UNIDAD','NUEVO',3,15,'Engrapador estándar de escritorio','SI',now],
  ];

  prSh.getRange(2, 1, demoProducts.length, APP.PRODUCT_HEADERS.length).setValues(demoProducts);

  const initialStock = {
    'BOR-INS-NEG-42': 8,
    'CAS-INS-AZU-M': 28,
    'CAS-INS-AZU-L': 4,
    'POL-INS-BLA-S': 25,
    'POL-INS-BLA-M': 5,
    'CHA-TAC-VER-L': 12,
    'CHA-FAE-NEG-XL': 3,
    'GUA-PRO-GRI-M': 0,
    'DET-LIQ-GEN-1L': 47,
    'LEJ-DES-GEN-1L': 0,
    'ESCO-LIM-GEN-STD': 4,
    'HOJ-GEN-GEN-A4': 63,
    'LAP-GEN-AZU-STD': 132,
    'COR-INS-NEG-UNI': 10,
    'ENG-GEN-GEN-STD': 5
  };

  const initialMovements = Object.keys(initialStock).filter(c => initialStock[c] > 0).map(c => [
    _id_('MOV'), dateStr, c, 'SALDO_INICIAL', initialStock[c], 'NUEVO', 'Carga inicial demo', 'Almacén Central', '', '', '', '', '', 'SISTEMA', 'Saldo inicial de inventario de demostración', now
  ]);

  const demoMovements = [
    [_id_('MOV'), '2026-09-18', 'BOR-INS-NEG-42', 'ENTRADA', 10, 'NUEVO', 'Compra', 'Proveedor CAT', '', '', '', '', '', 'almacenero', 'Lote septiembre', now],
    [_id_('MOV'), '2026-09-17', 'CAS-INS-AZU-M', 'SALIDA', 3, 'NUEVO', 'Entrega a personal', 'Área Operaciones', 'JUAN PEREZ QUISPE', '41872874', 'SERENAZGO', 'AGENTE', 'VALE-025', 'almacenero', 'Entrega uniforme personal nuevo', now],
    [_id_('MOV'), '2026-09-17', 'POL-INS-BLA-S', 'ENTRADA', 20, 'NUEVO', 'Compra', 'Textil SAC', '', '', '', '', '', 'almacenero', 'Reposición trimestral', now],
    [_id_('MOV'), '2026-09-16', 'DET-LIQ-GEN-1L', 'SALIDA', 5, 'NUEVO', 'Entrega a personal', 'Limpieza piso 3', 'ROSA HUANCA MAMANI', '72315890', 'LIMPIEZA', 'OPERARIA', 'VALE-026', 'almacenero', '', now],
    [_id_('MOV'), '2026-09-16', 'GUA-PRO-GRI-M', 'ENTRADA', 15, 'NUEVO', 'Compra', '3M Distribuidora', '', '', '', '', '', 'almacenero', 'Compra segundo semestre', now],
    [_id_('MOV'), '2026-09-15', 'HOJ-GEN-GEN-A4', 'SALIDA', 8, 'NUEVO', 'Entrega a personal', 'Secretaría General', 'MARIA FLORES CUEVA', '48209371', 'SECRETARIA GENERAL', 'SECRETARIA', 'PECOSA-001', 'supervisor', '', now],
    [_id_('MOV'), '2026-09-15', 'CHA-FAE-NEG-XL', 'SALIDA', 2, 'NUEVO', 'Entrega a personal', 'Seguridad Patrimonial', 'CARLOS RAMOS TORRES', '30194827', 'SEGURIDAD PATRIMONIAL', 'AGENTE', 'VALE-027', 'almacenero', 'Entrega para turno nocturno', now],
    [_id_('MOV'), '2026-09-14', 'LEJ-DES-GEN-1L', 'ENTRADA', 30, 'NUEVO', 'Compra', 'Clorox SAC', '', '', '', '', '', 'almacenero', '', now],
    [_id_('MOV'), '2026-09-14', 'LAP-GEN-AZU-STD', 'SALIDA', 12, 'NUEVO', 'Entrega a personal', 'RRHH', 'ANA GUTIERREZ SALAS', '53109284', 'RECURSOS HUMANOS', 'ASISTENTE', 'VALE-028', 'supervisor', 'Kit de escritorio nuevos empleados', now],
    [_id_('MOV'), '2026-09-13', 'POL-INS-BLA-M', 'SALIDA', 5, 'NUEVO', 'Entrega a personal', 'Logística', 'PEDRO MENDOZA LLANOS', '61023948', 'LOGISTICA', 'AUXILIAR', 'VALE-029', 'almacenero', 'Reposición uniforme', now],
    [_id_('MOV'), '2026-09-13', 'CAS-INS-AZU-L', 'AJUSTE_NEGATIVO', 1, 'USADO / REUTILIZABLE', 'Ajuste físico', 'Almacén', '', '', '', '', '', 'supervisor', 'Prenda deteriorada detectada en inventario', now],
    [_id_('MOV'), '2026-09-12', 'ESCO-LIM-GEN-STD', 'ENTRADA', 10, 'NUEVO', 'Compra', 'Distribuidora Hude', '', '', '', '', '', 'almacenero', '', now],
    [_id_('MOV'), '2026-09-11', 'COR-INS-NEG-UNI', 'SALIDA', 4, 'NUEVO', 'Entrega a personal', 'Vigilancia', 'LUIS VARGAS CONDORI', '70293847', 'VIGILANCIA', 'VIGILANTE', 'VALE-030', 'almacenero', '', now],
    [_id_('MOV'), '2026-09-10', 'ENG-GEN-GEN-STD', 'ENTRADA', 5, 'NUEVO', 'Compra', 'Artesco SAC', '', '', '', '', '', 'almacenero', '', now],
    [_id_('MOV'), '2026-09-09', 'BOR-INS-NEG-42', 'DEVOLUCION', 1, 'USADO / REUTILIZABLE', 'Devolución', 'Serenazgo', 'JUAN PEREZ QUISPE', '41872874', 'SERENAZGO', 'AGENTE', 'VALE-020', 'supervisor', 'Devolución de borceguí talla 42 por talla incorrecta', now],
  ];

  const allMovements = initialMovements.concat(demoMovements);
  mvSh.getRange(2, 1, allMovements.length, APP.MOVEMENT_HEADERS.length).setValues(allMovements);

  return { ok: true, count: demoProducts.length };
}

function _appendMovement_(ss, item) {
  const sh = ss.getSheetByName(APP.SHEETS.MOVEMENTS);
  const email = Session.getActiveUser().getEmail() || 'APP';
  
  const data = {
    'ID': _id_('ROW'),
    'MOVEMENT_ID': item.movementId || '',
    'FECHA': item.fecha,
    'CODIGO': item.codigo,
    'TIPO_MOVIMIENTO': item.tipoMovimiento,
    'CANTIDAD': Number(item.cantidad),
    'CONDICION': item.condicion || '',
    'MOTIVO': item.motivo || '',
    'ORIGEN_DESTINO': item.origenDestino || '',
    'BENEFICIARIO': item.beneficiario || '',
    'DNI_BENEFICIARIO': item.dniBeneficiario || '',
    'AREA_DESTINO': item.areaDestino || '',
    'CARGO_BENEFICIARIO': item.cargoBeneficiario || '',
    'DOCUMENTO_ENTREGA': item.documentoEntrega || '',
    'USUARIO': email,
    'OBSERVACION': item.observacion || '',
    'CREADO_EN': new Date()
  };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  headers.forEach((h, i) => {
    const key = String(h).trim().toUpperCase();
    if (data[key] !== undefined) {
      row[i] = data[key];
    }
  });
  sh.appendRow(row);
}

function _readSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('No existe la hoja ' + sheetName + '. Ejecuta setupSistema().');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function _buildInventory_(products, movements) {
  const map = {};
  products.forEach(p => {
    if (String(p.ACTIVO).toUpperCase() === 'NO') return;
    map[String(p.CODIGO).toUpperCase()] = {
      codigo: String(p.CODIGO).toUpperCase(),
      categoria: p.CATEGORIA || '', naturaleza: p.NATURALEZA || '', tipo: p.TIPO || '', modelo: p.MODELO || '', marca: p.MARCA || '', color: p.COLOR || '', talla: p.TALLA || '', presentacion: p.PRESENTACION || '', genero: p.GENERO || '', ubicacion: p.UBICACION || '', condicion: p.CONDICION || '',
      descripcion: p.DESCRIPCION || '', unidad: p.UNIDAD || 'UNIDAD', stockMin: Number(p.PUNTO_REORDEN || 0), stockMax: Number(p.STOCK_MAXIMO || 0), stock: 0,
      estado: 'NORMAL'
    };
  });
  movements.forEach(m => {
    const key = String(m.CODIGO || '').toUpperCase();
    if (!map[key]) return;
    map[key].stock += _movementDelta_(m);
  });
  return Object.values(map).map(i => {
    i.stock = Math.round((i.stock + Number.EPSILON) * 100) / 100;
    i.estado = i.stock <= 0 ? 'AGOTADO' : (i.stock <= i.stockMin ? (i.stock <= i.stockMin / 2 ? 'STOCK BAJO' : 'REORDENAR') : (i.stockMax > 0 && i.stock > i.stockMax ? 'SOBRESTOCK' : 'NORMAL'));
    return i;
  });
}

function _movementDelta_(m) {
  const qty = Number(m.CANTIDAD || 0);
  return ['ENTRADA','SALDO_INICIAL','AJUSTE_POSITIVO'].includes(String(m.TIPO_MOVIMIENTO).toUpperCase()) ? qty : -qty;
}

function _movementSort_(a,b) {
  const da = new Date(a.CREADO_EN || a.FECHA || 0).getTime();
  const db = new Date(b.CREADO_EN || b.FECHA || 0).getTime();
  return da - db;
}

function _publicProduct_(p) {
  return {
    codigo: String(p.CODIGO || ''), categoria: String(p.CATEGORIA || ''), naturaleza: String(p.NATURALEZA || ''), tipo: String(p.TIPO || ''), modelo: String(p.MODELO || ''), marca: String(p.MARCA || ''), color: String(p.COLOR || ''),
    talla: String(p.TALLA || ''), presentacion: String(p.PRESENTACION || ''), genero: String(p.GENERO || ''), ubicacion: String(p.UBICACION || ''), unidad: String(p.UNIDAD || 'UNIDAD'), condicion: String(p.CONDICION || ''),
    stockMin: Number(p.PUNTO_REORDEN || 0), stockMax: Number(p.STOCK_MAXIMO || 0), descripcion: String(p.DESCRIPCION || ''), activo: String(p.ACTIVO || 'SI')
  };
}

function _publicMovement_(m) {
  return {
    movementId: String(m.MOVEMENT_ID || ''),
    fecha: _formatDateValue_(m.FECHA), codigo: String(m.CODIGO || ''), tipoMovimiento: String(m.TIPO_MOVIMIENTO || ''),
    cantidad: Number(m.CANTIDAD || 0), condicion: String(m.CONDICION || ''), motivo: String(m.MOTIVO || ''), origenDestino: String(m.ORIGEN_DESTINO || ''),
    beneficiario: String(m.BENEFICIARIO || ''), dniBeneficiario: String(m.DNI_BENEFICIARIO || ''), areaDestino: String(m.AREA_DESTINO || ''), cargoBeneficiario: String(m.CARGO_BENEFICIARIO || ''), documentoEntrega: String(m.DOCUMENTO_ENTREGA || ''),
    usuario: String(m.USUARIO || ''), observacion: String(m.OBSERVACION || '')
  };
}

function _getConfigList_(ss, key) {
  const config = _readSheet_(ss, APP.SHEETS.CONFIG);
  const row = config.find(r => String(r.CLAVE).toUpperCase() === key);
  return row ? String(row.VALOR || '').split(',').map(v => v.trim()).filter(Boolean) : [];
}

function _unique_(arr) {
  return [...new Set(arr.map(v => String(v || '').trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b,'es'));
}

function _validateProduct_(f) {
  const required = [['codigo','Código'],['tipo','Tipo']]; // Removed talla and modelo as they are optional now
  required.forEach(([k,label]) => { if (!String(f[k] || '').trim()) throw new Error(label + ' es obligatorio.'); });
  if (Number(f.stockInicial || 0) < 0) throw new Error('El stock inicial no puede ser negativo.');
  if (Number(f.puntoReorden || 0) < 0) throw new Error('El stock mínimo / punto de reorden no puede ser negativo.');
}

function _validateMovement_(f) {
  if (!String(f.codigo || '').trim()) throw new Error('Selecciona un producto.');
  const types = ['ENTRADA','SALIDA','AJUSTE_POSITIVO','AJUSTE_NEGATIVO','DEVOLUCION'];
  if (!types.includes(String(f.tipoMovimiento || '').toUpperCase())) throw new Error('Tipo de movimiento no válido.');
  if (!(Number(f.cantidad) > 0)) throw new Error('La cantidad debe ser mayor que 0.');
  // if (!String(f.motivo || '').trim()) throw new Error('Indica el motivo.');
}

function _id_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
}

function _today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _formatDateValue_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function _csv_(value) {
  const s = String(value == null ? '' : value).replace(/"/g, '""');
  return '"' + s + '"';
}

function _resolveColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim().toUpperCase()] = i + 1; // 1-based index for getRange
  });
  return map;
}

function _updateRowBySchema_(sheet, rowNumber, dataMap, columnMap) {
  Object.keys(dataMap).forEach(key => {
    const colIndex = columnMap[key.toUpperCase()];
    if (colIndex) {
      sheet.getRange(rowNumber, colIndex).setValue(dataMap[key]);
    }
  });
}
