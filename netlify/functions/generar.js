'use strict';

const ExcelJS = require('exceljs');

const GITHUB_USER    = 'guionbajox2';
const GITHUB_REPO    = 'sistemapresupuestos';
const GITHUB_BRANCH  = 'main';
const PLANTILLA_PATH = 'netlify/functions/plantilla.xlsx';

const PLANTILLA_URL = 'https://raw.githubusercontent.com/' + GITHUB_USER + '/' + GITHUB_REPO + '/' + GITHUB_BRANCH + '/' + PLANTILLA_PATH;

const MAPA_CELDAS = {
  fechaSolicitud:           'L2',
  organizacion:             'H1',
  nombreLider:              'D22',
  nombreReceptor_L8:        'L8',
  nombreReceptor_G22:       'G22',
  tipoSolicitud:            'J9',
  nombreActividad:          'E5',
  fechaActividad:           'C6',
  horaActividad:            'H6',
  lugar:                    'D7',
  cantidadAsistencia:       'B8',
  dirigidaA:                'C9',
  proposito:                'A12',
  meta_edificarFe:          'J18',
  meta_diversionUnidad:     'J19',
  meta_crecimientoPersonal: 'J21',
  meta_fortalecerFamilias:  'J22',
  meta_obrasSalvacion:      'J23',
  totalSolicitado_K42:      'K42',
  totalSolicitado_L7:       'L7',
};

const PRODUCTO_INICIO_FILA = 29;

function sanitizeString(val) {
  if (typeof val !== 'string') return '';
  var s = val.slice(0, 500).trim();
  if (s.startsWith('=') || s.startsWith('+') || s.startsWith('-') || s.startsWith('@')) s = "'" + s;
  return s;
}

function sanitizeNumber(val, min, max) {
  if (min === undefined) min = 0;
  if (max === undefined) max = 9999999;
  var n = parseFloat(val);
  if (isNaN(n)) return 0;
  return Math.max(min, Math.min(max, n));
}

function generarTextoTipoSolicitud(tipo) {
  if (tipo === 'Anticipo')  return 'ANTICIPO | X | REEMBOLSO |   |';
  if (tipo === 'Reembolso') return 'ANTICIPO |   | REEMBOLSO | X |';
  return '';
}

function setCellValue(ws, cellRef, value) {
  var cell = ws.getCell(cellRef);
  var style = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  cell.value = value;
  cell.style = style;
}

function validarDatos(data) {
  var errores = [];
  if (!data.fechaSolicitud)  errores.push('fechaSolicitud requerido');
  if (!data.organizacion)    errores.push('organizacion requerido');
  if (!data.dirigidaA)    errores.push('Actividad dirigida a requerido');
  if (!data.nombreLider)     errores.push('nombreLider requerido');
  if (!data.nombreReceptor)  errores.push('nombreReceptor requerido');
  if (!data.tipoSolicitud)   errores.push('tipoSolicitud requerido');
  if (!data.nombreActividad) errores.push('nombreActividad requerido');
  if (!data.fechaActividad)  errores.push('fechaActividad requerido');
  if (!data.lugar)           errores.push('lugar requerido');
  if (!data.proposito)       errores.push('proposito requerido');
  if (!data.cantidadAsistencia || data.cantidadAsistencia < 1)
    errores.push('cantidadAsistencia debe ser >= 1');
  if (!Array.isArray(data.metas) || data.metas.length === 0)
    errores.push('Debe seleccionar al menos una meta');
  if (!Array.isArray(data.productos) || data.productos.length === 0)
    errores.push('Debe incluir al menos un producto');
  if (data.productos.length > 13)
    errores.push('Maximo 13 productos');
  return errores;
}

exports.handler = async function(event) {

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Metodo no permitido.' }) };
  }

  var data;
  try {
    if (!event.body) throw new Error('Cuerpo vacio.');
    data = JSON.parse(event.body);
  } catch (err) {
    return { statusCode: 400, body: JSON.stringify({ error: 'JSON invalido: ' + err.message }) };
  }

  var errores = validarDatos(data);
  if (errores.length > 0) {
    return { statusCode: 422, body: JSON.stringify({ error: errores.join('; ') }) };
  }

  try {
    console.log('[generar.js] Descargando plantilla desde:', PLANTILLA_URL);
    var response = await fetch(PLANTILLA_URL);
    if (!response.ok) {
      throw new Error('No se pudo descargar la plantilla. Status: ' + response.status);
    }
    var arrayBuffer = await response.arrayBuffer();
    var buffer = Buffer.from(arrayBuffer);

    var workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    var ws = workbook.worksheets[0];
    if (!ws) throw new Error('La plantilla no tiene hojas de trabajo.');

    setCellValue(ws, MAPA_CELDAS.fechaSolicitud,     sanitizeString(data.fechaSolicitud));
    setCellValue(ws, MAPA_CELDAS.organizacion,       sanitizeString(data.organizacion));
    setCellValue(ws, MAPA_CELDAS.nombreLider,        sanitizeString(data.nombreLider));
    setCellValue(ws, MAPA_CELDAS.nombreReceptor_L8,  sanitizeString(data.nombreReceptor));
    setCellValue(ws, MAPA_CELDAS.nombreReceptor_G22, sanitizeString(data.nombreReceptor));
    setCellValue(ws, MAPA_CELDAS.tipoSolicitud,      generarTextoTipoSolicitud(sanitizeString(data.tipoSolicitud)));
    setCellValue(ws, MAPA_CELDAS.nombreActividad,    sanitizeString(data.nombreActividad));
    setCellValue(ws, MAPA_CELDAS.fechaActividad,     sanitizeString(data.fechaActividad));
    setCellValue(ws, MAPA_CELDAS.horaActividad,      sanitizeString(data.horaActividad));
    setCellValue(ws, MAPA_CELDAS.lugar,              sanitizeString(data.lugar));
    setCellValue(ws, MAPA_CELDAS.cantidadAsistencia, sanitizeNumber(data.cantidadAsistencia, 1));
    setCellValue(ws, MAPA_CELDAS.dirigidaA,          sanitizeString(data.dirigidaA));
    setCellValue(ws, MAPA_CELDAS.proposito,          sanitizeString(data.proposito));

    var METAS_MAP = {
      'edificarFe':          MAPA_CELDAS.meta_edificarFe,
      'diversionUnidad':     MAPA_CELDAS.meta_diversionUnidad,
      'crecimientoPersonal': MAPA_CELDAS.meta_crecimientoPersonal,
      'fortalecerFamilias':  MAPA_CELDAS.meta_fortalecerFamilias,
      'obrasSalvacion':      MAPA_CELDAS.meta_obrasSalvacion,
    };

    Object.values(METAS_MAP).forEach(function(c) { setCellValue(ws, c, ''); });

    (Array.isArray(data.metas) ? data.metas : []).forEach(function(meta) {
      if (meta === 'otra') {
        if (data.metaOtraTexto) {
          setCellValue(ws, 'J25', 'X');
          setCellValue(ws, 'K25', sanitizeString(data.metaOtraTexto));
        }
      } else if (METAS_MAP[meta]) {
        setCellValue(ws, METAS_MAP[meta], 'X');
      }
    });

    (Array.isArray(data.productos) ? data.productos.slice(0, 13) : []).forEach(function(prod, idx) {
      var f = PRODUCTO_INICIO_FILA + idx;
      setCellValue(ws, 'A' + f, sanitizeNumber(prod.cantidad, 0));
      setCellValue(ws, 'B' + f, sanitizeString(prod.producto));
      setCellValue(ws, 'J' + f, sanitizeNumber(prod.costoUnitario, 0));
      setCellValue(ws, 'K' + f, sanitizeNumber(prod.costoTotal, 0));
    });

    var total = sanitizeNumber(data.totalSolicitado, 0);
    setCellValue(ws, MAPA_CELDAS.totalSolicitado_K42, total);
    setCellValue(ws, MAPA_CELDAS.totalSolicitado_L7,  total);

    var outputBuffer = await workbook.xlsx.writeBuffer();

    return {
      statusCode: 200,
      headers: {
        'Content-Type':           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':    'attachment; filename="solicitud_presupuesto.xlsx"',
        'Content-Length':         outputBuffer.length.toString(),
        'Cache-Control':          'no-store',
        'X-Content-Type-Options': 'nosniff',
      },
      body:            Buffer.from(outputBuffer).toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('[generar.js] Error interno:', err.message);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
