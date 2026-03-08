/**
 * netlify/functions/generar.js
 *
 * Netlify Serverless Function que:
 *  1. Recibe los datos del formulario como JSON
 *  2. Valida los datos con seguridad básica
 *  3. Carga la plantilla.xlsx existente (en la raíz del proyecto)
 *  4. Inserta valores en las celdas exactas sin modificar formatos
 *  5. Devuelve el buffer Excel para descarga
 */

'use strict';

// se agregaron estas 3 lineas sig.:
const GITHUB_USER     = 'guionbajox2';        // ← tu usuario de GitHub
const GITHUB_REPO     = 'sistemapresupuestos';    // ← nombre del repo
const GITHUB_BRANCH   = 'main';              // ← 'main' o 'master'

// se eliminaron:
// const ExcelJS = require('exceljs');
// const path    = require('path');
// const fs = require('fs');

/* =============================================
   MAPA DE CELDAS
   Cada clave corresponde al campo del formulario
   y su valor es la celda de destino en la plantilla.
   ============================================= */
const MAPA_CELDAS = {
  fechaSolicitud:    'L2',
  organizacion:      'H1',
  nombreLider:       'D22',
  nombreReceptor_L8: 'L8',
  nombreReceptor_G22:'G22',
  tipoSolicitud:     'J9',
  nombreActividad:   'E5',
  fechaActividad:    'C6',
  horaActividad:     'H6',
  lugar:             'D7',
  cantidadAsistencia:'B8',
  dirigidaA:         'C9',
  proposito:         'A12',

  // Metas (marcar X)
  meta_edificarFe:         'J18',
  meta_diversionUnidad:    'J19',
  meta_crecimientoPersonal:'J21',
  meta_fortalecerFamilias: 'J22',
  meta_obrasSalvacion:     'J23',

  // Productos: filas 29..41
  // Se insertan dinámicamente dentro del bucle

  // Totales
  totalSolicitado_K42: 'K42',
  totalSolicitado_L7:  'L7',
};

/* Celdas de inicio para los productos */
const PRODUCTO_INICIO_FILA = 29;
const COLUMNA_CANTIDAD      = 'A'; // columna A
const COLUMNA_PRODUCTO      = 'B'; // columna B
const COLUMNA_COSTO_UNIT    = 'J'; // columna J
const COLUMNA_COSTO_TOTAL   = 'K'; // columna K

/* =============================================
   SANITIZACIÓN BÁSICA
   ============================================= */

/**
 * Elimina caracteres potencialmente peligrosos en strings.
 * Para Excel: previene inyección de fórmulas (=, +, -, @).
 */
function sanitizeString(val) {
  if (typeof val !== 'string') return '';
  // Truncar a 500 caracteres máximo
  let s = val.slice(0, 500).trim();
  // Prevenir inyección de fórmulas Excel
  if (s.startsWith('=') || s.startsWith('+') || s.startsWith('-') || s.startsWith('@')) {
    s = "'" + s;
  }
  return s;
}

function sanitizeNumber(val, min = 0, max = 9999999) {
  const n = parseFloat(val);
  if (isNaN(n)) return 0;
  return Math.max(min, Math.min(max, n));
}

/* =============================================
   GENERAR TEXTO DINÁMICO TIPO SOLICITUD
   ============================================= */
function generarTextoTipoSolicitud(tipo) {
  if (tipo === 'Anticipo') {
    return 'ANTICIPO | X | REEMBOLSO |   |';
  } else if (tipo === 'Reembolso') {
    return 'ANTICIPO |   | REEMBOLSO | X |';
  }
  return '';
}

/* =============================================
   INSERTAR VALOR EN CELDA SIN MODIFICAR ESTILO
   ============================================= */

/**
 * Escribe un valor en una celda preservando completamente
 * el estilo y formato existente de la plantilla.
 *
 * @param {ExcelJS.Worksheet} ws - Hoja de trabajo
 * @param {string} cellRef - Referencia de celda ej: 'A1'
 * @param {string|number} value - Valor a insertar
 */
function setCellValue(ws, cellRef, value) {
  const cell = ws.getCell(cellRef);
  // Guardar estilo actual
  const existingStyle = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  // Asignar valor
  cell.value = value;
  // Restaurar estilo
  cell.style = existingStyle;
}

/* =============================================
   VALIDACIÓN DE DATOS RECIBIDOS
   ============================================= */
function validarDatos(data) {
  const errores = [];

  if (!data.fechaSolicitud)      errores.push('fechaSolicitud es obligatorio');
  if (!data.organizacion)        errores.push('organizacion es obligatorio');
  if (!data.nombreLider)         errores.push('nombreLider es obligatorio');
  if (!data.nombreReceptor)      errores.push('nombreReceptor es obligatorio');
  if (!data.tipoSolicitud)       errores.push('tipoSolicitud es obligatorio');
  if (!data.nombreActividad)     errores.push('nombreActividad es obligatorio');
  if (!data.fechaActividad)      errores.push('fechaActividad es obligatorio');
  if (!data.lugar)               errores.push('lugar es obligatorio');
  if (!data.proposito)           errores.push('proposito es obligatorio');
  if (!data.cantidadAsistencia || data.cantidadAsistencia < 1)
    errores.push('cantidadAsistencia debe ser >= 1');
  if (!Array.isArray(data.metas) || data.metas.length === 0)
    errores.push('Debe seleccionar al menos una meta');
  if (!Array.isArray(data.productos) || data.productos.length === 0)
    errores.push('Debe incluir al menos un producto');
  if (data.productos.length > 13)
    errores.push('Máximo 13 productos permitidos');

  return errores;
}

/* =============================================
   HANDLER PRINCIPAL
   ============================================= */
exports.handler = async function(event, context) {

  // Solo aceptar POST
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Método no permitido. Use POST.' }),
    };
  }

  /* --- 1. Parsear body con seguridad --- */
  let data;
  try {
    if (!event.body) throw new Error('El cuerpo de la solicitud está vacío.');
    data = JSON.parse(event.body);
  } catch (err) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'JSON inválido: ' + err.message }),
    };
  }

  /* --- 2. Validar datos --- */
  const errores = validarDatos(data);
  if (errores.length > 0) {
    return {
      statusCode: 422,
      body: JSON.stringify({ error: 'Datos inválidos: ' + errores.join('; ') }),
    };
  }

  try {
    /* --- 3. Cargar plantilla ---
       Netlify CLI y producción resuelven __dirname de forma distinta,
       por eso probamos varias rutas posibles hasta encontrar el archivo.
    ------------------------------------------------------------------ */
    const posiblesPaths = [
      // Producción Netlify: la función se ejecuta desde la raíz
      path.join(__dirname, '..', '..', 'plantilla.xlsx'),
      // Netlify CLI local (a veces agrega un nivel extra)
      path.join(__dirname, '..', '..', '..', 'plantilla.xlsx'),
      // Ruta absoluta desde process.cwd() — raíz real del proyecto
      path.join(process.cwd(), 'plantilla.xlsx'),
      // Por si cwd apunta a netlify/functions
      path.join(process.cwd(), '..', '..', 'plantilla.xlsx'),
    ];

    // Imprimir rutas intentadas para facilitar depuración
    console.log('[generar.js] Buscando plantilla en:');
    posiblesPaths.forEach(p => console.log('  -', p, '→', fs.existsSync(p) ? '✓ EXISTE' : '✗ no encontrada'));

    const templatePath = posiblesPaths.find(p => fs.existsSync(p));

    if (!templatePath) {
      return {
        statusCode: 500,
        body: JSON.stringify({
          error: 'plantilla.xlsx no encontrada en el servidor. Verifique que el archivo exista en la raíz del proyecto.',
          rutas_revisadas: posiblesPaths,
        }),
      };
    }

    console.log('[generar.js] Plantilla encontrada en:', templatePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Obtener la primera hoja (ajustar índice/nombre si es necesario)
    const ws = workbook.worksheets[0];
    if (!ws) {
      return {
        statusCode: 500,
        body: JSON.stringify({ error: 'La plantilla no contiene ninguna hoja de trabajo.' }),
      };
    }

    /* --- 4. Insertar valores en celdas --- */

    // Información general
    setCellValue(ws, MAPA_CELDAS.fechaSolicitud,    sanitizeString(data.fechaSolicitud));
    setCellValue(ws, MAPA_CELDAS.organizacion,      sanitizeString(data.organizacion));
    setCellValue(ws, MAPA_CELDAS.nombreLider,       sanitizeString(data.nombreLider));
    setCellValue(ws, MAPA_CELDAS.nombreReceptor_L8, sanitizeString(data.nombreReceptor));
    setCellValue(ws, MAPA_CELDAS.nombreReceptor_G22,sanitizeString(data.nombreReceptor));

    // Tipo de solicitud: texto dinámico
    const textoTipo = generarTextoTipoSolicitud(sanitizeString(data.tipoSolicitud));
    setCellValue(ws, MAPA_CELDAS.tipoSolicitud, textoTipo);

    // Detalles de la actividad
    setCellValue(ws, MAPA_CELDAS.nombreActividad,    sanitizeString(data.nombreActividad));
    setCellValue(ws, MAPA_CELDAS.fechaActividad,     sanitizeString(data.fechaActividad));
    setCellValue(ws, MAPA_CELDAS.horaActividad,      sanitizeString(data.horaActividad));
    setCellValue(ws, MAPA_CELDAS.lugar,              sanitizeString(data.lugar));
    setCellValue(ws, MAPA_CELDAS.cantidadAsistencia, sanitizeNumber(data.cantidadAsistencia, 1));
    setCellValue(ws, MAPA_CELDAS.dirigidaA,          sanitizeString(data.dirigidaA));
    setCellValue(ws, MAPA_CELDAS.proposito,          sanitizeString(data.proposito));

    // Metas: insertar 'X' en celdas correspondientes
    const METAS_MAP = {
      'edificarFe':          MAPA_CELDAS.meta_edificarFe,
      'diversionUnidad':     MAPA_CELDAS.meta_diversionUnidad,
      'crecimientoPersonal': MAPA_CELDAS.meta_crecimientoPersonal,
      'fortalecerFamilias':  MAPA_CELDAS.meta_fortalecerFamilias,
      'obrasSalvacion':      MAPA_CELDAS.meta_obrasSalvacion,
    };

    // Primero limpiar todas las celdas de metas (por si la plantilla tiene algo)
    Object.values(METAS_MAP).forEach(celda => setCellValue(ws, celda, ''));

    // Marcar las seleccionadas
    const metas = Array.isArray(data.metas) ? data.metas : [];
    metas.forEach(meta => {
      if (meta === 'otra') {
        // "Otra" meta: insertar texto personalizado (por ejemplo en J24 — ajustar si es necesario)
        // Si hay texto de otra meta, se inserta en la celda de metas disponible
        if (data.metaOtraTexto) {
          // Insertar 'X' y el texto en una celda cercana según diseño de plantilla
          // NOTA: ajustar la celda si la plantilla tiene una fila específica para "Otra"
          setCellValue(ws, 'J25', 'X');
          setCellValue(ws, 'K25', sanitizeString(data.metaOtraTexto));
        }
      } else if (METAS_MAP[meta]) {
        setCellValue(ws, METAS_MAP[meta], 'X');
      }
    });

    /* --- 5. Lista de productos --- */
    const productos = Array.isArray(data.productos) ? data.productos.slice(0, 13) : [];

    productos.forEach((prod, idx) => {
      const fila = PRODUCTO_INICIO_FILA + idx; // 29..41
      setCellValue(ws, `${COLUMNA_CANTIDAD}${fila}`,   sanitizeNumber(prod.cantidad, 0));
      setCellValue(ws, `${COLUMNA_PRODUCTO}${fila}`,   sanitizeString(prod.producto));
      setCellValue(ws, `${COLUMNA_COSTO_UNIT}${fila}`, sanitizeNumber(prod.costoUnitario, 0));
      setCellValue(ws, `${COLUMNA_COSTO_TOTAL}${fila}`,sanitizeNumber(prod.costoTotal, 0));
    });

    /* --- 6. Total general --- */
    const total = sanitizeNumber(data.totalSolicitado, 0);
    setCellValue(ws, MAPA_CELDAS.totalSolicitado_K42, total);
    setCellValue(ws, MAPA_CELDAS.totalSolicitado_L7,  total);

    /* --- 7. Generar buffer y devolver --- */
    const buffer = await workbook.xlsx.writeBuffer();

    return {
      statusCode: 200,
      headers: {
        'Content-Type':        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="solicitud_presupuesto.xlsx"',
        'Content-Length':      buffer.length.toString(),
        'Cache-Control':       'no-store',
        // Seguridad básica
        'X-Content-Type-Options': 'nosniff',
      },
      // Netlify Functions soporta body binario con isBase64Encoded
      body:            buffer.toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('[generar.js] Error interno:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: 'Error interno al generar el archivo Excel. Por favor intente nuevamente.',
        detail: process.env.NODE_ENV === 'development' ? err.message : undefined,
      }),
    };
  }
};
