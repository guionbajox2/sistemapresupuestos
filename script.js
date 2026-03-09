/**
 * script.js — Solicitud de Presupuesto
 * Maneja: validaciones en tiempo real, tabla dinámica,
 * cálculos automáticos y envío al backend Netlify Function.
 */

'use strict';

/* =============================================
   CONSTANTES Y CONFIGURACIÓN
   ============================================= */

const MAX_ROWS = 13;   // Filas 29..41 en Excel (13 filas)

/* Nombres únicos para atributos ARIA de cada fila */
let rowCounter = 0;

/* =============================================
   UTILIDADES GENERALES
   ============================================= */

/**
 * Selecciona un elemento por selector CSS, lanzando error si no existe.
 */
const $ = (selector, ctx = document) => ctx.querySelector(selector);
const $$ = (selector, ctx = document) => [...ctx.querySelectorAll(selector)];

/**
 * Formatea un número como "1,234.50 Bs."
 * Si el valor es 0 muestra "0 Bs."
 */
function formatBs(value) {
  const n = parseFloat(value) || 0;
  if (n === 0) return '0 Bs.';
  return n.toLocaleString('es-BO', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' Bs.';
}

/* =============================================
   TOOLTIP GLOBAL
   ============================================= */

const tooltip = $('#globalTooltip');
let tooltipTimeout;

function showTooltip(el, text) {
  clearTimeout(tooltipTimeout);
  tooltip.textContent = text;
  tooltip.setAttribute('aria-hidden', 'false');
  tooltip.classList.add('visible');
  positionTooltip(el);
}

function hideTooltip() {
  tooltipTimeout = setTimeout(() => {
    tooltip.classList.remove('visible');
    tooltip.setAttribute('aria-hidden', 'true');
  }, 120);
}

function positionTooltip(el) {
  const rect = el.getBoundingClientRect();
  const tw = tooltip.offsetWidth;
  let left = rect.left + rect.width / 2 - tw / 2;
  let top  = rect.top - tooltip.offsetHeight - 10 + window.scrollY;

  // Evitar salirse a la derecha
  if (left + tw > window.innerWidth - 12) left = window.innerWidth - tw - 12;
  if (left < 8) left = 8;
  // Si no hay espacio arriba, poner abajo
  if (top < window.scrollY + 8) top = rect.bottom + 8 + window.scrollY;

  tooltip.style.left = left + 'px';
  tooltip.style.top  = top + 'px';
}

// Registrar eventos para todos los info-btn
function initTooltips() {
  document.addEventListener('mouseover', e => {
    const btn = e.target.closest('.info-btn');
    if (btn && btn.dataset.tooltip) showTooltip(btn, btn.dataset.tooltip);
  });
  document.addEventListener('mouseout', e => {
    if (e.target.closest('.info-btn')) hideTooltip();
  });
  document.addEventListener('focusin', e => {
    const btn = e.target.closest('.info-btn');
    if (btn && btn.dataset.tooltip) showTooltip(btn, btn.dataset.tooltip);
  });
  document.addEventListener('focusout', e => {
    if (e.target.closest('.info-btn')) hideTooltip();
  });
}

/* =============================================
   VALIDACIÓN EN TIEMPO REAL
   ============================================= */

/**
 * Muestra un mensaje de error en el elemento de error asociado.
 * También aplica clase visual al input.
 */
function setFieldError(fieldId, message) {
  const field = document.getElementById(fieldId);
  const errorEl = document.getElementById(fieldId + '-error');
  if (!field || !errorEl) return;

  errorEl.textContent = message;
  field.classList.remove('is-valid');
  if (message) {
    field.classList.add('is-invalid');
    field.setAttribute('aria-invalid', 'true');
  } else {
    field.classList.remove('is-invalid');
    field.removeAttribute('aria-invalid');
    if (field.value.trim()) field.classList.add('is-valid');
  }
}

/** Limpia el error de un campo */
function clearFieldError(fieldId) {
  setFieldError(fieldId, '');
}

/** Valida que un input texto no esté vacío */
function validateRequired(fieldId, label) {
  const field = document.getElementById(fieldId);
  if (!field) return false;
  const val = field.value.trim();
  if (!val) {
    setFieldError(fieldId, `⚠ El campo "${label}" es obligatorio.`);
    return false;
  }
  clearFieldError(fieldId);
  return true;
}

/** Valida un número positivo */
function validatePositiveNumber(fieldId, label) {
  const field = document.getElementById(fieldId);
  if (!field) return false;
  const val = parseFloat(field.value);
  if (isNaN(val) || val < 1) {
    setFieldError(fieldId, `⚠ "${label}" debe ser un número mayor o igual a 1.`);
    return false;
  }
  clearFieldError(fieldId);
  return true;
}

/** Valida que un select tenga valor */
function validateSelect(fieldId, label) {
  const field = document.getElementById(fieldId);
  if (!field) return false;
  if (!field.value) {
    setFieldError(fieldId, `⚠ Seleccione una opción para "${label}".`);
    return false;
  }
  clearFieldError(fieldId);
  return true;
}

/* Registrar validación en tiempo real para campos texto/select */
function initFieldValidation() {
  // Texto / datalist
  const textFields = [
    { id: 'fechaSolicitud',    label: 'Fecha de la Solicitud' },
    { id: 'nombreLider',       label: 'Nombre del Líder' },
    { id: 'nombreReceptor',    label: 'Nombre de quien recibirá el presupuesto' },
    { id: 'nombreActividad',   label: 'Nombre de la Actividad' },
    { id: 'fechaActividad',    label: 'Fecha de la Actividad' },
    { id: 'lugar',             label: 'Lugar' },
    { id: 'proposito',         label: 'Propósito de la Actividad' },
  ];

  textFields.forEach(({ id, label }) => {
    const field = document.getElementById(id);
    if (!field) return;
    field.addEventListener('input', () => {
      if (field.value.trim()) clearFieldError(id);
    });
    field.addEventListener('blur', () => validateRequired(id, label));
  });

  // Select organización
  const org = document.getElementById('organizacion');
  org.addEventListener('change', () => validateSelect('organizacion', 'Organización'));

  // Select Dirigida A
  const dirigidaA = document.getElementById('dirigidaA');
  dirigidaA.addEventListener('change', () => validateSelect('dirigidaA', 'Actividad Dirigida a...'));

  // Number
  const qty = document.getElementById('cantidadAsistencia');
  qty.addEventListener('input', () => validatePositiveNumber('cantidadAsistencia', 'Cantidad de asistencia'));
  qty.addEventListener('blur',  () => validatePositiveNumber('cantidadAsistencia', 'Cantidad de asistencia'));

  // Radio: tipo solicitud
  $$('input[name="tipoSolicitud"]').forEach(radio => {
    radio.addEventListener('change', () => {
      const errorEl = document.getElementById('tipoSolicitud-error');
      if (errorEl) errorEl.textContent = '';
    });
  });
}

/* =============================================
   CHECKBOX: META "OTRA"
   ============================================= */

function initOtraCheckbox() {
  const otraCheck = document.getElementById('metaOtraCheck');
  const otraInput = document.getElementById('metaOtraTexto');
  if (!otraCheck || !otraInput) return;

  otraCheck.addEventListener('change', () => {
    otraInput.disabled = !otraCheck.checked;
    if (otraCheck.checked) {
      otraInput.focus();
    } else {
      otraInput.value = '';
    }
  });
}

/* =============================================
   TABLA DINÁMICA DE COMPRAS
   ============================================= */

function createTableRow() {
  const id = ++rowCounter;
  const tr = document.createElement('tr');
  tr.dataset.rowId = id;
  tr.classList.add('row-new');
  tr.setAttribute('role', 'row');

  tr.innerHTML = `
    <td class="col-qty">
      <input
        type="number"
        class="table-input qty-input"
        placeholder="1"
        min="1"
        value="1"
        aria-label="Cantidad del producto ${id}"
        data-row="${id}"
      />
    </td>
    <td class="col-product">
      <input
        type="text"
        class="table-input product-input"
        placeholder="Nombre del producto"
        aria-label="Nombre del producto ${id}"
        data-row="${id}"
      />
    </td>
    <td class="col-unitcost">
      <input
        type="number"
        class="table-input unitcost-input"
        placeholder="0.00"
        min="0"
        step="0.01"
        value="0"
        aria-label="Costo por unidad del producto ${id}"
        data-row="${id}"
      />
    </td>
    <td class="col-totalcost">
      <input
        type="number"
        class="table-input total-input"
        placeholder="0.00"
        min="0"
        step="0.01"
        value="0"
        aria-label="Costo total del producto ${id}"
        data-row="${id}"
      />
    </td>
    <td class="col-action">
      <button
        type="button"
        class="btn-delete-row"
        aria-label="Eliminar fila del producto ${id}"
        data-row="${id}"
        title="Eliminar fila"
      >
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">
          <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V6z"/>
          <path fill-rule="evenodd" d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1v1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3V2h11v1h-11z"/>
        </svg>
      </button>
    </td>
  `;

  // Listeners para cálculo automático
  const qtyInput      = tr.querySelector('.qty-input');
  const unitCostInput = tr.querySelector('.unitcost-input');
  const totalInput    = tr.querySelector('.total-input');

  function recalcRow() {
    const qty  = parseFloat(qtyInput.value)      || 0;
    const unit = parseFloat(unitCostInput.value) || 0;
    const calc = qty * unit;
    totalInput.value = calc.toFixed(2);
    recalcTotal();
  }

  qtyInput.addEventListener('input', recalcRow);
  unitCostInput.addEventListener('input', recalcRow);
  totalInput.addEventListener('input', recalcTotal); // edición manual

  // Botón eliminar
  tr.querySelector('.btn-delete-row').addEventListener('click', () => {
    if (getRowCount() <= 1) {
      alert('Debe haber al menos un producto en la lista.');
      return;
    }
    tr.remove();
    recalcTotal();
    updateProductsError();
  });

  return tr;
}

function getRowCount() {
  return $$('#purchasesBody tr').length;
}

function addRow() {
  if (getRowCount() >= MAX_ROWS) {
    alert('⚠ Límite de elementos alcanzado. Máximo 13 productos por solicitud.');
    return;
  }
  const tbody = document.getElementById('purchasesBody');
  const row = createTableRow();
  tbody.appendChild(row);
  // Enfocar primer input de la fila nueva 
  // C eliminó porq el cursor inicial al entrar a la web se dirige ahi (a la mitad de la web)
  // row.querySelector('.qty-input').focus();
}

function recalcTotal() {
  let sum = 0;
  $$('#purchasesBody .total-input').forEach(inp => {
    sum += parseFloat(inp.value) || 0;
  });
  document.getElementById('totalAmount').textContent = formatBs(sum);
}

function updateProductsError() {
  const errorEl = document.getElementById('products-error');
  const rows = $$('#purchasesBody tr');
  const hasProduct = rows.some(row => {
    const productInput = row.querySelector('.product-input');
    return productInput && productInput.value.trim() !== '';
  });
  if (errorEl) {
    errorEl.textContent = hasProduct ? '' : '';
  }
}

function initTable() {
  // Agregar una fila inicial
  addRow();

  document.getElementById('addRowBtn').addEventListener('click', addRow);
}

/* =============================================
   SERIALIZACIÓN DE DATOS DEL FORMULARIO
   ============================================= */

function serializeForm() {
  const form = document.getElementById('budgetForm');

  // Campos simples
  const data = {
    fechaSolicitud:    form.fechaSolicitud.value.trim(),
    organizacion:      form.organizacion.value,
    nombreLider:       form.nombreLider.value.trim(),
    nombreReceptor:    form.nombreReceptor.value.trim(),
    tipoSolicitud:     (form.tipoSolicitud ? ($$('input[name="tipoSolicitud"]:checked')[0]?.value || '') : ''),
    nombreActividad:   form.nombreActividad.value.trim(),
    fechaActividad:    form.fechaActividad.value.trim(),
    horaActividad:     form.horaActividad.value.trim(),
    lugar:             form.lugar.value.trim(),
    cantidadAsistencia: parseInt(form.cantidadAsistencia.value, 10) || 0,
    dirigidaA:         form.dirigidaA.value,
    proposito:         form.proposito.value.trim(),
  };

  // Metas (checkboxes)
  const metasChecked = $$('input[name="metas"]:checked').map(cb => cb.value);
  data.metas = metasChecked;
  data.metaOtraTexto = form.metaOtraTexto?.value.trim() || '';

  // Productos (tabla dinámica)
  const rows = $$('#purchasesBody tr');
  data.productos = rows.map(row => ({
    cantidad:      parseFloat(row.querySelector('.qty-input')?.value)      || 0,
    producto:      row.querySelector('.product-input')?.value.trim()       || '',
    costoUnitario: parseFloat(row.querySelector('.unitcost-input')?.value) || 0,
    costoTotal:    parseFloat(row.querySelector('.total-input')?.value)    || 0,
  })).filter(p => p.producto !== ''); // Ignorar filas sin nombre

  // Total general
  data.totalSolicitado = data.productos.reduce((acc, p) => acc + p.costoTotal, 0);

  return data;
}

/* =============================================
   VALIDACIÓN COMPLETA ANTES DE ENVIAR
   ============================================= */

function validateAll() {
  const errors = [];

  // Campos de texto obligatorios
  if (!validateRequired('fechaSolicitud', 'Fecha de la Solicitud'))   errors.push('Fecha de la Solicitud');
  if (!validateSelect('organizacion', 'Organización'))                 errors.push('Organización');
  if (!validateRequired('nombreLider', 'Nombre del Líder'))            errors.push('Nombre del Líder');
  if (!validateRequired('nombreReceptor', 'Nombre de quien recibirá')) errors.push('Nombre de quien recibirá el presupuesto');
  if (!validateRequired('nombreActividad', 'Nombre de la Actividad'))  errors.push('Nombre de la Actividad');
  if (!validateRequired('fechaActividad', 'Fecha de la Actividad'))    errors.push('Fecha de la Actividad');
  if (!validateRequired('lugar', 'Lugar'))                             errors.push('Lugar');
  if (!validatePositiveNumber('cantidadAsistencia', 'Asistencia'))     errors.push('Cantidad de Asistencia');
  if (!validateSelect('dirigidaA', 'Actividad Dirigida a'))              errors.push('Actividad Dirigida a');
  if (!validateRequired('proposito', 'Propósito de la Actividad'))     errors.push('Propósito de la Actividad');

  // Tipo de solicitud (radio)
  const tipoChecked = $$('input[name="tipoSolicitud"]:checked');
  if (tipoChecked.length === 0) {
    const errorEl = document.getElementById('tipoSolicitud-error');
    if (errorEl) errorEl.textContent = '⚠ Seleccione el tipo de solicitud (Anticipo o Reembolso).';
    errors.push('Tipo de Solicitud');
  }

  // Metas: al menos una
  const metasChecked = $$('input[name="metas"]:checked');
  if (metasChecked.length === 0) {
    const errorEl = document.getElementById('metas-error');
    if (errorEl) errorEl.textContent = '⚠ Seleccione al menos una meta de la actividad.';
    errors.push('Metas de la actividad');
  } else {
    const errorEl = document.getElementById('metas-error');
    if (errorEl) errorEl.textContent = '';
  }

  // Meta "Otra" con texto vacío
  const otraCheck = document.getElementById('metaOtraCheck');
  const otraInput = document.getElementById('metaOtraTexto');
  if (otraCheck?.checked && !otraInput?.value.trim()) {
    otraInput.classList.add('is-invalid');
    errors.push('Descripción de "Otra" meta');
  }

  // Productos: al menos uno con nombre
  const rows = $$('#purchasesBody tr');
  const validProducts = rows.filter(row => {
    const pi = row.querySelector('.product-input');
    return pi && pi.value.trim() !== '';
  });
  if (validProducts.length === 0) {
    const errorEl = document.getElementById('products-error');
    if (errorEl) errorEl.textContent = '⚠ Agregue al menos un producto con nombre en la lista de compras.';
    errors.push('Lista de compras (mínimo un producto)');
  } else {
    const errorEl = document.getElementById('products-error');
    if (errorEl) errorEl.textContent = '';
  }

  return errors;
}

/* =============================================
   ENVÍO AL BACKEND
   ============================================= */

function setButtonLoading(isLoading) {
  const btn = document.getElementById('submitBtn');
  const content = btn.querySelector('.btn-content');
  const loading = btn.querySelector('.btn-loading');

  btn.disabled = isLoading;

  if (isLoading) {
    content.hidden = true;
    loading.hidden = false;
    loading.removeAttribute('aria-hidden');
  } else {
    content.hidden = false;
    loading.hidden = true;
    loading.setAttribute('aria-hidden', 'true');
  }
}

async function handleSubmit(e) {
  e.preventDefault();

  // Ocultar errores previos
  const errorsBox = document.getElementById('formErrors');
  errorsBox.hidden = true;
  errorsBox.innerHTML = '';

  // Validar todo
  const errors = validateAll();

  if (errors.length > 0) {
    // Mostrar resumen de errores
    errorsBox.innerHTML = `
      <strong>Por favor corrija los siguientes campos antes de continuar:</strong>
      <ul>${errors.map(e => `<li>${e}</li>`).join('')}</ul>
    `;
    errorsBox.hidden = false;

    // Desplazarse al primer error
    const firstInvalid = document.querySelector('.is-invalid, [aria-invalid="true"]');
    if (firstInvalid) {
      firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      firstInvalid.focus();
    }
    return;
  }

  // Serializar datos
  const formData = serializeForm();

  // Iniciar estado de carga
  setButtonLoading(true);

  try {
    const response = await fetch('/.netlify/functions/generar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(formData),
    });

    if (!response.ok) {
      // Intentar leer mensaje de error del servidor
      let serverMsg = 'Error al generar el archivo. Por favor intente nuevamente.';
      try {
        const errData = await response.json();
        if (errData.error) serverMsg = errData.error;
      } catch {}
      throw new Error(serverMsg);
    }

    // Descargar el archivo Excel
    const blob = await response.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;

    // Nombre de archivo con fecha y organización
    const fecha = formData.fechaSolicitud.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ\s]/g, '').replace(/\s+/g, '_');
    const org   = formData.organizacion.replace(/\s+/g, '_');
    a.download  = `Solicitud_${org}_${fecha}.xlsx`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

  } catch (err) {
    console.error('Error al generar Excel:', err);
    errorsBox.innerHTML = `<strong>❌ Error:</strong> ${err.message}`;
    errorsBox.hidden = false;
    errorsBox.scrollIntoView({ behavior: 'smooth', block: 'center' });
  } finally {
    setButtonLoading(false);
  }
}

/* =============================================
   INICIALIZACIÓN
   ============================================= */

document.addEventListener('DOMContentLoaded', () => {
  initTooltips();
  initFieldValidation();
  initOtraCheckbox();
  initTable();

  const form = document.getElementById('budgetForm');
  form.addEventListener('submit', handleSubmit);
});

// 
document.addEventListener('DOMContentLoaded', function() { document.getElementById('fechaSolicitud').focus(); });
