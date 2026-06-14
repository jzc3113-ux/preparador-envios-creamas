const requiredBlockColumns = [
  'id_envio', 'id_campaña', 'nombre_campaña', 'bloque', 'correo', 'nombre', 'asunto', 'mensaje_html',
  'estado_envio', 'fecha_envio', 'error_envio', 'respondio', 'fecha_respuesta', 'tipo_respuesta',
  'estado_seguimiento', 'observacion'
];

const optionalFields = ['celular', 'dni', 'ciudad', 'institución', 'variable_1', 'variable_2', 'variable_3'];
const mappingFields = ['correo', 'nombre', ...optionalFields];
const state = { workbook: null, rows: [], headers: [], validation: null };

const els = {
  campaignName: document.querySelector('#campaignName'), excelFile: document.querySelector('#excelFile'),
  sheetChooser: document.querySelector('#sheetChooser'), sheetName: document.querySelector('#sheetName'),
  mappingGrid: document.querySelector('#mappingGrid'), subject: document.querySelector('#subject'),
  messageHtml: document.querySelector('#messageHtml'), blockSize: document.querySelector('#blockSize'),
  validateBtn: document.querySelector('#validateBtn'), downloadBtn: document.querySelector('#downloadBtn'),
  summaryGrid: document.querySelector('#summaryGrid'), statusMessage: document.querySelector('#statusMessage')
};

function normalizeHeader(value) { return String(value || '').trim(); }
function normalizeEmail(value) { return String(value || '').trim().toLowerCase(); }
function isValidEmail(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email); }
function safeFileName(name) { return normalizeHeader(name).replace(/[^a-z0-9áéíóúñü_-]+/gi, '_').replace(/^_+|_+$/g, '') || 'campaña'; }
function padBlock(number) { return String(number).padStart(3, '0'); }

function setStatus(message, isError = false) {
  els.statusMessage.textContent = message;
  els.statusMessage.style.background = isError ? '#fef2f2' : '#f0fdfa';
  els.statusMessage.style.color = isError ? '#991b1b' : '#115e59';
}

function readSelectedSheet() {
  const sheet = state.workbook.Sheets[els.sheetName.value];
  state.rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  state.headers = state.rows.length ? Object.keys(state.rows[0]).map(normalizeHeader) : [];
  renderMapping();
  els.downloadBtn.disabled = true;
  state.validation = null;
  updateSummary({ totalRows: state.rows.length, validCount: 0, invalidEmailCount: 0, duplicateCount: 0, emptyEmailCount: 0, blockCount: 0 });
  setStatus(`Hoja cargada: ${state.rows.length} filas detectadas. Revisa el mapeo y valida la base.`);
}

function renderMapping() {
  els.mappingGrid.innerHTML = '';
  mappingFields.forEach((field) => {
    const wrapper = document.createElement('div');
    const label = document.createElement('label');
    const select = document.createElement('select');
    label.textContent = field === 'nombre' || field === 'correo' ? `${field} *` : field;
    label.setAttribute('for', `map_${field}`);
    select.id = `map_${field}`;
    select.dataset.field = field;
    select.innerHTML = '<option value="">No usar</option>' + state.headers.map((header) => `<option value="${header}">${header}</option>`).join('');
    const guessed = state.headers.find((header) => header.toLowerCase().replaceAll(' ', '_') === field) || '';
    select.value = guessed;
    wrapper.append(label, select);
    els.mappingGrid.appendChild(wrapper);
  });
}

async function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const buffer = await file.arrayBuffer();
  state.workbook = XLSX.read(buffer, { type: 'array' });
  els.sheetName.innerHTML = state.workbook.SheetNames.map((name) => `<option value="${name}">${name}</option>`).join('');
  els.sheetChooser.classList.remove('hidden');
  readSelectedSheet();
}

function getMapping() {
  return Object.fromEntries(mappingFields.map((field) => [field, document.querySelector(`#map_${field}`)?.value || '']));
}

function validateRows() {
  const campaignName = normalizeHeader(els.campaignName.value);
  const subject = normalizeHeader(els.subject.value);
  const messageHtml = els.messageHtml.value.trim();
  const blockSize = Number(els.blockSize.value) || 250;
  const mapping = getMapping();
  if (!campaignName || !subject || !messageHtml) throw new Error('Completa nombre de campaña, asunto y mensaje HTML.');
  if (!mapping.correo || !mapping.nombre) throw new Error('Mapea las columnas obligatorias: correo y nombre.');
  if (!state.rows.length) throw new Error('Sube un Excel con al menos una fila.');

  const seen = new Set();
  const valid = [];
  const errors = [];
  let emptyEmailCount = 0;
  let invalidEmailCount = 0;
  let duplicateCount = 0;

  state.rows.forEach((row, index) => {
    const originalRow = index + 2;
    const email = normalizeEmail(row[mapping.correo]);
    const name = normalizeHeader(row[mapping.nombre]);
    const reasons = [];
    if (!email) { reasons.push('correo vacío'); emptyEmailCount += 1; }
    else if (!isValidEmail(email)) { reasons.push('correo inválido'); invalidEmailCount += 1; }
    if (email && seen.has(email)) { reasons.push('duplicado'); duplicateCount += 1; }
    if (!name) reasons.push('fila incompleta');

    const extra = Object.fromEntries(optionalFields.map((field) => [field, mapping[field] ? row[mapping[field]] || '' : '']));
    const base = { fila_origen: originalRow, correo: email, nombre: name, ...extra };
    if (reasons.length) {
      errors.push({ ...base, observacion: reasons.join(', ') });
    } else {
      seen.add(email);
      valid.push(base);
    }
  });

  const blockCount = Math.ceil(valid.length / blockSize);
  state.validation = { campaignName, campaignId: `CAMP-${Date.now()}`, subject, messageHtml, blockSize, valid, errors, summary: { totalRows: state.rows.length, validCount: valid.length, invalidEmailCount, duplicateCount, emptyEmailCount, blockCount } };
  updateSummary(state.validation.summary);
  els.downloadBtn.disabled = false;
  setStatus(`Validación lista: ${valid.length} registros válidos y ${errors.length} registros con observación.`);
}

function updateSummary(summary) {
  const values = [summary.totalRows, summary.validCount, summary.invalidEmailCount, summary.duplicateCount, summary.emptyEmailCount, summary.blockCount];
  els.summaryGrid.querySelectorAll('strong').forEach((node, index) => { node.textContent = values[index] || 0; });
}

async function excelTableBlob(sheetName, tableName, columns, rows, options = {}) {
  if (!window.ExcelJS) {
    throw new Error('No se pudo cargar ExcelJS. Revisa tu conexión a internet e intenta nuevamente.');
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Preparador de Envíos Crea+';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  const table = worksheet.addTable({
    name: tableName,
    displayName: tableName,
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: options.theme || 'TableStyleMedium2',
      showRowStripes: true
    },
    columns: columns.map((column) => ({ name: column, filterButton: true })),
    rows
  });
  table.commit();

  worksheet.columns.forEach((column, index) => {
    const header = columns[index];
    const widthByHeader = {
      mensaje_html: 48,
      observacion: 34,
      asunto: 32,
      nombre_campaña: 28,
      id_envio: 32,
      correo: 30,
      nombre: 28
    };
    column.width = widthByHeader[header] || 20;
  });

  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: options.headerColor || 'FFFF6B35' } };
    cell.alignment = { vertical: 'middle', wrapText: true };
  });

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'top', wrapText: cell.col === columns.indexOf('mensaje_html') + 1 };
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function buildBlockRows(records, blockNumber) {
  return records.map((record, index) => [
    `${state.validation.campaignId}-${padBlock(blockNumber)}-${String(index + 1).padStart(4, '0')}`,
    state.validation.campaignId, state.validation.campaignName, blockNumber, record.correo, record.nombre,
    state.validation.subject, state.validation.messageHtml, 'PENDIENTE', '', '', 'NO', '', '', 'SIN ENVIAR', ''
  ]);
}

async function generateZip() {
  if (!state.validation) validateRows();
  const zip = new JSZip();
  const { valid, errors, blockSize, summary, campaignName } = state.validation;
  for (let start = 0, block = 1; start < valid.length; start += blockSize, block += 1) {
    const rows = buildBlockRows(valid.slice(start, start + blockSize), block);
    zip.file(`bloque_${padBlock(block)}.xlsx`, await excelTableBlob('Envios', 'TablaEnvios', requiredBlockColumns, rows));
  }
  const errorColumns = ['fila_origen', 'correo', 'nombre', ...optionalFields, 'observacion'];
  const errorRows = errors.map((row) => errorColumns.map((key) => row[key] || ''));
  zip.file('reporte_errores.xlsx', await excelTableBlob('Errores', 'TablaErrores', errorColumns, errorRows, { theme: 'TableStyleMedium3', headerColor: 'FFB42318' }));
  const resumenColumns = ['campo', 'valor'];
  const resumenRows = [['nombre_campaña', campaignName], ['fecha_generacion', new Date().toISOString()], ['total_filas', summary.totalRows], ['validos', summary.validCount], ['errores', errors.length], ['duplicados', summary.duplicateCount], ['bloques_generados', summary.blockCount], ['tamaño_bloque', blockSize]];
  zip.file('resumen_campaña.xlsx', await excelTableBlob('Resumen', 'TablaResumen', resumenColumns, resumenRows, { theme: 'TableStyleMedium4', headerColor: 'FF0F766E' }));
  zip.file('plantilla_correo.html', state.validation.messageHtml);
  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, `${safeFileName(campaignName)}_power_automate.zip`);
  setStatus('ZIP generado correctamente con bloques, reportes y plantilla HTML.');
}

els.excelFile.addEventListener('change', handleFile);
els.sheetName.addEventListener('change', readSelectedSheet);
els.validateBtn.addEventListener('click', () => { try { validateRows(); } catch (error) { setStatus(error.message, true); } });
els.downloadBtn.addEventListener('click', () => { generateZip().catch((error) => setStatus(error.message, true)); });
renderMapping();
