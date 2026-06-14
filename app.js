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

function aoaToWorkbook(sheetName, rows, tableName) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const range = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length - 1, 0), c: rows[0].length - 1 } });
  worksheet['!autofilter'] = { ref: range };
  worksheet['!cols'] = rows[0].map(() => ({ wch: 22 }));
  if (tableName) worksheet['!tables'] = [{ name: tableName, ref: range, headerRow: true, totalsRow: false }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  return workbook;
}

function workbookToBlob(workbook) {
  const array = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

async function workbookToTableBlob(workbook, tableName, rowCount, columnCount) {
  const blob = workbookToBlob(workbook);
  const zip = await JSZip.loadAsync(blob);
  const tableRef = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rowCount - 1, 0), c: columnCount - 1 } });
  const headers = requiredBlockColumns.map((name, index) => `<tableColumn id="${index + 1}" name="${name}"/>`).join('');
  zip.file('xl/tables/table1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${tableName}" displayName="${tableName}" ref="${tableRef}" totalsRowShown="0"><autoFilter ref="${tableRef}"/><tableColumns count="${columnCount}">${headers}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);

  const contentTypes = await zip.file('[Content_Types].xml').async('string');
  if (!contentTypes.includes('/xl/tables/table1.xml')) {
    zip.file('[Content_Types].xml', contentTypes.replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
  }

  const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const relXml = zip.file(relPath)
    ? await zip.file(relPath).async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const tableRelId = relXml.includes('Id="rId1"') ? 'rIdTable1' : 'rId1';
  if (!relXml.includes('../tables/table1.xml')) {
    zip.file(relPath, relXml.replace('</Relationships>', `<Relationship Id="${tableRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`));
  }

  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  if (!sheetXml.includes('<tableParts')) {
    zip.file('xl/worksheets/sheet1.xml', sheetXml.replace('</worksheet>', `<tableParts count="1"><tablePart r:id="${tableRelId}"/></tableParts></worksheet>`));
  }
  return await zip.generateAsync({ type: 'blob' });
}

function buildBlockRows(records, blockNumber) {
  return [requiredBlockColumns, ...records.map((record, index) => [
    `${state.validation.campaignId}-${padBlock(blockNumber)}-${String(index + 1).padStart(4, '0')}`,
    state.validation.campaignId, state.validation.campaignName, blockNumber, record.correo, record.nombre,
    state.validation.subject, state.validation.messageHtml, 'PENDIENTE', '', '', 'NO', '', '', 'SIN ENVIAR', ''
  ])];
}

async function generateZip() {
  if (!state.validation) validateRows();
  const zip = new JSZip();
  const { valid, errors, blockSize, summary, campaignName } = state.validation;
  for (let start = 0, block = 1; start < valid.length; start += blockSize, block += 1) {
    const rows = buildBlockRows(valid.slice(start, start + blockSize), block);
    zip.file(`bloque_${padBlock(block)}.xlsx`, await workbookToTableBlob(aoaToWorkbook('Envios', rows), 'TablaEnvios', rows.length, requiredBlockColumns.length));
  }
  const errorRows = [['fila_origen', 'correo', 'nombre', ...optionalFields, 'observacion'], ...errors.map((r) => ['fila_origen', 'correo', 'nombre', ...optionalFields, 'observacion'].map((key) => r[key] || ''))];
  zip.file('reporte_errores.xlsx', workbookToBlob(aoaToWorkbook('Errores', errorRows)));
  const resumenRows = [['campo', 'valor'], ['nombre_campaña', campaignName], ['fecha_generacion', new Date().toISOString()], ['total_filas', summary.totalRows], ['validos', summary.validCount], ['errores', errors.length], ['duplicados', summary.duplicateCount], ['bloques_generados', summary.blockCount], ['tamaño_bloque', blockSize]];
  zip.file('resumen_campaña.xlsx', workbookToBlob(aoaToWorkbook('Resumen', resumenRows)));
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
