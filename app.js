import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';

const folderRules = new Map([
  ['谷仓-恒禾立:加州区-CA', '谷仓'], ['谷仓-恒禾立:萨凡纳区-GA', '谷仓'], ['谷仓-恒禾立:新泽西区-NJ', '谷仓'], ['谷仓-恒禾立:休斯顿区-HOU', '谷仓'], ['谷仓-恒禾立:芝加哥区-CHI', '谷仓'],
  ['海智链-恒禾立:海智链US-CA中大件A02', '海智链'], ['海智链-恒禾立:海智链US-CA中大件A03', '海智链'], ['海智链-恒禾立:海智链US-CA中大件A05', '海智链'], ['海智链-恒禾立:海智链US-CA中大件B02', '海智链'], ['海智链-恒禾立:海智链US-GA中大件A04', '海智链'], ['海智链-恒禾立:海智链US-NJ中大件A01', '海智链'], ['海智链-恒禾立:海智链US-NJ中大件A04', '海智链'],
  ['乐仓Lecangs-恒禾立:CAP', '乐仓'], ['乐仓Lecangs-恒禾立:CHI', '乐仓'], ['乐仓Lecangs-恒禾立:GAE', '乐仓'], ['乐仓Lecangs-恒禾立:HOU07', '乐仓'], ['乐仓Lecangs-恒禾立:MEM-R', '乐仓'], ['乐仓Lecangs-恒禾立:NJF02', '乐仓'], ['乐仓Lecangs-恒禾立:SCH', '乐仓'],
  ['新链仓-恒禾立:恒禾立-CA008', '链仓'], ['新链仓-恒禾立:恒禾立-GA006', '链仓'], ['新链仓-恒禾立:恒禾立-NJ006', '链仓'],
]);
const normalizedFolders = new Map([...folderRules].map(([name, folder]) => [warehouseKey(name), folder]));
let zipCoordinatesPromise;
let warehouseZipPromise;
let warehouseOutput;
let labelsOutput;

function text(value) { return value === null || value === undefined ? '' : String(value).trim(); }
function norm(value) { return text(value).replace(/[\s_-]/g, '').toUpperCase(); }
function warehouseKey(value) { return text(value).replace(/\s+/g, '').toUpperCase(); }
function normalProduct(value) { return text(value).replace(/[\s_\-/#（）()【】\[\]]+/g, '').toUpperCase(); }
function safeName(value) { return text(value).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/g, '') || 'unnamed'; }
function zipCode(value) { const match = text(value).match(/\d{1,5}/); return match ? match[0].padStart(5, '0') : ''; }
function todayPrefix() { const d = new Date(); return `${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`; }
function fieldIndex(headers, candidates) { const exact = new Map(headers.map((h, i) => [norm(h), i])); for (const item of candidates) if (exact.has(norm(item))) return exact.get(norm(item)); for (const item of candidates) { const needle = norm(item); const i = headers.findIndex(h => norm(h).includes(needle)); if (i >= 0) return i; } return -1; }
function setStatus(id, message, isError = false) { const box = document.querySelector(id); box.textContent = message; box.className = `status${isError ? ' error' : ''}`; box.classList.remove('hidden'); }
function hide(id) { document.querySelector(id).classList.add('hidden'); }
function download(blob, name) { const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1500); }
function counterObject(items) { const map = new Map(); items.forEach(item => map.set(item, (map.get(item) || 0) + 1)); return map; }

document.querySelectorAll('.drop input').forEach(input => input.addEventListener('change', () => { const box = input.closest('.drop'); const name = box.querySelector('.file-name'); const files = [...input.files]; name.textContent = files.length ? (files.length === 1 ? files[0].name : `已选择 ${files.length} 份文件`) : '选择文件'; box.classList.toggle('ready', files.length > 0); }));

async function workbookFromFile(file) { return XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false }); }
function matrix(book, sheetName) { return XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: '', raw: false }); }
function findSheet(book, needed) { for (const name of book.SheetNames) { const rows = matrix(book, name); const headers = rows[0] || []; if (needed.every(item => fieldIndex(headers, [item]) >= 0)) return { name, rows, headers }; } return null; }
function bookBlob(sheets) { const book = XLSX.utils.book_new(); for (const { name, rows } of sheets) XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name.slice(0, 31)); return new Blob([XLSX.write(book, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }
function addColumn(rows, name) { const headers = rows[0]; let index = headers.findIndex(h => text(h) === name); if (index < 0) { index = headers.length; headers.push(name); for (let i = 1; i < rows.length; i++) rows[i][index] = ''; } return index; }
function haversine(a, b) { const r = 3958.7613; const rad = degree => degree * Math.PI / 180; const dLat = rad(b.lat - a.lat), dLon = rad(b.lon - a.lon); const x = Math.sin(dLat / 2) ** 2 + Math.cos(rad(a.lat)) * Math.cos(rad(b.lat)) * Math.sin(dLon / 2) ** 2; return 2 * r * Math.asin(Math.min(1, Math.sqrt(x))); }
function groundZone(origin, destination, coordinates) { const a = coordinates.get(origin), b = coordinates.get(destination); if (!a || !b) return [null, null]; const miles = haversine(a, b); for (const [limit, zone] of [[150, 2], [300, 3], [600, 4], [1000, 5], [1400, 6], [1800, 7]]) if (miles <= limit) return [zone, miles]; return [8, miles]; }
async function warehouseZips() { if (!warehouseZipPromise) warehouseZipPromise = fetch('warehouse_zip_map.csv').then(r => { if (!r.ok) throw new Error('无法加载内置仓库邮编映射。'); return r.text(); }).then(csv => { const rows = csv.trim().split(/\r?\n/).slice(1); return new Map(rows.map(line => { const i = line.lastIndexOf(','); return [warehouseKey(line.slice(0, i)), zipCode(line.slice(i + 1))]; })); }); return warehouseZipPromise; }
async function zipCoordinates() { if (!zipCoordinatesPromise) zipCoordinatesPromise = fetch('us_zip_centroids.csv').then(r => { if (!r.ok) throw new Error('无法加载 ZIP 坐标数据。'); return r.text(); }).then(csv => { const lines = csv.trim().split(/\r?\n/); const headers = lines.shift().split(','); const zi = headers.indexOf('ZipCode'), lati = headers.indexOf('Latitude'), loni = headers.indexOf('Longitude'); const map = new Map(); for (const line of lines) { const parts = line.split(','); const zip = zipCode(parts[zi]), lat = Number(parts[lati]), lon = Number(parts[loni]); if (zip && Number.isFinite(lat) && Number.isFinite(lon)) map.set(zip, { lat, lon }); } return map; }); return zipCoordinatesPromise; }

async function runWarehouse() {
  const inventoryFile = document.querySelector('#inventory-file').files[0]; const exceptionFile = document.querySelector('#warehouse-exception-file').files[0];
  if (!inventoryFile || !exceptionFile) throw new Error('请上传海外仓库存和换仓异常单。');
  setStatus('#warehouse-status', '正在读取库存、异常单及 ZIP 坐标数据，请稍候…'); hide('#warehouse-result');
  const [inventoryBook, exceptionBook, coordinates, warehouseZipsMap] = await Promise.all([workbookFromFile(inventoryFile), workbookFromFile(exceptionFile), zipCoordinates(), warehouseZips()]);
  const inventorySheet = findSheet(inventoryBook, ['仓库', '产品名称', 'SKU', '可用量']); const stateSheet = findSheet(inventoryBook, ['仓库', '州']);
  if (!inventorySheet || !stateSheet) throw new Error('库存文件需包含“仓库、产品名称、SKU、可用量”及“仓库、州”工作表。');
  const ih = inventorySheet.headers, sh = stateSheet.headers;
  const ix = { warehouse: fieldIndex(ih, ['仓库']), product: fieldIndex(ih, ['产品名称']), sku: fieldIndex(ih, ['SKU']), available: fieldIndex(ih, ['可用量']) };
  const sx = { warehouse: fieldIndex(sh, ['仓库']), state: fieldIndex(sh, ['州']) };
  const states = new Map(stateSheet.rows.slice(1).filter(row => text(row[sx.warehouse]) && text(row[sx.state])).map(row => [warehouseKey(row[sx.warehouse]), text(row[sx.state])]));
  const bySku = new Map(), byProduct = new Map();
  for (const row of inventorySheet.rows.slice(1)) { const warehouse = text(row[ix.warehouse]), product = text(row[ix.product]), sku = text(row[ix.sku]), available = Number(row[ix.available]); if (!warehouse || !product || !sku || !(available > 0)) continue; const item = { warehouse, product, sku, available, state: states.get(warehouseKey(warehouse)) || '' }; for (const [map, id] of [[bySku, sku], [byProduct, normalProduct(product)]]) { if (!map.has(id)) map.set(id, []); map.get(id).push(item); } }
  if (!bySku.size) throw new Error('库存文件没有可用量大于 0 的有效 SKU。');
  const orderSheet = findSheet(exceptionBook, ['发货仓库', '指定州', '产品名称', '三方仓SKU']);
  if (!orderSheet) throw new Error('异常订单文件未找到包含发货仓库、指定州、产品名称、三方仓SKU的主表。');
  const rows = orderSheet.rows, h = rows[0]; const c = { warehouse: fieldIndex(h, ['发货仓库']), state: fieldIndex(h, ['指定州']), product: fieldIndex(h, ['产品名称']), sku: fieldIndex(h, ['三方仓SKU']), originZip: fieldIndex(h, ['原指定仓邮编']), recipientZip: fieldIndex(h, ['收货邮编', '收件人邮编']) };
  if (c.originZip < 0 || c.recipientZip < 0) throw new Error('异常单需包含“原指定仓邮编”和“收货邮编”或“收件人邮编”。');
  const audit = { result: addColumn(rows, '换仓匹配结果'), note: addColumn(rows, '换仓说明'), state: addColumn(rows, '匹配仓州'), stock: addColumn(rows, '匹配库存可用量'), zone: addColumn(rows, 'FedEx Zone 校验'), originZone: addColumn(rows, '原指定仓 Zone'), targetZone: addColumn(rows, '候选仓 Zone'), difference: addColumn(rows, 'Zone 差') };
  const summary = new Map(), results = []; let matched = 0, replaced = 0;
  for (let i = 1; i < rows.length; i++) { const row = rows[i]; const sku = text(row[c.sku]), state = text(row[c.state]), product = text(row[c.product]); if (!sku && !state && !product) continue; const [originalZone] = groundZone(zipCode(row[c.originZip]), zipCode(row[c.recipientZip]), coordinates); const candidates = (bySku.get(sku) || byProduct.get(normalProduct(product)) || []).filter(item => item.available >= 1); const sameState = candidates.filter(item => warehouseKey(item.state) === warehouseKey(state)); let selected = null, targetZone = null, difference = null, result, note, zone, targetState = '', stock = '';
    if (sameState.length) { selected = [...sameState].sort((a, b) => b.available - a.available || a.warehouse.localeCompare(b.warehouse))[0]; [targetZone] = groundZone(warehouseZipsMap.get(warehouseKey(selected.warehouse)) || '', zipCode(row[c.recipientZip]), coordinates); difference = targetZone !== null && originalZone !== null ? Math.abs(targetZone - originalZone) : null; result = selected.sku !== sku ? '已换 SKU 并匹配同州仓' : '已匹配同州仓'; note = '指定州内存在可用候选仓'; zone = difference === null ? '候选仓或收货邮编无法计算 Zone' : 'ZIP 距离自动计算（同州）'; }
    else if (candidates.length) { const eligible = candidates.map(item => { const [candidateZone] = groundZone(warehouseZipsMap.get(warehouseKey(item.warehouse)) || '', zipCode(row[c.recipientZip]), coordinates); return { item, candidateZone, difference: candidateZone !== null && originalZone !== null ? Math.abs(candidateZone - originalZone) : null }; }).filter(item => item.difference !== null && item.difference < 2).sort((a, b) => a.difference - b.difference || b.item.available - a.item.available || a.item.warehouse.localeCompare(b.item.warehouse)); if (eligible.length) { ({ item: selected, candidateZone: targetZone, difference } = eligible[0]); result = selected.sku === sku ? '已换仓（Zone 合规）' : '已换仓并替换 SKU（Zone 合规）'; note = `原仓 Zone ${originalZone}，候选仓 Zone ${targetZone}，差值 ${difference} 小于 2`; zone = 'ZIP 距离自动计算，Zone 差小于 2'; } else { const possible = candidates.map(item => { const [candidateZone] = groundZone(warehouseZipsMap.get(warehouseKey(item.warehouse)) || '', zipCode(row[c.recipientZip]), coordinates); return { candidateZone, difference: candidateZone !== null && originalZone !== null ? Math.abs(candidateZone - originalZone) : null }; }).filter(item => item.difference !== null).sort((a, b) => a.difference - b.difference); if (possible.length) ({ candidateZone: targetZone, difference } = possible[0]); result = '需人工复核'; note = '跨州候选仓 Zone 与原指定仓 Zone 的差值大于等于 2，或缺少可计算的仓库/收货邮编'; zone = 'Zone 差大于等于 2，违规或缺少邮编'; } }
    else { result = '无可用库存'; note = '库存表中未找到原三方仓 SKU 或完全一致产品名称的可用库存'; zone = '不适用'; }
    if (selected) { stock = selected.available; selected.available -= 1; row[c.warehouse] = selected.warehouse; row[c.sku] = selected.sku; targetState = selected.state; matched++; if (selected.sku !== sku) replaced++; }
    row[audit.result] = result; row[audit.note] = note; row[audit.state] = targetState; row[audit.stock] = stock; row[audit.zone] = zone; row[audit.originZone] = originalZone ?? ''; row[audit.targetZone] = targetZone ?? ''; row[audit.difference] = difference ?? ''; summary.set(result, (summary.get(result) || 0) + 1); results.push(result);
  }
  warehouseOutput = { blob: bookBlob(exceptionBook.SheetNames.map(name => name === orderSheet.name ? { name, rows } : { name, rows: matrix(exceptionBook, name) })), name: `${todayPrefix()}换仓后的异常单.xlsx` };
  document.querySelector('#warehouse-total').textContent = results.length; document.querySelector('#warehouse-matched').textContent = matched; document.querySelector('#warehouse-replaced').textContent = replaced; document.querySelector('#warehouse-manual').textContent = summary.get('需人工复核') || 0; document.querySelector('#warehouse-result').classList.remove('hidden'); setStatus('#warehouse-status', `处理完成：${results.length} 行订单。`);
}

function extractFields(pageText) { const poSku = pageText.match(/PO#:\s*([A-Z0-9]+)\s+([A-Z0-9][A-Z0-9-]+)/i); const tracking = pageText.match(/(?:TRK#|MPS#)\s*([0-9 ][0-9 ]{8,})/i); if (!poSku || !tracking) return null; return { po: poSku[1].toUpperCase(), sku: poSku[2].toUpperCase(), tracking: tracking[1].replace(/\D/g, '') }; }
function exceptionRowsBook(exceptions) { return bookBlob([{ name: '处理异常', rows: [['类型', '原因', 'PO', 'SKU', 'Tracking', '异常单行号', '来源'], ...exceptions.map(x => [x.type || '', x.reason || '', x.po || '', x.sku || '', x.tracking || '', x.row || '', x.source || ''])] }]); }
function folderFor(warehouse) { return normalizedFolders.get(warehouseKey(warehouse)) || null; }

async function runLabels() {
  const exceptionFile = document.querySelector('#label-exception-file').files[0]; const pdfFiles = [...document.querySelector('#pdf-files').files];
  if (!exceptionFile || !pdfFiles.length) throw new Error('请上传面单匹配异常单和至少一份 PDF。');
  setStatus('#label-status', '正在读取异常单并逐页解析 PDF，请保持页面打开…'); hide('#label-result');
  const book = await workbookFromFile(exceptionFile); const sheetName = book.SheetNames[0]; const rows = matrix(book, sheetName); const headers = rows[0] || [];
  const c = { po: fieldIndex(headers, ['平台单号', '平台PO', '订单号', 'PO']), sku: fieldIndex(headers, ['MSKU', 'Supplier SKU', '供应商SKU', 'SKU']), warehouse: fieldIndex(headers, ['发货仓库', '仓库']) };
  if (Object.values(c).some(x => x < 0)) throw new Error('无法识别表格列：PO、SKU 或发货仓库。');
  const trackingColumn = addColumn(rows, 'Tracking'); const matchColumn = addColumn(rows, '匹配'); const records = [], byKey = new Map(), byTracking = new Map(), occurrence = new Map(), exceptions = [];
  for (let i = 1; i < rows.length; i++) { const row = rows[i]; const po = text(row[c.po]).toUpperCase(), sku = text(row[c.sku]).toUpperCase(); if (!po && !sku) continue; if (!po || !sku) { exceptions.push({ type: '表格字段缺失', reason: '缺少 PO 或 SKU', po, sku, row: i + 1 }); continue; } const base = `${po}_${sku}`, count = (occurrence.get(base) || 0) + 1; occurrence.set(base, count); const record = { rowIndex: i, rowNumber: i + 1, po, sku, base, match: count === 1 ? base : `${base}_${String(count).padStart(2, '0')}`, warehouse: text(row[c.warehouse]), tracking: text(row[trackingColumn]), used: false }; records.push(record); if (!byKey.has(base)) byKey.set(base, []); byKey.get(base).push(record); if (record.tracking) { if (byTracking.has(record.tracking)) exceptions.push({ type: 'Tracking 冲突', reason: '异常单中存在重复 Tracking', tracking: record.tracking, row: i + 1 }); else byTracking.set(record.tracking, record); } }
  if (records.length && !records.some(record => record.warehouse)) throw new Error('面单匹配异常单的“发货仓库”列没有任何仓库数据，无法分类面单。');
  const labels = []; let pages = 0;
  for (const file of pdfFiles) { const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise; for (let p = 1; p <= pdf.numPages; p++) { pages++; const page = await pdf.getPage(p); const content = await page.getTextContent(); const fields = extractFields(content.items.map(item => item.str).join(' ')); if (!fields) exceptions.push({ type: 'Tracking 无法识别', reason: '未能同时识别 PO、SKU 与 Tracking', source: `${file.name} 第 ${p} 页` }); else labels.push({ ...fields, file, page: p, source: `${file.name} 第 ${p} 页` }); } }
  const selected = [], seenTracking = new Set();
  for (const label of labels) { const base = `${label.po}_${label.sku}`; if (seenTracking.has(label.tracking)) { exceptions.push({ type: 'Tracking 冲突', reason: '面单中出现重复 Tracking', ...label }); continue; } seenTracking.add(label.tracking); let record = byTracking.get(label.tracking); if (record && record.base !== base) { exceptions.push({ type: 'Tracking 冲突', reason: 'Tracking 对应的 PO/SKU 与异常单不一致', ...label, row: record.rowNumber }); continue; } if (!record) record = (byKey.get(base) || []).find(item => !item.used && !item.tracking); if (!record) { exceptions.push({ type: '面单未匹配', reason: '异常单中没有可用的 PO/SKU 行', ...label }); continue; } record.used = true; selected.push({ label, record }); }
  for (const record of records) { rows[record.rowIndex][matchColumn] = record.match; if (record.used) rows[record.rowIndex][trackingColumn] = selected.find(item => item.record === record).label.tracking; else exceptions.push({ type: '缺少面单', reason: '异常单未找到对应面单', po: record.po, sku: record.sku, tracking: record.tracking, row: record.rowNumber }); }
  const zip = new JSZip(), pdfSources = new Map(), outputCounts = new Map();
  for (const { label, record } of selected) { const folder = folderFor(record.warehouse); if (!folder) { exceptions.push({ type: '仓库无法识别', reason: record.warehouse || '仓库为空', po: record.po, sku: record.sku, tracking: label.tracking, row: record.rowNumber }); continue; } if (!pdfSources.has(label.file)) pdfSources.set(label.file, await PDFLib.PDFDocument.load(await label.file.arrayBuffer())); const out = await PDFLib.PDFDocument.create(); const [copied] = await out.copyPages(pdfSources.get(label.file), [label.page - 1]); out.addPage(copied); zip.file(`面单整理结果/${folder}/${safeName(record.match)}.pdf`, await out.save()); outputCounts.set(folder, (outputCounts.get(folder) || 0) + 1); }
  const prefix = todayPrefix(); zip.file(`${prefix}异常单_已匹配Tracking.xlsx`, bookBlob(book.SheetNames.map(name => name === sheetName ? { name, rows } : { name, rows: matrix(book, name) }))); zip.file(`${prefix}处理异常清单.xlsx`, exceptionRowsBook(exceptions)); labelsOutput = { blob: await zip.generateAsync({ type: 'blob' }), name: `${prefix}异常单_面单分类结果.zip` };
  document.querySelector('#label-pages').textContent = pages; document.querySelector('#label-recognized').textContent = labels.length; document.querySelector('#label-matched').textContent = selected.length; document.querySelector('#label-exceptions').textContent = exceptions.length; const counts = document.querySelector('#label-counts'); counts.replaceChildren(...[...outputCounts].map(([name, count]) => { const li = document.createElement('li'); li.textContent = `${name}: ${count} 份面单`; return li; })); document.querySelector('#label-result').classList.remove('hidden'); setStatus('#label-status', `处理完成：分类 ZIP 已生成，共 ${[...outputCounts.values()].reduce((a, b) => a + b, 0)} 张正式面单。`);
}

function wire(buttonId, statusId, handler) { document.querySelector(buttonId).addEventListener('click', async () => { const button = document.querySelector(buttonId); button.disabled = true; const old = button.textContent; button.textContent = '处理中…'; try { await handler(); } catch (error) { console.error(error); setStatus(statusId, error.message || '处理失败，请检查文件格式。', true); } finally { button.disabled = false; button.textContent = old; } }); }
wire('#warehouse-run', '#warehouse-status', runWarehouse); wire('#label-run', '#label-status', runLabels);
document.querySelector('#warehouse-download').addEventListener('click', () => warehouseOutput && download(warehouseOutput.blob, warehouseOutput.name));
document.querySelector('#labels-download').addEventListener('click', () => labelsOutput && download(labelsOutput.blob, labelsOutput.name));
