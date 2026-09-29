import * as XLSX from 'xlsx';
import './app.css';

const $ = (selector) => document.querySelector(selector);
let schools = [];
const DRIVE_FOLDER_ID = '1q4PHnubRVIsJ3LwBCbMpJNFYj-ztcQ96';
const GOOGLE_DRIVE_API_KEY = import.meta.env.VITE_GOOGLE_DRIVE_API_KEY || '';
let driveFiles = [];
let driveSourceLoaded = false;
let forceDriveRefresh = false;
let manualInventoryFiles = [];
const selectedSchools = new Set();

const FOLDER_DB_NAME = 'auditoria-escolar-config';
const FOLDER_STORE_NAME = 'settings';

function openFolderDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(FOLDER_DB_NAME, 2);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(FOLDER_STORE_NAME)) request.result.createObjectStore(FOLDER_STORE_NAME);
      if (!request.result.objectStoreNames.contains('drive-files')) request.result.createObjectStore('drive-files');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getCachedDriveFile(file) {
  const db = await openFolderDb();
  const record = await new Promise((resolve, reject) => { const tx=db.transaction('drive-files','readonly'); const request=tx.objectStore('drive-files').get(file.id); request.onsuccess=()=>resolve(request.result||null); request.onerror=()=>reject(request.error); });
  if (!record || record.modifiedTime !== file.modifiedTime) return null;
  return new File([record.blob], record.name, { type: record.mimeType || file.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', lastModified: record.lastModified || Date.now() });
}

async function cacheDriveFile(file, blob) {
  const db = await openFolderDb();
  await new Promise((resolve, reject) => { const tx=db.transaction('drive-files','readwrite'); tx.objectStore('drive-files').put({ id:file.id, name:file.name, mimeType:file.mimeType, modifiedTime:file.modifiedTime, lastModified:file.modifiedTime ? Date.parse(file.modifiedTime) : Date.now(), blob }, file.id); tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error); });
}

async function saveFolderHandle(handle) {
  const db = await openFolderDb();
  await new Promise((resolve, reject) => { const tx=db.transaction(FOLDER_STORE_NAME,'readwrite'); tx.objectStore(FOLDER_STORE_NAME).put(handle,'inventory-folder'); tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error); });
}

async function getFolderHandle() {
  const db = await openFolderDb();
  return new Promise((resolve, reject) => { const tx=db.transaction(FOLDER_STORE_NAME,'readonly'); const request=tx.objectStore(FOLDER_STORE_NAME).get('inventory-folder'); request.onsuccess=()=>resolve(request.result||null); request.onerror=()=>reject(request.error); });
}

async function filesFromFolder(handle) {
  const files = [];
  for await (const entry of handle.values()) {
    if (entry.kind === 'file' && entry.name.toLowerCase().endsWith('.xlsx') && !entry.name.startsWith('~$')) files.push(await entry.getFile());
  }
  return files;
}

async function useManualFolder(handle) {
  if (handle.queryPermission && await handle.queryPermission({ mode: 'read' }) !== 'granted') await handle.requestPermission({ mode: 'read' });
  manualInventoryFiles = await filesFromFolder(handle);
  driveSourceLoaded = false;
  renderDriveMatrixOptions();
  $('#inventoryFileLabel').textContent = manualInventoryFiles.length ? `${manualInventoryFiles.length} arquivo(s) .xlsx selecionado(s) da pasta manual.` : 'Nenhum arquivo .xlsx encontrado nessa pasta.';
}

const labels = {
  OK: 'Sem divergência',
  DIVERGENCIA: 'Com divergência',
  NAO_RECEBIDO: 'Não recebido',
  DUPLICADO: 'Duplicado',
  ERRO: 'Erro de leitura',
  FORA_DA_MATRIZ: 'Fora da matriz'
};

const rankLabels = {
  all: 'Todos os status',
  equipment_ok: 'Equipamentos OK',
  maintenance: 'Garantia / manutenção',
  damaged: 'Danificados / inservíveis'
};

const SERIAL_FIELDS = ['NUMERO DE SERIE', 'NÚMERO DE SÉRIE', 'SERIAL', 'Nº DE SERIE', 'Nº DE SÉRIE'];
const NO_SERIAL = new Set(['', 'SEM SERIAL', 'N/A', 'NA', '-', 'NAO INFORMADO', 'NÃO INFORMADO'].map(normalize));

function normalize(value) {
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function getValue(item, fields) {
  const normalized = new Map(Object.entries(item).map(([key, value]) => [normalize(key), value]));
  for (const field of fields) {
    const value = normalized.get(normalize(field));
    if (value !== undefined && value !== null && value !== '') return normalize(value);
  }
  return '';
}

function cieText(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return String(Math.trunc(value));
  return String(value ?? '').trim().replace(/\.0$/, '');
}

function expectedInfo(header) {
  const text = normalize(header);
  const match = text.match(/\(([^)]+)\)/);
  const manufacturer = match ? normalize(match[1]) : text.includes('TES') ? 'TES' : text.includes('MOVPLAN') ? 'MOVPLAN' : '';
  if (['PLATAFORMA', 'REC.', 'RECARGA', 'GABINETE', 'CARRINHO'].some(token => text.includes(token))) return ['RECARGA', manufacturer];
  if (text.includes('DESKTOP')) return ['DESKTOP', manufacturer];
  if (['NOTEBOOK BAS', 'NOTEBOOK SALA', 'NOTEBOOK PLUS'].some(token => text.includes(token))) return ['NOTEBOOK_BASICO', manufacturer];
  if (text.includes('NOTEBOOK AVAN')) return ['NOTEBOOK_AVANCADO', manufacturer];
  if (text.includes('CHROMEBOOK')) return ['CHROMEBOOK', manufacturer];
  if (text.includes('TV') || text.includes('EDUCATRON')) return ['TV', manufacturer];
  if (text.includes('TABLET')) return ['TABLET', manufacturer];
  if (text.includes('SMARTPHONE')) return ['SMARTPHONE', manufacturer];
  return ['', manufacturer];
}

async function workbookFromFile(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: 'array', cellDates: true });
}

async function listDriveFiles() {
  if (!GOOGLE_DRIVE_API_KEY) throw new Error('Google Drive não configurado nesta publicação. Na Vercel, adicione a variável VITE_GOOGLE_DRIVE_API_KEY e faça um novo deploy.');
  const fields = 'nextPageToken,files(id,name,mimeType,size,modifiedTime)';
  const files = [];
  let pageToken = '';
  do {
    const params = new URLSearchParams({
      q: `'${DRIVE_FOLDER_ID}' in parents and trashed = false`,
      fields, pageSize: '1000', orderBy: 'name', key: GOOGLE_DRIVE_API_KEY
    });
    if (pageToken) params.set('pageToken', pageToken);
    const response = await fetch(`https://www.googleapis.com/drive/v3/files?${params}`);
    if (!response.ok) throw new Error(response.status === 403 ? 'Google Drive recusou o acesso (403). Confira se a variável VITE_GOOGLE_DRIVE_API_KEY foi configurada na Vercel, se a Drive API está ativa e se o domínio está liberado na chave.' : `Não foi possível acessar o Google Drive (${response.status}). Verifique a chave e o compartilhamento da pasta.`);
    const data = await response.json();
    files.push(...(data.files || []));
    pageToken = data.nextPageToken || '';
  } while (pageToken);
  return files.filter(file => file.name.toLowerCase().endsWith('.xlsx'));
}

async function downloadDriveFile(file, { forceRefresh = false } = {}) {
  if (!forceRefresh) {
    const cached = await getCachedDriveFile(file);
    if (cached) return cached;
  }
  const params = new URLSearchParams({ alt: 'media', key: GOOGLE_DRIVE_API_KEY });
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 45000);
  let response;
  try { response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?${params}`, { signal: controller.signal }); }
  catch (error) { throw new Error(`Tempo excedido ao baixar ${file.name} do Google Drive.`); }
  finally { clearTimeout(timeout); }
  if (!response.ok) throw new Error(`Não foi possível baixar ${file.name} do Google Drive (${response.status}).`);
  const blob = await response.blob();
  await cacheDriveFile(file, blob);
  return new File([blob], file.name, { type: file.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', lastModified: file.modifiedTime ? Date.parse(file.modifiedTime) : Date.now() });
}

async function downloadDriveFiles(files, { forceRefresh = false } = {}) {
  const downloaded = [];
  const batchSize = 6;
  for (let index = 0; index < files.length; index += batchSize) {
    const batch = files.slice(index, index + batchSize);
    downloaded.push(...await Promise.all(batch.map(file => downloadDriveFile(file, { forceRefresh }))));
    $('#scanButton').textContent = `Baixando ${Math.min(index + batch.length, files.length)}/${files.length}…`;
  }
  return downloaded;
}

function renderDriveMatrixOptions() {
  const select = $('#driveMatrixSelect');
  const preferred = driveFiles.find(file => normalize(file.name).includes('COMPRAS CENTRALIZADAS'));
  select.innerHTML = '<option value="">Escolha a matriz entre os arquivos do Drive</option>' + driveFiles.map(file => `<option value="${escapeHtml(file.id)}" ${file.id === preferred?.id ? 'selected' : ''}>${escapeHtml(file.name)}</option>`).join('');
  select.classList.toggle('hidden', !driveSourceLoaded);
  if (preferred) $('#matrixFileLabel').textContent = `Matriz selecionada automaticamente: ${preferred.name}`;
}

async function readMatrix(file) {
  const wb = await workbookFromFile(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  let headerIndex = -1, cieIndex = -1, schoolIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    const normalized = rows[i].map(normalize);
    if (normalized.includes('CIE') && normalized.includes('NOME_ESCOLA')) {
      headerIndex = i;
      cieIndex = normalized.indexOf('CIE');
      schoolIndex = normalized.indexOf('NOME_ESCOLA');
      break;
    }
  }
  if (headerIndex < 0) throw new Error('A matriz não contém as colunas CIE e NOME_ESCOLA.');
  const headers = rows[headerIndex].map(value => String(value ?? '').trim());
  const result = new Map();
  for (const row of rows.slice(headerIndex + 1)) {
    const cie = cieText(row[cieIndex]);
    if (!/^\d+$/.test(cie)) continue;
    const schoolName = String(row[schoolIndex] ?? '').trim();
    const expected = new Map();
    headers.forEach((header, index) => {
      const normHeader = normalize(header);
      if (index === cieIndex || index === schoolIndex || normHeader === 'URE' || normHeader.includes('POS SED')) return;
      const raw = row[index] ?? 0;
      let quantity = 0;
      if (typeof raw === 'string') quantity = parseInt(raw.replaceAll('.', '').replace(',', '.'), 10) || 0;
      else quantity = parseInt(Number(raw) || 0, 10) || 0;
      if (quantity <= 0) return;
      let [category, manufacturer] = expectedInfo(header);
      if (!category) return;
      if (normHeader.includes('NOTEBOOK SALA DE AULA 2025') && manufacturer === 'POSITIVO' && normalize(schoolName).includes('SARAH SALVESTRO') && quantity === 15) {
        category = 'CHROMEBOOK'; manufacturer = 'SAMSUNG';
      }
      const key = `${category}|${manufacturer}`;
      const entry = expected.get(key) || { category, manufacturer, quantity: 0, names: [] };
      entry.quantity += quantity;
      entry.names.push(header);
      expected.set(key, entry);
    });
    result.set(cie, { name: schoolName, expected });
  }
  return result;
}

function classify(item) {
  const sheet = normalize(item._sheet || '');
  if (sheet.includes('PLATAFORMA') || sheet.includes('RECARGA')) return 'RECARGA';
  const manufacturer = getValue(item, ['FABRICANTE', 'MARCA', 'MARCA / FABRICANTE']);
  const model = getValue(item, ['MODELO', 'TIPO / MODELO', 'MODELO EQUIPAMENTO', 'DESCRICAO', 'DESCRIÇÃO', 'OUTROS']);
  const category = getValue(item, ['CATEGORIA DO EQL', 'CATEGORIA DO EQUIPAMENTO', 'CATEGORIA', 'TIPO', 'TIPO DE EQUIPAMENTO']);
  if (model.includes('CHROME') || category.includes('CHROMEBOOK')) return 'CHROMEBOOK';
  if (category.includes('SMARTPHONE') || model.includes('SMARTPHONE') || model.includes('PHONE') || ['MOTOROLA', 'XIAOMI'].some(x => manufacturer.includes(x))) return 'SMARTPHONE';
  if (category.includes('TABLET') || model.includes('TABLET') || ` ${model}`.includes(' TAB')) return 'TABLET';
  if (['TV', 'EDUCATRON'].some(x => category.includes(x) || model.includes(x)) || ['LG', 'PHILCO', 'AOC'].some(x => manufacturer.includes(x))) return 'TV';
  if (category.includes('DESKTOP') || ['DESKTOP', 'PC', 'ALL IN ONE'].some(x => model.includes(x))) return 'DESKTOP';
  if (manufacturer.includes('POSITIVO')) {
    if (['N1110', 'N1210', 'N2110', 'MOTION'].some(x => model.includes(x))) return 'NOTEBOOK_BASICO';
    if (['POS-PIQ', 'C1400', 'C8256', 'PRESLEY'].some(x => model.includes(x))) return 'DESKTOP';
    if (['MASTER', 'EXPERT', 'VISION', 'N8440'].some(x => model.includes(x))) return 'NOTEBOOK_AVANCADO';
    return 'NOTEBOOK_BASICO';
  }
  if (manufacturer.includes('LENOVO') || model.includes('LENOVO')) {
    if (['11AA', '11JA', '11BL', '10RR', '10T7'].some(x => model.includes(x)) || category.includes('DESKTOP')) return 'DESKTOP';
    return 'NOTEBOOK_BASICO';
  }
  if (['ULTRA', 'MULTI', 'MULTILASER'].some(x => manufacturer.includes(x))) return 'NOTEBOOK_BASICO';
  if (manufacturer.includes('SAMSUNG')) return /SM-|MOTO|GALAXY/.test(model) ? 'SMARTPHONE' : 'NOTEBOOK_BASICO';
  if (['DELL', 'HP', 'ACER', 'ASUS'].some(x => manufacturer.includes(x))) return 'NOTEBOOK_BASICO';
  if (manufacturer.includes('TES')) return 'RECARGA';
  if (category.includes('NOTEBOOK') || model.includes('NOTEBOOK')) return 'NOTEBOOK_BASICO';
  return '';
}

function aliases(manufacturer) {
  const mapping = {
    MULTI: ['MULTILASER', 'ULTRA'], POSITIVO: ['POSITIVO TECNOLOGIA SA', 'POSITIVO INFORMATICA'],
    SAMSUNG: ['SAMSUNG ELECTRONICS'], DELL: ['DELL INC.'], HP: ['HEWLETT-PACKARD'], LENOVO: ['LENOVO'], TES: ['TES'], MOVPLAN: ['MOVPLAN']
  };
  return [manufacturer, ...(mapping[manufacturer] || [])];
}

function equipmentCondition(item) {
  const status = getValue(item, ['STATUS DO EQUIPAMENTO']);
  if (['DANIFICADO', 'INSERVIVEL', 'RUIM'].some(token => status.includes(token))) return 'damaged';
  if (['MANUTENCAO', 'GARANTIA', 'CHAMADO ABERTO', 'AGUARDANDO PECA'].some(token => status.includes(token))) return 'maintenance';
  if (['DISPONIVEL', 'OK', 'BOM', 'EQUIPAMENTO FUNCIONANDO'].includes(status)) return 'equipment_ok';
  return 'status_unclassified';
}

function readInventoryWorkbook(wb) {
  const records = [];
  const seen = new Set();
  let noSerialCounter = 0;
  for (const sheetName of wb.SheetNames) {
    if (normalize(sheetName).includes('LISTAS')) continue;
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
    for (const raw of rows) {
      const item = { ...raw, _sheet: sheetName };
      const serial = getValue(item, SERIAL_FIELDS);
      const uniqueKey = serial && !NO_SERIAL.has(serial) ? `SERIAL|${serial}` : `NO_SERIAL|${noSerialCounter++}`;
      if (!seen.has(uniqueKey)) { seen.add(uniqueKey); records.push(item); }
    }
  }
  return records;
}

function auditWorkbook(wb, matrixEntry) {
  const expected = new Map([...matrixEntry.expected.entries()].map(([key, value]) => [key, { ...value, found: 0, equipment_ok: 0, maintenance: 0, damaged: 0, status_unclassified: 0 }]));
  const counts = { equipment_ok: 0, maintenance: 0, damaged: 0, status_unclassified: 0 };
  for (const item of readInventoryWorkbook(wb)) {
    const condition = equipmentCondition(item);
    counts[condition]++;
    const category = classify(item);
    if (!category) continue;
    const manufacturer = getValue(item, ['FABRICANTE', 'MARCA', 'MARCA / FABRICANTE']);
    const model = getValue(item, ['MODELO', 'TIPO / MODELO', 'MODELO EQUIPAMENTO', 'DESCRICAO', 'DESCRIÇÃO', 'OUTROS']);
    const area = `${manufacturer} ${model}`;
    let matched = false;
    for (const entry of expected.values()) {
      if (category !== entry.category) continue;
      const wanted = entry.manufacturer;
      if (!wanted || aliases(wanted).some(alias => area.includes(alias))) {
        entry.found++;
        entry[condition]++;
        matched = true;
        break;
      }
    }
    if (!matched && category === 'RECARGA') {
      const fallback = [...expected.values()].find(entry => entry.category === 'RECARGA');
      if (fallback) { fallback.found++; fallback[condition]++; }
    }
  }
  const details = [];
  let expectedTotal = 0, foundTotal = 0;
  for (const entry of expected.values()) {
    const difference = entry.found - entry.quantity;
    details.push({ ...entry, difference, status: difference < 0 ? 'FALTANDO' : difference > 0 ? 'SOBRANDO' : 'OK' });
    expectedTotal += entry.quantity; foundTotal += entry.found;
  }
  return { details, expectedTotal, foundTotal, counts };
}

async function executeAudit(matrixFile, inventoryFiles) {
  const matrix = await readMatrix(matrixFile);
  const filesByCie = new Map();
  const invalidFiles = [];
  for (const file of inventoryFiles) {
    const match = file.name.match(/^\s*(\d{4,9})\s*[-–—]/);
    if (!match) { invalidFiles.push(file.name); continue; }
    const list = filesByCie.get(match[1]) || [];
    list.push(file); filesByCie.set(match[1], list);
  }
  const cies = [...new Set([...matrix.keys(), ...filesByCie.keys()])].sort((a,b) => a.length - b.length || a.localeCompare(b));
  const rows = [];
  let processed = 0;
  for (const cie of cies) {
    const matrixEntry = matrix.get(cie);
    const files = filesByCie.get(cie) || [];
    const schoolName = matrixEntry?.name || (files[0] ? files[0].name.replace(/^\s*\d+\s*[-–—]\s*|\s*-\s*INVENT.RIO.*$/gi, '').replace(/\.xlsx$/i,'') : '');
    const base = { cie, school_name: schoolName, file_name: files[0]?.name || '', file_mtime: '', expected: 0, found: 0, missing: 0, extra: 0, equipment_ok: 0, maintenance: 0, damaged: 0, status_unclassified: 0, details: [], error: '' };
    if (!matrixEntry) base.status = 'FORA_DA_MATRIZ', base.error = 'CIE não encontrado na matriz';
    else if (!files.length) {
      const total = [...matrixEntry.expected.values()].reduce((sum, x) => sum + x.quantity, 0);
      Object.assign(base, { status: 'NAO_RECEBIDO', expected: total, missing: total });
    } else if (files.length > 1) {
      Object.assign(base, { status: 'DUPLICADO', error: `${files.length} arquivos encontrados para o mesmo CIE`, file_name: files.map(f => f.name).join(' | ') });
    } else {
      try {
        const wb = await workbookFromFile(files[0]);
        const { details, expectedTotal, foundTotal, counts } = auditWorkbook(wb, matrixEntry);
        const missing = details.reduce((sum,d) => sum + Math.max(0, -d.difference), 0);
        const extra = details.reduce((sum,d) => sum + Math.max(0, d.difference), 0);
        Object.assign(base, { status: missing === 0 && extra === 0 ? 'OK' : 'DIVERGENCIA', expected: expectedTotal, found: foundTotal, missing, extra, details, ...counts });
      } catch (error) { Object.assign(base, { status: 'ERRO', error: error.message || String(error) }); }
      processed++;
    }
    rows.push(base);
    if (processed % 5 === 0) await new Promise(resolve => setTimeout(resolve, 0));
  }
  return { schools: rows, last_run: { run_at: new Date().toISOString(), total_files: inventoryFiles.length }, upload: { inventory_files: inventoryFiles.length, invalid_files: invalidFiles, matrix_file: matrixFile.name } };
}

function fmtDate(value) { if (!value) return '—'; return new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short' }).format(new Date(value)); }
function escapeHtml(value) { return String(value ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
function showNotice(message, error = false) { const node = $('#notice'); node.textContent = message; node.classList.toggle('error', error); node.classList.remove('hidden'); }
function sum(rows, field) { return rows.reduce((total, row) => total + Number(row[field] || 0), 0); }
function scopedSchools() { return selectedSchools.size ? schools.filter(row => selectedSchools.has(row.cie)) : schools; }
function rankValue(row, rank) { return rank === 'all' ? Number(row.equipment_ok||0)+Number(row.maintenance||0)+Number(row.damaged||0) : Number(row[rank]||0); }

function render(data) {
  $('#databasePath').textContent = 'Processamento local no navegador';
  schools = data.schools || [];
  for (const cie of [...selectedSchools]) if (!schools.some(row => row.cie === cie)) selectedSchools.delete(cie);
  renderSchoolOptions(); renderDashboard(data.last_run);
}

function renderDashboard(run = null) {
  const scoped = scopedSchools(), rank = $('#rankFilter').value;
  $('#totalFiles').textContent = scoped.filter(row => row.file_name).length;
  $('#equipmentOk').textContent = sum(scoped,'equipment_ok').toLocaleString('pt-BR');
  $('#maintenanceCount').textContent = sum(scoped,'maintenance').toLocaleString('pt-BR');
  $('#damagedCount').textContent = sum(scoped,'damaged').toLocaleString('pt-BR');
  const tracked = sum(scoped,'equipment_ok') + sum(scoped,'maintenance') + sum(scoped,'damaged');
  const unclassified = sum(scoped,'status_unclassified');
  $('#statusCoverage').textContent = `${tracked.toLocaleString('pt-BR')} equipamentos com status classificados${unclassified ? ` · ${unclassified.toLocaleString('pt-BR')} sem status informado` : ''}.`;
  document.querySelectorAll('.metric').forEach(card => card.classList.remove('active'));
  const activeId = {equipment_ok:'equipmentOk',maintenance:'maintenanceCount',damaged:'damagedCount'}[rank]; if (activeId) $(`#${activeId}`).closest('.metric').classList.add('active');
  const scopeText = selectedSchools.size ? `${scoped.length} escola(s) selecionada(s)` : `${scoped.length} escolas`;
  const runAt = run?.run_at || document.body.dataset.lastRun; if (run?.run_at) document.body.dataset.lastRun = run.run_at;
  $('#lastRun').textContent = runAt ? `Último processamento: ${fmtDate(runAt)} · ${scopeText} · Ranking: ${rankLabels[rank]}` : 'Nenhuma auditoria executada.';
  renderRows();
}

function renderRows() {
  const query = $('#search').value.trim().toLocaleUpperCase('pt-BR'), rank = $('#rankFilter').value;
  const filtered = scopedSchools().filter(row => !query || `${row.cie} ${row.school_name}`.toLocaleUpperCase('pt-BR').includes(query)).sort((a,b) => rankValue(b,rank)-rankValue(a,rank) || a.school_name.localeCompare(b.school_name,'pt-BR'));
  $('#results').innerHTML = filtered.length ? filtered.map((row,index)=>`<tr data-cie="${escapeHtml(row.cie)}" tabindex="0"><td class="number rank">${index+1}º</td><td><strong>${escapeHtml(row.cie)}</strong></td><td class="school">${escapeHtml(row.school_name)}</td><td class="file" title="${escapeHtml(row.file_name)}">${escapeHtml(row.file_name||'—')}</td><td><span class="badge ${escapeHtml(row.status)}">${escapeHtml(labels[row.status]||row.status)}</span></td><td class="number">${row.expected}</td><td class="number">${row.found}</td><td class="number">${row.missing}</td><td class="number">${row.extra}</td><td class="number status-ok">${row.equipment_ok||0}</td><td class="number status-maintenance">${row.maintenance||0}</td><td class="number status-damaged">${row.damaged||0}</td></tr>`).join('') : '<tr><td colspan="12" class="empty">Nenhuma escola encontrada para esse filtro.</td></tr>';
}

function updateSchoolPickerLabel() { const button=$('#schoolPickerButton'); if(!selectedSchools.size) button.textContent='Todas as escolas'; else if(selectedSchools.size===1){const row=schools.find(item=>selectedSchools.has(item.cie)); button.textContent=row?`${row.cie} · ${row.school_name}`:'1 escola selecionada';} else button.textContent=`${selectedSchools.size} escolas selecionadas`; }
function renderSchoolOptions(){const query=$('#schoolPickerSearch').value.trim().toLocaleUpperCase('pt-BR'); const options=schools.filter(row=>!query||`${row.cie} ${row.school_name}`.toLocaleUpperCase('pt-BR').includes(query)); $('#schoolOptions').innerHTML=options.map(row=>`<label class="school-option"><input type="checkbox" value="${escapeHtml(row.cie)}" ${selectedSchools.has(row.cie)?'checked':''}><span><strong>${escapeHtml(row.cie)}</strong> — ${escapeHtml(row.school_name)}</span></label>`).join(''); updateSchoolPickerLabel();}

function openDetails(cie) {
  const row=schools.find(item=>item.cie===cie); if(!row)return;
  $('#detailTitle').textContent=`${row.cie} · ${row.school_name}`;
  const summary=`<div class="detail-summary"><div><span>Esperado</span><strong>${row.expected}</strong></div><div><span>Encontrado</span><strong>${row.found}</strong></div><div><span>Faltando</span><strong>${row.missing}</strong></div><div><span>Sobrando</span><strong>${row.extra}</strong></div><div class="status-ok"><span>Equipamentos OK</span><strong>${row.equipment_ok||0}</strong></div><div class="status-maintenance"><span>Garantia / manutenção</span><strong>${row.maintenance||0}</strong></div><div class="status-damaged"><span>Danificados / inservíveis</span><strong>${row.damaged||0}</strong></div></div>`;
  const error=row.error?`<p class="error-text">${escapeHtml(row.error)}</p>`:'';
  const comparison=row.details?.length?`<h3 class="detail-section-title">Comparação com a matriz</h3><div class="table-wrap comparison-table-wrap"><table class="comparison-table"><thead><tr><th>Equipamento</th><th>Grupo</th><th>Fabricante</th><th class="number">Esperado</th><th class="number">Encontrado</th><th class="number status-ok">OK</th><th class="number status-maintenance">Garantia / manutenção</th><th class="number status-damaged">Danificados / inservíveis</th></tr></thead><tbody>${row.details.map(item=>`<tr><td>${escapeHtml(item.names.join(' + '))}</td><td>${escapeHtml(item.category)}</td><td>${escapeHtml(item.manufacturer||'—')}</td><td class="number">${item.quantity}</td><td class="number found-cell"><span class="found-bubble">${item.found}</span></td><td class="number status-ok">${item.equipment_ok||0}</td><td class="number status-maintenance">${item.maintenance||0}</td><td class="number status-damaged">${item.damaged||0}</td></tr>`).join('')}</tbody></table></div>`:'';
  $('#detailBody').innerHTML=summary+error+comparison; $('#detailDialog').showModal();
}

$('#matrixFile').addEventListener('change',e=>{$('#matrixFileLabel').textContent=e.target.files[0]?.name||'Selecione a planilha usada como matriz.'; if(e.target.files[0] && driveSourceLoaded) $('#driveMatrixSelect').value='';});
$('#inventoryFiles').addEventListener('change',e=>{const files=[...e.target.files].filter(f=>f.name.toLowerCase().endsWith('.xlsx')&&!f.name.startsWith('~$')); manualInventoryFiles=files; driveSourceLoaded=false; renderDriveMatrixOptions(); $('#inventoryFileLabel').textContent=files.length?`${files.length} arquivo(s) .xlsx selecionado(s).`:'Selecione uma pasta local com os inventários .xlsx.';});
$('#chooseFolderButton').addEventListener('click',async()=>{if(!window.showDirectoryPicker)return showNotice('Seu navegador não oferece seleção persistente de pastas. Use o seletor de arquivos abaixo.',true);try{const handle=await window.showDirectoryPicker({mode:'read'});await saveFolderHandle(handle);await useManualFolder(handle);$('#lastFolderButton').classList.remove('hidden');showNotice(`Pasta manual carregada: ${manualInventoryFiles.length} arquivo(s) disponível(is).`);}catch(error){if(error.name!=='AbortError')showNotice(error.message||String(error),true);}});
$('#lastFolderButton').addEventListener('click',async()=>{try{const handle=await getFolderHandle();if(!handle)return showNotice('Nenhuma pasta manual foi memorizada.',true);await useManualFolder(handle);showNotice(`Última pasta carregada: ${manualInventoryFiles.length} arquivo(s) disponível(is).`);}catch(error){showNotice('Não foi possível reabrir a última pasta. Selecione-a novamente.',true);}});
getFolderHandle().then(handle=>{$('#lastFolderButton').classList.toggle('hidden',!handle);}).catch(()=>{});
$('#loadDriveButton').addEventListener('click', async()=>{
  const button=$('#loadDriveButton'); button.disabled=true; button.textContent='Carregando…';
  try {
    driveFiles = await listDriveFiles(); driveSourceLoaded = true; forceDriveRefresh = true; manualInventoryFiles = []; renderDriveMatrixOptions();
    $('#driveFileLabel').textContent = `${driveFiles.length} arquivo(s) .xlsx encontrado(s). Clique em processar para baixar ou atualizar os arquivos.`;
    showNotice(`Pasta do Google Drive carregada: ${driveFiles.length} arquivo(s) disponível(is). O próximo processamento atualizará o cache.`);
  } catch(error) { showNotice(error.message||String(error),true); }
  finally { button.disabled=false; button.textContent='Atualizar arquivos da pasta'; }
});
$('#driveMatrixSelect').addEventListener('change',()=>{ if($('#driveMatrixSelect').value) $('#matrixFile').value=''; });
$('#scanButton').addEventListener('click', async()=>{
  const button=$('#scanButton');
  let matrix=$('#matrixFile').files[0]; let inventories=manualInventoryFiles.length ? manualInventoryFiles : [...$('#inventoryFiles').files].filter(f=>f.name.toLowerCase().endsWith('.xlsx')&&!f.name.startsWith('~$'));
  if (driveSourceLoaded) {
    const matrixId=$('#driveMatrixSelect').value;
    if(!matrixId)return showNotice('Escolha a planilha matriz entre os arquivos do Google Drive.',true);
    button.disabled=true; button.textContent='Baixando arquivos…'; showNotice(`Baixando ${driveFiles.length} planilha(s) do Google Drive…`);
    try { matrix=await downloadDriveFile(driveFiles.find(file=>file.id===matrixId), { forceRefresh: forceDriveRefresh }); inventories=await downloadDriveFiles(driveFiles.filter(file=>file.id!==matrixId), { forceRefresh: forceDriveRefresh }); forceDriveRefresh = false; }
    catch(error) { showNotice(error.message||String(error),true); button.disabled=false; button.textContent='Processar auditoria'; return; }
  }
  if(!matrix)return showNotice('Selecione a planilha matriz de compras.',true); if(!inventories.length)return showNotice('Carregue a pasta do Google Drive ou selecione os inventários no computador.',true);
  button.disabled=true; button.textContent='Processando…'; showNotice(`Processando ${inventories.length} arquivo(s) localmente no navegador…`);
  try{const data=await executeAudit(matrix,inventories); render(data); const invalid=data.upload.invalid_files.length; showNotice(`Auditoria concluída: ${data.last_run.total_files} arquivo(s) analisado(s)${invalid?` · ${invalid} arquivo(s) ignorado(s) por nome inválido`:''}. Nenhuma planilha foi enviada ao servidor.`);}catch(error){console.error(error);showNotice(error.message||String(error),true);}finally{button.disabled=false;button.textContent='Processar auditoria';}
});
$('#search').addEventListener('input',renderRows); $('#rankFilter').addEventListener('change',()=>renderDashboard());
$('#schoolPickerButton').addEventListener('click',()=>{const menu=$('#schoolPickerMenu');menu.classList.toggle('hidden');$('#schoolPickerButton').setAttribute('aria-expanded',String(!menu.classList.contains('hidden')));});
$('#schoolPickerSearch').addEventListener('input',renderSchoolOptions); $('#schoolOptions').addEventListener('change',event=>{if(!event.target.matches('input[type=checkbox]'))return;if(event.target.checked)selectedSchools.add(event.target.value);else selectedSchools.delete(event.target.value);updateSchoolPickerLabel();renderDashboard();});
$('#selectAllSchools').addEventListener('click',()=>{selectedSchools.clear();renderSchoolOptions();renderDashboard();}); $('#clearSchools').addEventListener('click',()=>{selectedSchools.clear();renderSchoolOptions();renderDashboard();});
document.addEventListener('click',event=>{if(!event.target.closest('.school-picker')){$('#schoolPickerMenu').classList.add('hidden');$('#schoolPickerButton').setAttribute('aria-expanded','false');}});
$('#results').addEventListener('click',event=>{const row=event.target.closest('tr[data-cie]');if(row)openDetails(row.dataset.cie);}); $('#results').addEventListener('keydown',event=>{const row=event.target.closest('tr[data-cie]');if(row&&['Enter',' '].includes(event.key))openDetails(row.dataset.cie);});
$('#closeDialog').addEventListener('click',()=>$('#detailDialog').close()); $('#detailDialog').addEventListener('click',event=>{if(event.target===$('#detailDialog'))$('#detailDialog').close();});

$('#databasePath').textContent='Processamento local no navegador';
