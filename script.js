import { parseWorkbook } from './excel-parser.js';
import { countWorkbook, sumResults, CATEGORIES } from './counter.js';

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const statusEl = document.getElementById('status');
const resultsEl = document.getElementById('results');
const summaryEl = document.getElementById('summary');

const progressEl = document.getElementById('progress');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');

let processing = false;

/* ---------- Drag & drop ---------- */

dropzone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
  const files = [...(e.target.files ?? [])];
  if (files.length) handleFiles(files);
  fileInput.value = '';
});

['dragenter', 'dragover'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add('hover');
  });
});

['dragleave', 'drop'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove('hover');
  });
});

dropzone.addEventListener('drop', (e) => {
  const files = [...(e.dataTransfer?.files ?? [])]
    .filter((f) => /\.xlsx?$/i.test(f.name));
  if (files.length) handleFiles(files);
});

/* ---------- Обработка ---------- */
window.importedFiles = [];

async function handleFiles(files) {
  if (processing) return;
  processing = true;

  resultsEl.innerHTML = '';
  summaryEl.classList.add('hidden');
  summaryEl.innerHTML = '';

  window.importedFiles = [];

  const total = files.length;
  showProgress(0, total);

  const results = [];
  const errors = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    progressText.textContent = `${i + 1} из ${total}: ${file.name}`;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {type: 'array' });
      window.importedFiles.push({name: file.name, workbook });
      


      const { sheets } = parseWorkbook(buffer);
      const counted = countWorkbook(sheets);
      results.push({ fileName: file.name, ...counted });
    } catch (err) {
      console.error('Ошибка при обработке', file.name, err);
      errors.push({ fileName: file.name, message: err.message });
    }

    setProgress(((i + 1) / total) * 100);
    // даём браузеру отрисовать UI
    await new Promise((r) => setTimeout(r, 0));
  }

  hideProgress();
  statusEl.textContent = `Готово: ${results.length} из ${total} обработано${
    errors.length ? `, ошибок: ${errors.length}` : ''
  }`;

  if (results.length) {
    renderSummary(sumResults(results), results.length);
  }

  renderResults(results, errors);
  processing = false;
}

/* ---------- Прогресс ---------- */

function showProgress(done, total) {
  progressEl.classList.remove('hidden');
  setProgress((done / total) * 100);
  progressText.textContent = `0 из ${total}`;
}

function setProgress(percent) {
  progressFill.style.width = percent + '%';
}

function hideProgress() {
  progressEl.classList.add('hidden');
}

/* ---------- Рендер ---------- */

function renderSummary(sum, fileCount) {
  summaryEl.classList.remove('hidden');
  summaryEl.innerHTML = '';

  const h = document.createElement('h2');
  h.textContent = `Общий итог по ${fileCount} файлам`;
  summaryEl.append(h);

  const grid = document.createElement('div');
  grid.className = 'summary-grid';
  grid.append(
    renderSummaryTable('День', sum.day),
    renderSummaryTable('Ночь', sum.night),
  );
  summaryEl.append(grid);
}

function renderSummaryTable(title, data) {
  const wrap = document.createElement('div');

  const h = document.createElement('h3');
  h.textContent = title;
  wrap.append(h);

  const table = document.createElement('table');
  table.className = 'counts';

  const tbody = document.createElement('tbody');
  for (const key of CATEGORIES) {
    const value = data.counts[key];
    if (key === 'UNKNOWN' && value === 0) continue;

    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${key}</td><td>${value}</td>`;
    if (key === 'LBX') tr.classList.add('lbx');
    if (key === 'UNKNOWN') tr.classList.add('unknown');
    tbody.append(tr);
  }

  table.append(tbody);
  wrap.append(table);

  return wrap;
}

function renderResults(results, errors) {
  for (const r of results) {
    resultsEl.append(renderFileCard(r));
  }
  if (errors.length) {
    resultsEl.append(renderErrors(errors));
  }
}

function renderFileCard(r) {
  const card = document.createElement('section');
  card.className = 'card file-card';

  const title = document.createElement('h3');
  title.className = 'file-title';
  title.textContent = r.fileName;
  card.append(title);

  const loc = document.createElement('div');
  loc.className = 'sheet-loc';
  loc.textContent = `Локация в файле: ${r.rawLocation ?? '—'}`;
  card.append(loc);

  const grid = document.createElement('div');
  grid.className = 'file-grid';
  grid.append(
    renderSummaryTable('День', r.day),
    renderSummaryTable('Ночь', r.night),
  );
  card.append(grid);

  return card;
}

function renderErrors(errors) {
  const card = document.createElement('section');
  card.className = 'card error-card';

  const h = document.createElement('h3');
  h.textContent = `Ошибки (${errors.length})`;
  card.append(h);

  const ul = document.createElement('ul');
  for (const e of errors) {
    const li = document.createElement('li');
    li.textContent = `${e.fileName}: ${e.message}`;
    ul.append(li);
  }
  card.append(ul);

  return card;
}