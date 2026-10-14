// print.js

/**
 * Проверяет лист: есть ли данные в столбце C начиная со строки 10.
 * Читает ячейки напрямую по координатам.
 */
function sheetHasDataInColumnC(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws['!ref']) return false;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const COL_C = 2; // A=0, B=1, C=2
  const START_ROW = 9; // строка 10 Excel = 0-индекс 9

  for (let r = START_ROW; r <= range.e.r; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: COL_C });
    const cell = ws[addr];
    if (!cell || cell.v == null) continue;

    // ❌ Числа игнорируем (это итоги, суммы, 0 и т.п., а не ФИО)
    if (typeof cell.v === 'number') continue;

    const s = String(cell.v).trim();
    if (s === '') continue;

    // ❌ Числовые строки типа "0", "5", "123" тоже игнорируем
    if (/^-?\d+([.,]\d+)?$/.test(s)) continue;

    // ✅ Нашли непустую нецифровую ячейку — значит, лист печатаем
    return true;
  }
  return false;
}

function printAllSheets(importedFiles) {
  if (!importedFiles || !importedFiles.length) {
    alert('Сначала импортируйте файлы');
    return;
  }

  const MAX_SHEETS = 2;
  const skipped = [];
  const perPass = [[], []];

  for (let pass = 0; pass < MAX_SHEETS; pass++) {
    importedFiles.forEach((file) => {
      const sheetName = file.workbook.SheetNames[pass];
      if (!sheetName) return;

      if (!sheetHasDataInColumnC(file.workbook, sheetName)) {
        console.log('❌ Пропускаю:', file.name, '→', sheetName);
        skipped.push({ fileName: file.name, sheetName });
        return;
      }

      console.log('✅ Печатаю:', file.name, '→', sheetName);
      perPass[pass].push({ file, sheetName });
    });
  }

  if (skipped.length) {
    const lines = skipped.map(
      (s, i) => `${i + 1}. ${s.fileName} — лист «${s.sheetName}»`
    );
    alert(
      'Эти листы не отправлены на печать (в столбце C, начиная со строки 10, нет данных):\n\n' +
        lines.join('\n')
    );
  }

  const totalToPrint = perPass[0].length + perPass[1].length;
  if (!totalToPrint) return;

  const iframe = document.createElement('iframe');
  iframe.style.cssText =
    'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;
  doc.open();
  const userCss = window.PrintSettings
  ? PrintSettings.getCss()
  : `@page { size: A4 portrait; margin: 10mm; }
     table { border-collapse: collapse; width: 100%; font-size: 12px; }
     td, th { border: 1px solid #999; padding: 6px 10px; }
     h3 { font-size: 14px; margin: 0 0 6px; }`;

doc.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>
  body { font-family: Arial, sans-serif; margin: 0; }
  .sheet { page-break-after: always; }
  .sheet:last-child { page-break-after: auto; }
  ${userCss}
</style></head><body>`);

  const esc = (s) =>
    String(s).replace(
      /[&<>"']/g,
      (m) =>
        ({
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#39;',
        }[m])
    );

  for (let pass = 0; pass < MAX_SHEETS; pass++) {
    perPass[pass].forEach(({ file, sheetName }) => {
      const sheet = file.workbook.Sheets[sheetName];
      const html = XLSX.utils.sheet_to_html(sheet, { editable: false });

      doc.write(`<div class="sheet">
        <h3>${esc(file.name)} — ${esc(sheetName)}</h3>
        ${html}
      </div>`);
    });
  }

  doc.write('</body></html>');
  doc.close();

  setTimeout(() => {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
  }, 500);
}
