// XLSX приходит из <script> в index.html (глобальная переменная)

/**
 * Стоп-слова в столбце с именем → это не сотрудник, а должность.
 * Регистр и алфавит — любые.
 */
 const NON_EMPLOYEE_PATTERNS = [
  /coffee[\s-]*lady/i,
  /кофе[\s-]*леди/i,
  /кофеледи/i,
  /cleaning[\s-]*lady/i,
  /клининг[\s-]*леди/i,
  /горничн/i,
  /раздач/i,
  /\bcleaning\b/i,
  /\bcoffee\b/i,
];

/**
 * Стоп-названия компании-подрядчика в столбце G (индекс 6).
 */
const CONTRACTOR_PATTERNS = [
  /eurasia\s*support\s*services/i,
  /eurasia\s*ss\b/i,
  /eurasian\s*service\s*solutions/i,
  /eurasia\s*service\s*solutions/i,
  /\beuss\b/i,
];

export function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheets = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: false,
      defval: '',
    });

    const isNight = /night|ночь/i.test(sheetName);
    const rawLocation = findRawLocation(rows);
    const employees = findEmployees(rows);

    sheets.push({ name: sheetName, isNight, rawLocation, employees });
  }

  return { sheets };
}

function findRawLocation(rows) {
  const MAX_COL = 10;
  for (const row of rows) {
    const limit = Math.min(row.length, MAX_COL);
    for (let c = 0; c < limit; c++) {
      const cell = String(row[c] ?? '');
      const idx = cell.search(/НАЗВАНИЕ\s*ОБЪЕКТА/i);
      if (idx !== -1) {
        const leftInCell = cell.slice(0, idx).trim();
        if (leftInCell) return leftInCell;

        for (let k = c - 1; k >= 0; k--) {
          const left = String(row[k] ?? '').trim();
          if (left) return left;
        }

        const wholeRow = row
          .slice(0, MAX_COL)
          .map((v) => String(v ?? ''))
          .join(' ')
          .trim();
        return wholeRow || null;
      }
    }
  }
  return null;
}

function findEmployees(rows) {
  const MAX_COL = 10;
  const CONTRACTOR_COL = 6; // столбец G
  let headerRow = -1;
  let nameCol = -1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const limit = Math.min(row.length, MAX_COL);
    for (let c = 0; c < limit; c++) {
      const cell = String(row[c] ?? '');
      if (/Ф\.?\s*И\.?\s*О\.?|Name\s*&\s*Surname/i.test(cell)) {
        headerRow = r;
        nameCol = c;
        break;
      }
    }
    if (headerRow !== -1) break;
  }

  if (headerRow === -1) return [];

  const employees = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    const nameCell = String(row[nameCol] ?? '').trim();

    if (!nameCell) continue;
    if (/^\d+([.,]\d+)?$/.test(nameCell)) break;

    // ❌ Стоп-слово в имени (должность, а не человек)
    if (NON_EMPLOYEE_PATTERNS.some((re) => re.test(nameCell))) continue;

    // ❌ Название подрядчика в столбце G
    const contractor = String(row[CONTRACTOR_COL] ?? '').trim();
    if (CONTRACTOR_PATTERNS.some((re) => re.test(contractor))) continue;

    const rowText = row
      .slice(0, MAX_COL)
      .map((v) => String(v ?? ''))
      .join(' ');

    employees.push({ name: nameCell, rowText });
  }

  return employees;
}