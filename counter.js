import { resolveLocation, isLunchbox } from './classifier.js';

export const CATEGORIES = [
  'ABK', 'MMA', 'OXD', 'SP1', 'SP2', 'WS1', 'WS2', 'LBX', 'UNKNOWN',
];

export function emptyCounts() {
  return Object.fromEntries(CATEGORIES.map((k) => [k, 0]));
}

/**
 * Считает один лист.
 * → { counts, unknownRaw }
 */
export function countSheet(sheet) {
  const counts = emptyCounts();
  const unknownRaw = new Set();

  const loc = resolveLocation(sheet.rawLocation);

  for (const emp of sheet.employees) {
    if (isLunchbox(emp.rowText)) {
      counts.LBX++;
      continue;
    }
    if (loc.key === 'UNKNOWN') {
      counts.UNKNOWN++;
      if (loc.raw) unknownRaw.add(loc.raw);
    } else {
      counts[loc.key]++;
    }
  }

  return { counts, unknownRaw: [...unknownRaw] };
}

/**
 * Считает весь workbook, разделяя на День и Ночь.
 * → {
 *     day:   { counts, unknownRaw },
 *     night: { counts, unknownRaw },
 *     rawLocation: 'ABK'   // из первого найденного листа
 *   }
 */
export function countWorkbook(sheets) {
  const day = { counts: emptyCounts(), unknownRaw: new Set() };
  const night = { counts: emptyCounts(), unknownRaw: new Set() };
  let rawLocation = null;

  for (const sheet of sheets) {
    if (!rawLocation && sheet.rawLocation) rawLocation = sheet.rawLocation;

    const target = sheet.isNight ? night : day;
    const { counts, unknownRaw } = countSheet(sheet);

    for (const key of CATEGORIES) target.counts[key] += counts[key];
    for (const r of unknownRaw) target.unknownRaw.add(r);
  }

  return {
    day: { counts: day.counts, unknownRaw: [...day.unknownRaw] },
    night: { counts: night.counts, unknownRaw: [...night.unknownRaw] },
    rawLocation,
  };
}

/** Суммирует несколько результатов (по файлам) в один. */
export function sumResults(results) {
  const day = { counts: emptyCounts(), unknownRaw: new Set() };
  const night = { counts: emptyCounts(), unknownRaw: new Set() };

  for (const r of results) {
    for (const key of CATEGORIES) {
      day.counts[key] += r.day.counts[key];
      night.counts[key] += r.night.counts[key];
    }
    for (const u of r.day.unknownRaw) day.unknownRaw.add(u);
    for (const u of r.night.unknownRaw) night.unknownRaw.add(u);
  }

  return {
    day: { counts: day.counts, unknownRaw: [...day.unknownRaw] },
    night: { counts: night.counts, unknownRaw: [...night.unknownRaw] },
  };
}