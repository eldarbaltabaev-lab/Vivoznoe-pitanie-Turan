// settings.js

(function () {
  const STORAGE_KEY = 'print-settings-v1';

  const DEFAULTS = {
    // Границы таблицы
    borderColor: '#999999',
    borderWidth: 1,        // px
    borderStyle: 'solid',  // solid / dashed / dotted

    // Текст
    fontSize: 12,          // px
    titleSize: 14,         // px (h3 — заголовок листа)

    // Отступы внутри ячейки
    padY: 6,               // сверху/снизу
    padX: 10,              // слева/справа

    // Отступы страницы @page
    margin: 10,            // mm со всех сторон
    pageSize: 'A4 portrait' // A4 portrait / A4 landscape / A3 portrait / A3 landscape
  };

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return { ...DEFAULTS };
      const parsed = JSON.parse(raw);
      // мержим с дефолтами, чтобы новые поля не терялись
      return { ...DEFAULTS, ...parsed };
    } catch (e) {
      console.warn('Не удалось прочитать настройки, использую дефолты', e);
      return { ...DEFAULTS };
    }
  }

  function save(partial) {
    const current = load();
    const next = { ...current, ...partial };
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch (e) {
      console.warn('Не удалось сохранить настройки', e);
    }
    return next;
  }

  function reset() {
    localStorage.removeItem(STORAGE_KEY);
    return { ...DEFAULTS };
  }

  function getCss() {
    const s = load();
    return `
      @page { size: ${s.pageSize}; margin: ${s.margin}mm; }
      table { border-collapse: collapse; width: 100%; font-size: ${s.fontSize}px; }
      td, th {
        border: ${s.borderWidth}px ${s.borderStyle} ${s.borderColor};
        padding: ${s.padY}px ${s.padX}px;
      }
      h3 { font-size: ${s.titleSize}px; margin: 0 0 6px; }
    `;
  }

  // Наружу — через глобальный объект, чтобы print.js (обычный скрипт) видел
  window.PrintSettings = { DEFAULTS, load, save, reset, getCss };
})();