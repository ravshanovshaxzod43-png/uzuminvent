// ==UserScript==
// @name         Uzum Ревизор ПВЗ
// @namespace    https://dp.uzum.uz/
// @version      1.8.1
// @description  Ревизия склада ПВЗ Uzum Market: проверка ID точки, потоварная инвентаризация, приёмка, возвраты, пересорт, партнёрские заказы (SX/RX), время работы, история и Excel-акт
// @author       Хожиев Сухроб
// @match        https://dp.uzum.uz/*
// @icon         https://dp.uzum.uz/favicon.ico
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  if (window.__uzumInvLoaded) return;
  window.__uzumInvLoaded = true;

  const VERSION = '1.8.1';
  const API = 'https://api-wms.uzum.uz/de/delivery-point';
  const API_V3 = 'https://api-wms.uzum.uz/de/v3/delivery-point';
  const HISTORY_KEY = 'uzum_inventory_history';
  const HISTORY_LIMIT = 100;
  const NO_CELL = 'БЕЗ_ЯЧЕЙКИ';
  const BTN_GAP = 40;
  const RING_R = 32;
  const RING_C = 2 * Math.PI * RING_R;
  let IDLE_LIMIT_MS = 2 * 60 * 1000;
  const SETTINGS_KEY = 'uzum_inventory_settings';
  const DPKEY_RE = /^[A-Za-z0-9_-]{1,40}$/;
  const jwtRegex = /eyJ[a-zA-Z0-9_-]{25,}\.eyJ[a-zA-Z0-9_-]{25,}\.[a-zA-Z0-9_-]{25,}/g;
  const PARTNER_RE = /^(SX|RX)(?![A-Z])/i;

  // ---------------------------------------------------------------- telegram
  // Token va chat ID lar sozlamalar bo'limida qo'lda kiritiladi.
  const TG_BOT_PREFIX = '9289308469:';
  const TG_TOKEN_RE = /^9289308469:[A-Za-z0-9_-]{20,}$/;
  const TG_CHAT_RE = /^-?\d{3,}$/;
  const getTgToken = () => String(getSettings().tgToken || '').trim();
  const getTgChats = () => String(getSettings().tgChats || '')
    .split(/[\s,;]+/).map(s => s.trim()).filter(Boolean);

  // ---------------------------------------------------------------- settings
  // Har bir funksiya sozlamalar bo'limidan yoqilib/o'chiriladi.
  const SETTINGS_DEFAULTS = {
    sound: true,          // skaner ovozlari
    tgToken: '',          // Telegram bot tokeni (9289308469:... bilan boshlanadi)
    tgChats: '',          // chat ID lar, vergul bilan ajratilgan (1 yoki bir nechta)
    autoFocus: true,      // skaner maydonini avtomatik fokus
    idleMinutes: 2,       // tanaffus chegarasi (daqiqa), ish vaqti uchun
    lang: 'auto',         // auto | ru | uz
    hotkeys: true         // Ctrl+Z — oxirgi skanni bekor qilish
  };
  const SETTINGS_META = {
    sound: { type: 'bool' },
    tgToken: { type: 'str' },
    tgChats: { type: 'str' },
    autoFocus: { type: 'bool' },
    idleMinutes: { type: 'num', min: 1, max: 10 },
    lang: { type: 'enum', values: ['auto', 'ru', 'uz'] },
    hotkeys: { type: 'bool' }
  };

  function getSettings() {
    let saved = {};
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      if (parsed && typeof parsed === 'object') saved = parsed;
    } catch (e) {}
    const out = { ...SETTINGS_DEFAULTS };
    Object.keys(SETTINGS_DEFAULTS).forEach(k => {
      if (saved[k] !== undefined) out[k] = saved[k];
    });
    return out;
  }

  function setSetting(key, value) {
    const meta = SETTINGS_META[key];
    if (!meta) return;
    let v = value;
    if (meta.type === 'bool') v = !!value;
    if (meta.type === 'str') v = String(value == null ? '' : value).trim();
    if (meta.type === 'num') v = Math.min(meta.max, Math.max(meta.min, Number(value) || SETTINGS_DEFAULTS[key]));
    if (meta.type === 'enum' && !meta.values.includes(value)) v = SETTINGS_DEFAULTS[key];
    const all = getSettings();
    all[key] = v;
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(all)); } catch (e) {}
    applySettings();
  }

  // Sozlamalarni ishlayotgan kodga qo'llash
  function applySettings() {
    const s = getSettings();
    IDLE_LIMIT_MS = s.idleMinutes * 60 * 1000;
    SOUND_ON = !!s.sound;
    AUTO_FOCUS = !!s.autoFocus;
    LANG_PREF = s.lang;
    HOTKEYS_ON = !!s.hotkeys;
  }
  let SOUND_ON = true;
  let AUTO_FOCUS = true;
  let LANG_PREF = 'auto';
  let HOTKEYS_ON = true;
  applySettings();

  // ---------------------------------------------------------------- i18n
  const I18N = {
    ru: {
      appName: 'Ревизор ПВЗ',
      navNew: 'Новая ревизия',
      navHist: 'История ревизий',
      navSet: 'Настройки',
      setTitle: 'Настройки',
      setSub: 'Включайте функции и настраивайте их под себя. Изменения сохраняются автоматически.',
      setSound: 'Звуки сканера',
      setSoundD: 'Звуковые сигналы при каждом скане',
      setTg: 'Отчёт в Telegram',
      setTgD: 'Отправлять итог ревизии в Telegram при сохранении',
      setFocus: 'Автофокус сканера',
      setFocusD: 'Возвращать курсор в поле ввода после клика по странице',
      setIdle: 'Порог простоя (мин)',
      setIdleD: 'Пауза дольше этого значения не учитывается в рабочем времени',
      setLang: 'Язык интерфейса',
      setLangD: 'Авто — определяется по языку сайта',
      langAuto: 'Авто',
      setReset: 'Сбросить настройки',
      setResetOk: 'Настройки сброшены к значениям по умолчанию.',
      setOn: 'Вкл',
      setOff: 'Выкл',
      setHot: 'Горячие клавиши',
      setHotD: 'Ctrl+Z — отменить последний скан (на экране сканирования)',
      setTgTok: 'Telegram bot токен',
      setTgTokD: 'Обязателен. Начинается с 9289308469: — без него ревизию не начать',
      setTgChat: 'Telegram ID чатов',
      setTgChatD: 'Обязательны. Несколько ID через запятую — отчёт уйдёт во все',
      tgNeedToken: '⚠️ Введите Telegram bot токен в Настройках. Без токена ревизия не начинается.',
      tgBadToken: p => `⚠️ Неверный токен. Он должен начинаться с ${p}`,
      tgTokenFormatOk: 'Формат токена верный',
      tgNeedChat: '⚠️ Введите хотя бы один Telegram ID чата в Настройках.',
      tgBadChat: '⚠️ Неверный ID чата. Допустимы только цифры (можно с «-» в начале).',
      tgCheckFail: '⚠️ Токен не прошёл проверку в Telegram (getMe). Проверьте токен и подключение.',
      close: 'Закрыть',
      stepParams: 'Параметры',
      stepScan: 'Сканирование',
      stepReport: 'Итоги',
      setupTitle: 'Параметры ревизии',
      setupHint: 'Проверьте ПВЗ и ID точки, затем выгрузите остатки.',
      scanTitle: pvz => `Сканирование · ${pvz}`,
      scanSub: 'Сначала отсканируйте ячейку, затем товары на ней.',
      reportTitle: pvz => `Итоги ревизии · ${pvz}`,
      reportSub: 'Сводка по всем позициям и рекомендации по каждой.',
      histTitle: 'Архив ревизий',
      histSub: 'Ревизии, сохранённые в этом браузере.',
      pvz: 'ПВЗ',
      dpKeyLabel: 'ID точки (dpKey)',
      issueLabel: 'Заказов к выдаче',
      ordersFromPage: 'Берётся со страницы «Заказы клиентов».',
      loadBtn: 'Выгрузить остатки и начать',
      loadingBtn: 'Проверка...',
      statOrders: 'Заказов к выдаче',
      statPlan: 'Товаров по плану',
      statScanned: 'Отсканировано',
      barLabel: (f, tot) => `Найдено ${f} из ${tot} шт`,
      timerLabel: 'Ish vaqti',
      cellActive: 'Активная ячейка',
      cellPlaceholder: 'Отсканируйте ячейку',
      lastBarcode: 'Последний код',
      scanPlaceholder: 'Сканируйте код или введите вручную и нажмите Enter',
      hintNeedCell: 'Сначала отсканируйте штрихкод ячейки',
      hintScanItems: 'Сканируйте товары на этой полке',
      undo: 'Отменить скан',
      restartBtn: 'Новая ревизия',
      restartConfirm: 'Текущая ревизия будет сброшена. Продолжить?',
      cellItems: 'Учтено в текущей ячейке',
      finish: 'Итоги ревизии',
      emptyCell: 'Ячейка пуста. Сканируйте заказы на этой полке.',
      colOrderBarcode: 'Заказ / Код',
      colQty: 'Кол-во',
      badgeAcc: 'В ПРИЁМКЕ',
      badgeRet: 'ВОЗВРАТ',
      badgeForeign: 'ЧУЖОЙ ПВЗ',
      badgeBank: 'UZUM BANK',
      badgePartner: 'ПАРТНЁР',
      orderLabel: n => `Заказ №${n}`,
      pcs: n => `${n} шт`,
      cellWord: 'ячейка',
      msgNeedCell: '⚠️ Сначала отсканируйте штрихкод ячейки (полки)!',
      msgAlreadyFull: pid => `ℹ️ Заказ №${pid} уже учтён полностью.`,
      msgAccepted: (pid, name) => `✓ Заказ №${pid} (${name}) принят.`,
      msgSearching: '🔍 Проверка статуса в WMS...',
      msgAcceptanceHint: '⚠️ В ПРИЁМКЕ. Оставить в ячейке! Завершить приёмку в WMS.',
      msgAcceptanceDup: 'ℹ️ Эта позиция приёмки уже учтена.',
      msgReturnHint: '🔴 ВОЗВРАТ / ОТМЕНЁН! Срочно убрать с полки и отправить на склад.',
      msgReturnDup: 'ℹ️ Эта позиция возврата уже учтена.',
      msgForeignHint: '⛔ Чужой товар или посылка. Не числится на этом ПВЗ — отложить в возврат!',
      msgUndone: code => `↩️ Скан «${code}» отменён.`,
      msgNothingToUndo: 'Нечего отменять.',
      statusAcceptance: 'Размещён / В приёмке',
      statusReturn: 'Завершён (к возврату)',
      statusAcceptanceShort: 'В приёмке',
      statusOk: 'Совпадает',
      statusShort: 'Недостача',
      statusMis: 'Не на полке',
      statusBoth: 'Недостача и не на полке',
      notFoundShort: 'Не найден',
      errNoToken: 'Токен не найден. Обновите вкладку «Заказы» (F5) и повторите.',
      errNoAccess: s => `Нет доступа к WMS API (${s}). Обновите страницу (F5) и повторите.`,
      errApi: s => `Ошибка WMS API (${s}).`,
      errNoDpKey: 'ID ПВЗ не указан. Введите ID точки (dpKey) вручную.',
      loadAuth: 'Авторизация в WMS...',
      loadList: k => `Выгрузка реестра заказов (dpKey: ${k})...`,
      loadCollected: n => `Собрано заказов к выдаче: ${n}...`,
      loadItems: (a, b) => `Загрузка состава товаров: ${a} из ${b}`,
      confirmFailed: n => `Не удалось загрузить состав ${n} заказов (для них план принят как 1 шт).\nПродолжить ревизию?`,
      errLoadCancelled: 'Загрузка отменена. Повторите попытку.',
      dpChecking: 'Проверка ID точки в WMS...',
      dpEmpty: 'ID точки (dpKey) не определён. Введите его вручную и нажмите кнопку ещё раз.',
      dpFormat: 'ID точки может содержать только латинские буквы, цифры, «-» и «_».',
      dpInvalid: id => `ID ${id} не подтверждён WMS. Проверьте значение или обновите страницу (F5).`,
      dpOk: id => `ID ${id} подтверждён WMS`,
      dpNotChecked: 'ID ещё не проверен. Нажмите «Выгрузить остатки».',
      dpMismatch: (page, api) => `На странице «К выдаче»: ${page}, а в WMS по этому ID: ${api}. ID может быть неверным. Продолжить?`,
      alreadySaved: 'Эта ревизия уже сохранена. Для новой нажмите «Новая ревизия».',
      savedOk: pvz => `✓ Ревизия ПВЗ ${pvz} сохранена в базу!\nАкт доступен во вкладке «История».`,
      saveFail: '⚠️ Не удалось сохранить ревизию в базу браузера (нет места). Акт .xlsx будет скачан.',
      confirmDelete: 'Удалить эту ревизию из базы?',
      histEmpty: 'В базе пока нет сохранённых ревизий. После проверки нажмите «Итоги ревизии» → «Завершить целиком».',
      histNotFound: 'Ничего не найдено.',
      histSearch: 'Поиск по PID или штрихкоду',
      chipOrders: n => `Заказов ${n}`,
      chipPlan: n => `План ${n}`,
      chipFact: n => `Факт ${n}`,
      chipShort: n => `Недостач ${n}`,
      chipMis: n => `Не на полке ${n}`,
      chipAcc: n => `Приёмка ${n}`,
      chipRet: n => `Возвраты ${n}`,
      chipPartner: n => `Партнёры ${n}`,
      chipTime: (tot, act) => `Время ${tot} · работа ${act}`,
      btnXlsx: 'Скачать .xlsx',
      btnSave: 'Завершить целиком',
      btnSaved: 'Сохранено ✓',
      btnExcel: 'Скачать .xlsx',
      btnBack: 'Вернуться к сканированию',
      deleteTitle: 'Удалить из базы',
      metaOrders: 'Заказов',
      metaPlan: 'По плану',
      metaScanned: 'Отсканировано',
      timeStart: 'Начало',
      timeTotal: 'Общее время',
      timeActive: 'Рабочее время',
      timeSpeed: 'Скорость',
      speedVal: n => `${n} шт/ч`,
      speedNA: '—',
      tileShort: 'Недостачи',
      tileMis: 'Не на полке',
      tileAcc: 'В приёмке',
      tileRet: 'Возвраты',
      tileSur: 'Чужие',
      tilePartner: 'Партнёрские',
      ofTotal: n => `из ${n}`,
      issuesPill: n => `проблем: ${n}`,
      secShort: 'Недостачи',
      secMis: 'Заказы не на своей полке',
      secAcc: 'В приёмке (не приняты полностью)',
      secRet: 'Отменённые / возвраты',
      secSur: 'Чужие товары / не числятся',
      secPartner: 'Партнёрские заказы (SX / RX)',
      pos: n => `${n} поз.`,
      noShort: 'Недостач не обнаружено. Все клиентские заказы на месте!',
      noMis: 'Все заказы лежат строго на своих местах.',
      noAcc: 'Позиций в статусе приёмки на полках нет.',
      noRet: 'Отменённых заказов на полках не обнаружено.',
      noSur: 'Чужих товаров на складе не обнаружено.',
      noPartner: 'Партнёрских заказов (SX / RX) в этой ревизии нет.',
      thPid: 'Заказ (PID)',
      thMissing: 'Не хватает',
      thShelf: 'Полка по плану',
      thClient: 'Клиент',
      thQty: 'Кол-во',
      thPlanCell: 'По плану',
      thActualCell: 'Лежит фактически',
      thBarcodeLoc: 'Код / Грузоместо',
      thRec: 'Рекомендация',
      thBarcode: 'Код',
      thExtra: 'Лишнее кол-во',
      thFoundAt: 'Где найден',
      thStatus: 'Статус',
      thPlanFact: 'План / Факт',
      thLocation: 'Где лежит',
      recAcc: 'Не отправлять на склад! Оставить в ячейке, провести приёмку.',
      recRet: 'Снять с полки! Оформить отгрузку на возвратный склад.',
      recPS: 'Товара нет на складе — проверить у партнёра.',
      recPM: 'Переложить на полку по плану.',
      recPB: 'Переложить на полку и проверить недостачу у партнёра.',
      recPO: 'Замечаний нет.',
      cellPrefix: 'ячейка ',
      unknown: 'Не определён',
      mgrName: 'Хожиев Сухроб',
      noName: '—',
      tgTitle: '📦 Ревизия ПВЗ завершена',
      tgPvz: (p, k) => `ПВЗ: ${p} (ID: ${k})`,
      tgDate: (d, tm) => `Дата: ${d} ${tm}`,
      tgAuditor: a => `Ревизор: ${a}`,
      tgOrders: n => `Заказов: ${n}`,
      tgPlan: n => `План: ${n} шт`,
      tgFact: n => `Факт: ${n} шт`,
      tgTime: (tot, act) => `⏱ Время: ${tot} (рабочее: ${act})`,
      tgShort: (c, u) => `🔴 Недостачи: ${c} поз. (-${u} шт)`,
      tgMis: (c, u) => `🟡 Не на полке: ${c} поз. (${u} шт)`,
      tgAcc: (c, u) => `⚠️ В приёмке: ${c} поз. (${u} шт)`,
      tgRet: (c, u) => `🔴 Возвраты: ${c} поз. (${u} шт)`,
      tgSur: (c, u) => `⛔ Чужие: ${c} поз. (+${u} шт)`,
      tgPartner: (c, u) => `🤝 Партнёрские (SX/RX) с проблемой: ${c} поз. (${u} шт)`,
      sheetName: 'Акт ревизии',
      xTitle: 'АКТ ИНВЕНТАРИЗАЦИИ СКЛАДА ПВЗ',
      xPvz: 'Пункт выдачи:',
      xDate: 'Дата и время:',
      xMgr: 'Территориальный управляющий:',
      xAuditor: 'Ревизор:',
      xStatus: 'Статус проверки:',
      xStatusDone: 'Завершена целиком (сохранено в базу)',
      xSummary: '=== ИТОГОВАЯ СВОДКА РЕВИЗИИ ===',
      xHdrInd: 'Показатель',
      xHdrQty: 'Количество',
      xHdrStatus: 'Статус',
      xSumOrders: 'Всего клиентских заказов в плане WMS:',
      xSumPlan: 'Товаров по плану:',
      xSumScanned: 'Всего отсканировано товаров:',
      xSumShort: 'Чистая недостача (не найдены на складе):',
      xSumMis: 'Заказы не на своей полке:',
      xSumAcc: 'Заказы в приёмке (не приняты полностью):',
      xSumRet: 'Отменённые / возвраты (к отгрузке на склад):',
      xSumSur: 'Чужие товары (не числятся за ПВЗ):',
      xSumPartner: 'Партнёрские заказы SX/RX с проблемой:',
      xSumTime: 'Затраченное время (общее / рабочее):',
      xStPlanToIssue: 'План к выдаче',
      xStPlanPcs: 'План штук',
      xStFact: 'Факт на складе',
      xStShort: n => `Не хватает: -${n} шт`,
      xStMis: n => `${n} шт требуют перемещения`,
      xStAcc: n => `${n} шт (оставить в ячейке, провести приёмку)`,
      xStRet: n => `${n} шт (срочно снять с полки на склад)`,
      xStSur: n => `+${n} шт в коробку возвратов`,
      xStPartner: n => `${n} шт с проблемой`,
      xStTime: (act, s, e) => `Рабочее: ${act} · с ${s} до ${e}`,
      xS1: 'РАЗДЕЛ 1. НЕДОСТАЧИ (ЗАКАЗЫ, КОТОРЫХ НЕТ НА СКЛАДЕ)',
      xS2: 'РАЗДЕЛ 2. ЗАКАЗЫ НЕ НА СВОЕЙ ПОЛКЕ (ТОВАР НА СКЛАДЕ, НО НЕ НА ПЛАНОВОЙ ПОЛКЕ)',
      xS3: 'РАЗДЕЛ 3. ТОВАРЫ В ПРИЁМКЕ (НЕ ПРИНЯТЫ ПОЛНОСТЬЮ — НЕ ОТПРАВЛЯТЬ НА СКЛАД!)',
      xS4: 'РАЗДЕЛ 4. ОТМЕНЁННЫЕ ЗАКАЗЫ И ВОЗВРАТЫ (СРОЧНО УБРАТЬ С ПОЛКИ И ОТПРАВИТЬ НА СКЛАД)',
      xS5: 'РАЗДЕЛ 5. ЧУЖИЕ ТОВАРЫ И НЕОПОЗНАННЫЕ ИЗЛИШКИ',
      xS6: 'РАЗДЕЛ 6. ПАРТНЁРСКИЕ ЗАКАЗЫ (SX / RX) — ВСЕ ЗАКАЗЫ И ИХ СТАТУС',
      xOkShort: '✓ Недостач не обнаружено. Все клиентские заказы на месте!',
      xOkMis: '✓ Все найденные заказы лежат строго на своих местах.',
      xOkAcc: '✓ Товаров в статусе незавершённой приёмки не обнаружено.',
      xOkRet: '✓ Забытых отменённых заказов или возвратов на полках не обнаружено.',
      xOkSur: '✓ Лишних или чужих посылок на складе не обнаружено.',
      xOkPartner: 'Партнёрских заказов (SX / RX) в ревизии нет.',
      xHdrS1: ['№', 'Номер заказа (PID)', 'ID в WMS', 'Код (SKU / Место)', 'Не хватает', 'Полка по плану', 'Где найден', 'Клиент', 'Телефон'],
      xHdrS2: ['№', 'Номер заказа (PID)', 'ID в WMS', 'Код (SKU / Место)', 'Кол-во', 'Полка по плану', 'Где лежит по факту', 'Клиент', 'Действие'],
      xHdrS3: ['№', 'Считанный код / Место', 'ID заказа', 'Кол-во', 'Полка', 'Статус в WMS', '', '', 'Рекомендуемое действие'],
      xHdrS5: ['№', 'Считанный код', 'Кол-во', 'Где найден (полка)', 'Статус', '', '', '', 'Рекомендуемое действие'],
      xHdrS6: ['№', 'Заказ (PID)', 'ID в WMS', 'Статус', 'План (шт)', 'Факт (шт)', 'Где лежит', 'Клиент', 'Рекомендация'],
      xNotFound: 'НЕ НАЙДЕН',
      xActS2: n => `Переложить в ячейку ${n}`,
      xActS3: 'Оставить в ячейке! Завершить приёмку в WMS',
      xActS4: 'Снять с полки! Оформить отгрузку на возвратный склад',
      xActS5: 'Отложить в коробку возвратов',
      xStatusRetShort: 'Завершён (отгрузка)',
      xForeignStatus: 'Чужой ПВЗ / не числится'
    },

    uz: {
      appName: 'PVZ inventarizatsiya',
      navNew: 'Yangi inventarizatsiya',
      navHist: 'Inventarizatsiyalar tarixi',
      navSet: 'Sozlamalar',
      setTitle: 'Sozlamalar',
      setSub: 'Funksiyalarni yoqing va o\'zingizga moslang. O\'zgarishlar avtomatik saqlanadi.',
      setSound: 'Skaner ovozlari',
      setSoundD: 'Har bir skanerlashda ovozli signal',
      setTg: 'Telegram hisoboti',
      setTgD: 'Saqlashda inventarizatsiya natijasini Telegram ga yuborish',
      setFocus: 'Skaner avtofokusi',
      setFocusD: 'Sahifani bosgandan keyin kursorni kiritish maydoniga qaytarish',
      setIdle: 'Turib qolish chegarasi (daq)',
      setIdleD: 'Bu vaqtdan uzoq pauza ish vaqtiga kiritilmaydi',
      setLang: 'Interfeys tili',
      setLangD: 'Avto — sayt tiliga qarab aniqlanadi',
      langAuto: 'Avto',
      setReset: 'Sozlamalarni tiklash',
      setResetOk: 'Sozlamalar standart qiymatlarga qaytarildi.',
      setOn: 'Yoqilgan',
      setOff: 'O\'chirilgan',
      setHot: 'Tezkor tugmalar',
      setHotD: 'Ctrl+Z — oxirgi skanni bekor qilish (skanerlash ekranida)',
      setTgTok: 'Telegram bot tokeni',
      setTgTokD: 'Majburiy. 9289308469: bilan boshlanadi — bo\'lmasa revizani boshlab bo\'lmaydi',
      setTgChat: 'Telegram chat ID lari',
      setTgChatD: 'Majburiy. Bir nechta ID vergul bilan — hisobot hammasiga boradi',
      tgNeedToken: '⚠️ Sozlamalarda Telegram bot tokenini kiriting. Tokensiz revizani boshlab bo\'lmaydi.',
      tgBadToken: p => `⚠️ Token noto'g'ri. U ${p} bilan boshlanishi kerak`,
      tgTokenFormatOk: 'Token formati to\'g\'ri',
      tgNeedChat: '⚠️ Sozlamalarda kamida bitta Telegram chat ID kiriting.',
      tgBadChat: '⚠️ Chat ID noto\'g\'ri. Faqat raqamlar (boshida «-» bo\'lishi mumkin).',
      tgCheckFail: '⚠️ Token Telegram tekshiruvidan o\'tmadi (getMe). Tokenni va internetni tekshiring.',
      close: 'Yopish',
      stepParams: 'Parametrlar',
      stepScan: 'Skanerlash',
      stepReport: 'Natija',
      setupTitle: 'Inventarizatsiya parametrlari',
      setupHint: 'PVZ va nuqta ID sini tekshiring, so\'ng qoldiqlarni yuklang.',
      scanTitle: pvz => `Skanerlash · ${pvz}`,
      scanSub: 'Avval katakni skanerlang, so\'ng undagi mahsulotlarni.',
      reportTitle: pvz => `Inventarizatsiya natijalari · ${pvz}`,
      reportSub: 'Barcha pozitsiyalar bo\'yicha xulosa va tavsiyalar.',
      histTitle: 'Inventarizatsiyalar arxivi',
      histSub: 'Ushbu brauzerda saqlangan inventarizatsiyalar.',
      pvz: 'PVZ',
      dpKeyLabel: 'Nuqta ID (dpKey)',
      issueLabel: 'Berishga buyurtmalar',
      ordersFromPage: 'Sahifadagi «Mijoz buyurtmalari» bo\'limidan olinadi.',
      loadBtn: 'Qoldiqlarni yuklab boshlash',
      loadingBtn: 'Tekshirilmoqda...',
      statOrders: 'Berishga buyurtmalar',
      statPlan: 'Reja bo\'yicha mahsulotlar',
      statScanned: 'Skanerlangan',
      barLabel: (f, tot) => `Topildi ${f} / ${tot} dona`,
      timerLabel: 'Ish vaqti',
      cellActive: 'Faol katak',
      cellPlaceholder: 'Katakni skanerlang',
      lastBarcode: 'Oxirgi kod',
      scanPlaceholder: 'Skanerlang yoki kodni qo\'lda kiriting va Enter bosing',
      hintNeedCell: 'Avval katak shtrixkodini skanerlang',
      hintScanItems: 'Ushbu javondagi mahsulotlarni skanerlang',
      undo: 'Skanni bekor qilish',
      restartBtn: 'Yangi inventarizatsiya',
      restartConfirm: 'Joriy inventarizatsiya bekor qilinadi. Davom etasizmi?',
      cellItems: 'Joriy katakda hisobga olingan',
      finish: 'Natijalarni ko\'rish',
      emptyCell: 'Katak bo\'sh. Ushbu javondagi buyurtmalarni skanerlang.',
      colOrderBarcode: 'Buyurtma / Kod',
      colQty: 'Miqdor',
      badgeAcc: 'QABULDA',
      badgeRet: 'QAYTARISH',
      badgeForeign: 'BOSHQA PVZ',
      badgeBank: 'UZUM BANK',
      badgePartner: 'HAMKOR',
      orderLabel: n => `Buyurtma №${n}`,
      pcs: n => `${n} dona`,
      cellWord: 'katak',
      msgNeedCell: '⚠️ Avval katak (javon) shtrixkodini skanerlang!',
      msgAlreadyFull: pid => `ℹ️ Buyurtma №${pid} to'liq hisobga olingan.`,
      msgAccepted: (pid, name) => `✓ Buyurtma №${pid} (${name}) qabul qilindi.`,
      msgSearching: '🔍 WMS da holat tekshirilmoqda...',
      msgAcceptanceHint: '⚠️ QABULDA. Katakda qoldiring! WMS da qabulni yakunlang.',
      msgAcceptanceDup: 'ℹ️ Bu qabul pozitsiyasi allaqachon hisobga olingan.',
      msgReturnHint: '🔴 QAYTARISH / BEKOR QILINGAN! Zudlik bilan javondan olib, omborga yuboring.',
      msgReturnDup: 'ℹ️ Bu qaytarish pozitsiyasi allaqachon hisobga olingan.',
      msgForeignHint: '⛔ Boshqa mahsulot yoki jo\'natma. Ushbu PVZ ga tegishli emas — qaytarishga ajratib qo\'ying!',
      msgUndone: code => `↩️ «${code}» skani bekor qilindi.`,
      msgNothingToUndo: 'Bekor qilish uchun hech narsa yo\'q.',
      statusAcceptance: 'Joylashtirilgan / Qabulda',
      statusReturn: 'Yakunlangan (qaytarishga)',
      statusAcceptanceShort: 'Qabulda',
      statusOk: 'Mos',
      statusShort: 'Kamomad',
      statusMis: 'Joyida emas',
      statusBoth: 'Kamomad va joyida emas',
      notFoundShort: 'Topilmadi',
      errNoToken: 'Token topilmadi. "Buyurtmalar" yorlig\'ini yangilang (F5) va qaytadan urining.',
      errNoAccess: s => `WMS API ga kirish yo'q (${s}). Sahifani yangilang (F5) va qaytadan urining.`,
      errApi: s => `WMS API xatosi (${s}).`,
      errNoDpKey: 'PVZ ID kiritilmagan. Nuqta ID sini (dpKey) qo\'lda kiriting.',
      loadAuth: 'WMS ga avtorizatsiya...',
      loadList: k => `Buyurtmalar reestri yuklanmoqda (dpKey: ${k})...`,
      loadCollected: n => `Berishga yig'ilgan buyurtmalar: ${n}...`,
      loadItems: (a, b) => `Mahsulotlar tarkibi yuklanmoqda: ${a} / ${b}`,
      confirmFailed: n => `${n} ta buyurtma tarkibini yuklab bo'lmadi (ular uchun reja 1 dona deb olindi).\nInventarizatsiyani davom ettirasizmi?`,
      errLoadCancelled: 'Yuklash bekor qilindi. Qaytadan urining.',
      dpChecking: 'PVZ ID WMS da tekshirilmoqda...',
      dpEmpty: 'Nuqta ID (dpKey) aniqlanmadi. Qo\'lda kiriting va tugmani yana bosing.',
      dpFormat: 'Nuqta ID faqat lotin harf, raqam, «-» va «_» belgilaridan iborat bo\'lishi kerak.',
      dpInvalid: id => `ID ${id} WMS tomonidan tasdiqlanmadi. Qiymatni tekshiring yoki sahifani yangilang (F5).`,
      dpOk: id => `ID ${id} WMS tomonidan tasdiqlandi`,
      dpNotChecked: 'ID hali tekshirilmagan. «Qoldiqlarni yuklab boshlash» tugmasini bosing.',
      dpMismatch: (page, api) => `Sahifadagi «Berishga»: ${page} ta, WMS da bu ID bo'yicha: ${api} ta. ID noto'g'ri bo'lishi mumkin. Davom etasizmi?`,
      alreadySaved: 'Bu inventarizatsiya allaqachon saqlangan. Yangi inventarizatsiya uchun «Yangi inventarizatsiya» tugmasini bosing.',
      savedOk: pvz => `✓ ${pvz} inventarizatsiyasi bazaga saqlandi!\nAkt «Tarix» bo'limida mavjud.`,
      saveFail: '⚠️ Inventarizatsiyani brauzer bazasiga saqlab bo\'lmadi (joy yetarli emas). .xlsx akt yuklab olinadi.',
      confirmDelete: 'Bu inventarizatsiyani bazadan o\'chirasizmi?',
      histEmpty: 'Bazada hali saqlangan inventarizatsiyalar yo\'q. Tekshiruvdan keyin «Butunlay yakunlash» tugmasini bosing.',
      histNotFound: 'Hech narsa topilmadi.',
      histSearch: 'PID yoki shtrixkod bo\'yicha qidiruv',
      chipOrders: n => `Buyurtmalar ${n}`,
      chipPlan: n => `Reja ${n}`,
      chipFact: n => `Fakt ${n}`,
      chipShort: n => `Kamomad ${n}`,
      chipMis: n => `Joyida emas ${n}`,
      chipAcc: n => `Qabul ${n}`,
      chipRet: n => `Qaytarish ${n}`,
      chipPartner: n => `Hamkor ${n}`,
      chipTime: (tot, act) => `Vaqt ${tot} · ish ${act}`,
      btnXlsx: 'Yuklab olish .xlsx',
      btnSave: 'Butunlay yakunlash',
      btnSaved: 'Saqlangan ✓',
      btnExcel: 'Yuklab olish .xlsx',
      btnBack: 'Skanerlashga qaytish',
      deleteTitle: 'Bazadan o\'chirish',
      metaOrders: 'Buyurtmalar',
      metaPlan: 'Reja bo\'yicha',
      metaScanned: 'Skanerlangan',
      timeStart: 'Boshlangan',
      timeTotal: 'Umumiy vaqt',
      timeActive: 'Ish vaqti',
      timeSpeed: 'Tezlik',
      speedVal: n => `${n} dona/soat`,
      speedNA: '—',
      tileShort: 'Kamomad',
      tileMis: 'Joyida emas',
      tileAcc: 'Qabulda',
      tileRet: 'Qaytarish',
      tileSur: 'Begona',
      tilePartner: 'Hamkor',
      ofTotal: n => `${n} tadan`,
      issuesPill: n => `muammoli: ${n}`,
      secShort: 'Kamomad',
      secMis: 'Noto\'g\'ri javondagi buyurtmalar',
      secAcc: 'Qabulda (to\'liq qabul qilinmagan)',
      secRet: 'Bekor qilingan / qaytarishlar',
      secSur: 'Begona mahsulotlar / ro\'yxatda yo\'q',
      secPartner: 'Hamkor buyurtmalari (SX / RX)',
      pos: n => `${n} poz.`,
      noShort: 'Kamomad aniqlanmadi. Barcha mijoz buyurtmalari joyida!',
      noMis: 'Barcha buyurtmalar o\'z javonlarida.',
      noAcc: 'Javonlarda qabul holatidagi pozitsiyalar yo\'q.',
      noRet: 'Javonlarda bekor qilingan buyurtmalar topilmadi.',
      noSur: 'Omborda begona mahsulotlar topilmadi.',
      noPartner: 'Ushbu inventarizatsiyada hamkor buyurtmalari (SX / RX) yo\'q.',
      thPid: 'Buyurtma (PID)',
      thMissing: 'Yetishmaydi',
      thShelf: 'Reja bo\'yicha javon',
      thClient: 'Mijoz',
      thQty: 'Miqdor',
      thPlanCell: 'Reja bo\'yicha',
      thActualCell: 'Haqiqatda turibdi',
      thBarcodeLoc: 'Kod / Joy',
      thRec: 'Tavsiya',
      thBarcode: 'Kod',
      thExtra: 'Ortiqcha miqdor',
      thFoundAt: 'Qayerda topildi',
      thStatus: 'Holat',
      thPlanFact: 'Reja / Fakt',
      thLocation: 'Qayerda turibdi',
      recAcc: 'Omborga yubormang! Katakda qoldiring, qabulni o\'tkazing.',
      recRet: 'Javondan oling! Qaytarish omboriga jo\'nating.',
      recPS: 'Omborda mahsulot yo\'q — hamkor bilan tekshiring.',
      recPM: 'Rejadagi javonga qo\'ying.',
      recPB: 'Javonga qo\'ying va hamkor bilan kamomadni tekshiring.',
      recPO: 'Muammo yo\'q.',
      cellPrefix: 'katak ',
      unknown: 'Noma\'lum',
      mgrName: 'Xojiyev Suxrob',
      noName: '—',
      tgTitle: '📦 PVZ inventarizatsiyasi yakunlandi',
      tgPvz: (p, k) => `PVZ: ${p} (ID: ${k})`,
      tgDate: (d, tm) => `Sana: ${d} ${tm}`,
      tgAuditor: a => `Inventarizatsiya o'tkazuvchi: ${a}`,
      tgOrders: n => `Buyurtmalar: ${n}`,
      tgPlan: n => `Reja: ${n} dona`,
      tgFact: n => `Fakt: ${n} dona`,
      tgTime: (tot, act) => `⏱ Vaqt: ${tot} (ish: ${act})`,
      tgShort: (c, u) => `🔴 Kamomad: ${c} poz. (-${u} dona)`,
      tgMis: (c, u) => `🟡 Joyida emas: ${c} poz. (${u} dona)`,
      tgAcc: (c, u) => `⚠️ Qabulda: ${c} poz. (${u} dona)`,
      tgRet: (c, u) => `🔴 Qaytarishlar: ${c} poz. (${u} dona)`,
      tgSur: (c, u) => `⛔ Begona: ${c} poz. (+${u} dona)`,
      tgPartner: (c, u) => `🤝 Hamkor (SX/RX) muammoli: ${c} poz. (${u} dona)`,
      sheetName: 'Inventarizatsiya akti',
      xTitle: 'PVZ OMBOR INVENTARIZATSIYA AKTI',
      xPvz: 'Berish punkti:',
      xDate: 'Sana va vaqt:',
      xMgr: 'Hududiy boshqaruvchi:',
      xAuditor: 'Inventarizatsiya o\'tkazuvchi:',
      xStatus: 'Tekshiruv holati:',
      xStatusDone: 'To\'liq yakunlangan (bazaga saqlangan)',
      xSummary: '=== INVENTARIZATSIYA UMUMIY NATIJASI ===',
      xHdrInd: 'Ko\'rsatkich',
      xHdrQty: 'Miqdor',
      xHdrStatus: 'Holat',
      xSumOrders: 'WMS rejasidagi mijoz buyurtmalari:',
      xSumPlan: 'Reja bo\'yicha mahsulotlar:',
      xSumScanned: 'Jami skanerlangan mahsulotlar:',
      xSumShort: 'Sof kamomad (omborda topilmadi):',
      xSumMis: 'Noto\'g\'ri javondagi buyurtmalar:',
      xSumAcc: 'Qabuldagi buyurtmalar (to\'liq qabul qilinmagan):',
      xSumRet: 'Bekor qilingan / qaytarishlar (omborga jo\'natish):',
      xSumSur: 'Begona mahsulotlar (PVZ ga tegishli emas):',
      xSumPartner: 'Hamkor (SX/RX) buyurtmalari muammoli:',
      xSumTime: 'Sarflangan vaqt (umumiy / ish):',
      xStPlanToIssue: 'Berish rejasi',
      xStPlanPcs: 'Reja donalari',
      xStFact: 'Ombordagi fakt',
      xStShort: n => `Yetishmaydi: -${n} dona`,
      xStMis: n => `${n} dona joyini o'zgartirish kerak`,
      xStAcc: n => `${n} dona (katakda qoldiring, qabulni o'tkazing)`,
      xStRet: n => `${n} dona (zudlik bilan javondan olib omborga)`,
      xStSur: n => `+${n} dona qaytarish qutisiga`,
      xStPartner: n => `${n} dona muammoli`,
      xStTime: (act, s, e) => `Ish vaqti: ${act} · ${s} dan ${e} gacha`,
      xS1: 'BO\'LIM 1. KAMOMAD (OMBORDA YO\'Q BUYURTMALAR)',
      xS2: 'BO\'LIM 2. NOTO\'G\'RI JAVONDAGI BUYURTMALAR (OMBORDA, LEKIN BOSHQA JAVONDA)',
      xS3: 'BO\'LIM 3. QABULDAGI MAHSULOTLAR (TO\'LIQ QABUL QILINMAGAN — OMBORGA YUBORMANG!)',
      xS4: 'BO\'LIM 4. BEKOR QILINGAN BUYURTMALAR VA QAYTARISHLAR (ZUDLIK BILAN JAVONDAN OLING)',
      xS5: 'BO\'LIM 5. BEGONA MAHSULOTLAR VA ANIQLANMAGAN ORTIQCHALAR',
      xS6: 'BO\'LIM 6. HAMKOR BUYURTMALARI (SX / RX) — BARCHASI VA HOLATI',
      xOkShort: '✓ Kamomad aniqlanmadi. Barcha mijoz buyurtmalari joyida!',
      xOkMis: '✓ Topilgan buyurtmalar barchasi o\'z joyida.',
      xOkAcc: '✓ Tugallanmagan qabul holatidagi mahsulotlar topilmadi.',
      xOkRet: '✓ Javonlarda unutilgan bekor qilingan buyurtma yoki qaytarish topilmadi.',
      xOkSur: '✓ Omborda ortiqcha yoki begona jo\'natmalar topilmadi.',
      xOkPartner: 'Ushbu inventarizatsiyada hamkor buyurtmalari (SX / RX) yo\'q.',
      xHdrS1: ['№', 'Buyurtma raqami (PID)', 'WMS ID', 'Kod (SKU / Joy)', 'Yetishmaydi', 'Reja bo\'yicha javon', 'Qayerda topildi', 'Mijoz', 'Telefon'],
      xHdrS2: ['№', 'Buyurtma raqami (PID)', 'WMS ID', 'Kod (SKU / Joy)', 'Miqdor', 'Reja bo\'yicha javon', 'Haqiqatda qayerda', 'Mijoz', 'Harakat'],
      xHdrS3: ['№', 'O\'qilgan kod / Joy', 'Buyurtma ID', 'Miqdor', 'Javon', 'WMS holati', '', '', 'Tavsiya etilgan harakat'],
      xHdrS5: ['№', 'O\'qilgan kod', 'Miqdor', 'Qayerda topildi (javon)', 'Holat', '', '', '', 'Tavsiya etilgan harakat'],
      xHdrS6: ['№', 'Buyurtma (PID)', 'WMS ID', 'Holat', 'Reja (dona)', 'Fakt (dona)', 'Qayerda turibdi', 'Mijoz', 'Tavsiya'],
      xNotFound: 'TOPILMADI',
      xActS2: n => `${n} katakka o'tkazish`,
      xActS3: 'Katakda qoldiring! WMS da qabulni yakunlang',
      xActS4: 'Javondan oling! Qaytarish omboriga jo\'nating',
      xActS5: 'Qaytarish qutisiga ajratib qo\'ying',
      xStatusRetShort: 'Yakunlangan (jo\'natma)',
      xForeignStatus: 'Boshqa PVZ / ro\'yxatda yo\'q'
    }
  };

  let LANG = 'ru';
  let lastLocaleAt = 0;

  function t(key, ...args) {
    const dict = I18N[LANG] || I18N.ru;
    let v = dict[key];
    if (v === undefined) v = I18N.ru[key];
    if (v === undefined) return key;
    return typeof v === 'function' ? v(...args) : v;
  }

  // ---------------------------------------------------------------- storage / locale
  function readAllStorage() {
    const out = [];
    for (const s of [sessionStorage, localStorage]) {
      try {
        for (let i = 0; i < s.length; i++) {
          const key = s.key(i);
          const raw = s.getItem(key);
          if (raw == null) continue;
          let parsed;
          try { parsed = JSON.parse(raw); } catch (e) { parsed = undefined; }
          out.push({ key, raw, parsed });
        }
      } catch (e) {}
    }
    return out;
  }

  function findDeep(obj, key, depth = 0) {
    if (!obj || typeof obj !== 'object' || depth > 6) return undefined;
    if (Object.prototype.hasOwnProperty.call(obj, key)) return obj[key];
    for (const v of Object.values(obj)) {
      const r = findDeep(v, key, depth + 1);
      if (r !== undefined) return r;
    }
    return undefined;
  }

  function detectLocale() {
    const items = readAllStorage();
    const direct = items.find(i => i.key === 'locale');
    if (direct && typeof direct.parsed !== 'object') {
      const v = String(direct.parsed ?? direct.raw);
      return /^uz/i.test(v) ? 'uz' : 'ru';
    }
    for (const i of items) {
      const loc = findDeep(i.parsed, 'locale');
      if (typeof loc === 'string' && loc) return /^uz/i.test(loc) ? 'uz' : 'ru';
    }
    return 'ru';
  }

  // Storage'ni har safar o'qimaslik uchun 30 sekundda bir marta yangilanadi
  function refreshLocale(force) {
    const now = Date.now();
    if (!force && now - lastLocaleAt < 30000) return;
    lastLocaleAt = now;
    LANG = LANG_PREF === 'auto' ? detectLocale() : LANG_PREF;
  }

  function titleCase(s) {
    return String(s || '').toLowerCase().replace(/(^|[\s-])(\p{L})/gu, (m, p, c) => p + c.toUpperCase());
  }

  function getAuditorName() {
    let info = null;
    for (const i of readAllStorage()) {
      const u = findDeep(i.parsed, 'userInfo');
      if (u && typeof u === 'object' && (u.firstName || u.lastName)) { info = u; break; }
    }
    if (info) return titleCase(`${info.firstName || ''} ${info.lastName || ''}`.trim());
    return t('unknown');
  }

  // ---------------------------------------------------------------- utils
  function escHtml(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function maskPhone(p) {
    if (!p) return '';
    const s = String(p);
    if (s.length <= 8) return '***';
    return s.slice(0, 6) + '***' + s.slice(-2);
  }

  function safeCustomer(c) {
    c = c || {};
    return { firstName: c.firstName || '', lastName: c.lastName || '', phone: maskPhone(c.phone) };
  }

  function fullName(c) {
    return `${c?.firstName || ''} ${c?.lastName || ''}`.trim() || t('noName');
  }

  function sanitizeFileName(s) {
    return String(s).replace(/[\\/:*?"<>|]+/g, '_');
  }

  function $(id) { return document.getElementById(id); }

  const pad2 = n => String(n).padStart(2, '0');

  // ms -> HH:MM:SS
  function fmtDur(ms) {
    const s = Math.floor(Math.max(0, ms || 0) / 1000);
    return `${pad2(Math.floor(s / 3600))}:${pad2(Math.floor((s % 3600) / 60))}:${pad2(s % 60)}`;
  }

  // ms -> HH:MM (soat:daqiqa)
  function fmtClock(ms) {
    if (!ms) return '—';
    const d = new Date(ms);
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }

  function normOrderId(o) {
    const v = o?.b2cOrderId ?? o?.orderId ?? o?.id;
    return v == null ? '' : String(v);
  }

  function normPid(o) {
    return String(o?.b2cOrderPublicId ?? o?.b2cPublicOrderId ?? o?.orderNumber ?? '').trim();
  }

  function normBarcode(o) {
    return String(o?.b2cOrderBarcode ?? o?.orderBarcode ?? '').trim();
  }

  // PID yoki shtrixkod boshidagi SX / RX belgisi
  function partnerTagOf(...vals) {
    for (const v of vals) {
      const s = String(v == null ? '' : v).trim();
      const m = s.match(PARTNER_RE);
      if (m) return m[1].toUpperCase();
    }
    return '';
  }

  // UZUM BANK buyurtmasimi. SX/RX (hamkor) buyurtmalari bank hisoblanmaydi.
  function isBankOrder(o) {
    if (!o) return false;
    if (partnerTagOf(normPid(o), normBarcode(o))) return false;
    if (o.aerOrder === true) return true;
    return normBarcode(o).toUpperCase().startsWith('UZUM-BANK');
  }

  function isPartnerByType(...objs) {
    return objs.some(o => {
      if (!o || typeof o !== 'object') return false;
      if (o.partner === true || o.isPartner === true) return true;
      const typeStr = [o.orderType, o.type, o.deliveryType, o.sourceType]
        .filter(v => typeof v === 'string')
        .join(' ')
        .toUpperCase();
      return /PARTNER|HAMKOR/.test(typeStr);
    });
  }

  function detectKind(header, withItems, pidStr, barcodeStr) {
    if (isBankOrder(header)) return { kind: 'BANK', tag: '' };
    const tag = partnerTagOf(pidStr, barcodeStr);
    if (tag || isPartnerByType(header, withItems)) return { kind: 'PARTNER', tag };
    return { kind: 'ORDER', tag: '' };
  }

  function isItemable(h) {
    return !!h && h.b2cOrderId != null && !isBankOrder(h);
  }

  const CELL_KEYS = ['cellBarcode', 'cellKey', 'cell', 'cellInfo', 'cellCode', 'cellNumber', 'cellName'];
  function extractCell(...srcs) {
    for (const src of srcs) {
      if (!src || typeof src !== 'object') continue;
      for (const k of CELL_KEYS) {
        let v = src[k];
        if (v && typeof v === 'object') v = v.barcode || v.code || v.name || v.value || '';
        if (v != null && String(v).trim() !== '') return String(v).trim();
      }
    }
    return NO_CELL;
  }

  function isNoCell(c) {
    return !c || c === NO_CELL || c === 'null';
  }

  function cellLabel(c) {
    return isNoCell(c) ? t('unknown') : c;
  }

  function normCell(s) {
    return String(s).replace(/\s+/g, '').toUpperCase();
  }

  function kindTag(r) {
    if (!r) return '';
    if (r.kind === 'BANK') return t('badgeBank');
    if (r.kind === 'PARTNER') return r.tag ? `${t('badgePartner')} · ${r.tag}` : t('badgePartner');
    return '';
  }

  function pidText(r) {
    const base = r.pid || r.orderId;
    const k = kindTag(r);
    return k ? `${base} (${k})` : String(base);
  }

  function kindBadgeHtml(r) {
    if (!r || !r.kind || r.kind === 'ORDER') return '';
    const cls = r.kind === 'BANK' ? 's-info' : 's-par';
    return ` <span class="uz-tag ${cls}">${escHtml(kindTag(r))}</span>`;
  }

  function pidHtml(r) {
    return `<b>${escHtml(r.pid || r.orderId)}</b>${kindBadgeHtml(r)}`;
  }

  function statusText(st) {
    const map = { OK: t('statusOk'), SHORT: t('statusShort'), MIS: t('statusMis'), BOTH: t('statusBoth') };
    return map[st] || '';
  }

  function statusCls(st) {
    return { OK: 's-ok', SHORT: 's-bad', MIS: 's-warn', BOTH: 's-bad' }[st] || 's-mut';
  }

  function partnerAction(st) {
    const map = { OK: t('recPO'), SHORT: t('recPS'), MIS: t('recPM'), BOTH: t('recPB') };
    return map[st] || '';
  }

  function plainFound(list) {
    return list && list.length ? list.map(f => `${f.cell} (${f.qty})`).join(', ') : t('notFoundShort');
  }

  function safeCustomerName(c) {
    return escHtml(c?.firstName || '');
  }

  // ---------------------------------------------------------------- token
  function decodeJwt(token) {
    try {
      const base64Url = token.split('.')[1];
      const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      return JSON.parse(decodeURIComponent(atob(base64).split('').map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)).join('')));
    } catch (e) { return null; }
  }

  function isTokenValid(token) {
    if (!token || typeof token !== 'string') return false;
    const payload = decodeJwt(token);
    return !!(payload && payload.exp && (payload.exp * 1000 > Date.now() + 10000));
  }

  function findTokenAnywhere() {
    if (isTokenValid(window.__uzumToken)) return window.__uzumToken;
    try {
      const appEl = document.querySelector('#app') || document.querySelector('#root');
      const pinia = appEl?.__vue_app__?.config?.globalProperties?.$pinia;
      if (pinia) {
        const matches = JSON.stringify(pinia.state.value).match(jwtRegex);
        if (matches) for (const tk of matches) if (isTokenValid(tk)) return tk;
      }
    } catch (e) {}
    for (const s of [sessionStorage, localStorage]) {
      try {
        for (let i = 0; i < s.length; i++) {
          const matches = s.getItem(s.key(i))?.match(jwtRegex);
          if (matches) for (const tk of matches) if (isTokenValid(tk)) return tk;
        }
      } catch (e) {}
    }
    return null;
  }

  function captureAuth(value) {
    if (value && /Bearer/i.test(value)) {
      const raw = String(value).replace(/Bearer\s+/i, '').trim();
      if (isTokenValid(raw)) window.__uzumToken = raw;
    }
  }

  function captureDpKeyFromUrl(url) {
    if (url && url.includes('/delivery-point/')) {
      const m = url.match(/[?&]dpKey=([^&]+)/);
      if (m) window.__lastLiveDpKey = m[1];
    }
  }

  function readAuthFromHeaders(h) {
    if (!h) return null;
    try {
      if (typeof Headers !== 'undefined' && h instanceof Headers) return h.get('authorization');
      if (Array.isArray(h)) {
        const f = h.find(x => String(x[0]).toLowerCase() === 'authorization');
        return f ? f[1] : null;
      }
      return h.authorization || h.Authorization || null;
    } catch (e) { return null; }
  }

  if (!window.__fetchPatched) {
    const origFetch = window.fetch;
    window.fetch = function (...args) {
      try {
        const first = args[0];
        const urlStr = typeof first === 'string' ? first : (first?.url || '');
        captureDpKeyFromUrl(urlStr);
        let auth = readAuthFromHeaders(args[1]?.headers);
        if (!auth && first && typeof first === 'object') auth = readAuthFromHeaders(first.headers);
        captureAuth(auth);
      } catch (e) {}
      return origFetch.apply(this, args);
    };
    window.__fetchPatched = true;
  }

  if (!window.__xhrPatched) {
    const oOpen = XMLHttpRequest.prototype.open;
    const oSet = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
      try { captureDpKeyFromUrl(String(url)); } catch (e) {}
      return oOpen.call(this, method, url, ...rest);
    };
    XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
      try { if (/^authorization$/i.test(k)) captureAuth(v); } catch (e) {}
      return oSet.call(this, k, v);
    };
    window.__xhrPatched = true;
  }

  async function fetchJson(url, headers, timeoutMs = 20000, retries = 1) {
    let lastErr;
    for (let attempt = 0; attempt <= retries; attempt++) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), timeoutMs);
      try {
        const res = await fetch(url, { headers, signal: ctrl.signal });
        clearTimeout(timer);
        if (res.status === 401 || res.status === 403) {
          const err = new Error(t('errNoAccess', res.status));
          err.isAccessError = true;
          throw err;
        }
        if (!res.ok) throw new Error(t('errApi', res.status));
        return await res.json();
      } catch (e) {
        clearTimeout(timer);
        lastErr = e;
        if (e && e.isAccessError) throw e;
      }
    }
    throw lastErr;
  }

  function authHeaders() {
    const token = findTokenAnywhere();
    if (!token) throw new Error(t('errNoToken'));
    return { 'accept': 'application/json', 'authorization': `Bearer ${token}` };
  }

  // ---------------------------------------------------------------- sahifa sarlavhasi (PVZ va ID)
  function insideOurUi(el) {
    return !!(el.closest && (el.closest('#uzum-inv-modal') || el.closest('#uzum-inv-btn')));
  }

  function findPvzDropdown() {
    const all = Array.from(document.querySelectorAll('*'));
    for (const el of all) {
      if (insideOurUi(el)) continue;
      const r = el.getBoundingClientRect();
      if (r.top >= 0 && r.top <= 80 && r.right > window.innerWidth * 0.4 && r.width > 20 && r.height > 15) {
        const text = el.textContent?.trim() || '';
        const m = text.match(/^([A-Za-zА-Яа-яЁё]{2,8}-\d{1,4})/);
        if (m && !text.includes('build') && !text.includes('version') && text.length < 25) {
          let target = el;
          while (target.parentElement && target.parentElement.getBoundingClientRect().top <= 80) {
            const siblings = Array.from(target.parentElement.children);
            const hasSiblingButtons = siblings.some(s => s !== target && (s.textContent.includes('выдаче') || s.textContent.includes('Отменить') || s.tagName === 'BUTTON'));
            if (hasSiblingButtons) return target;
            if (target.parentElement.tagName === 'HEADER' || target.parentElement.tagName === 'BODY') break;
            target = target.parentElement;
          }
          return target;
        }
      }
    }
    return null;
  }

  function detectPvzNameFromHeader() {
    const dd = findPvzDropdown();
    if (dd) {
      const m = dd.textContent.trim().match(/\b([A-Za-zА-Яа-яЁё]{2,8}-\d{1,4})\b/);
      if (m) return m[1];
    }
    const allEls = Array.from(document.querySelectorAll('header *, nav *, div, button, span'));
    for (const el of allEls) {
      if (insideOurUi(el)) continue;
      const r = el.getBoundingClientRect();
      if (r.top >= 0 && r.top <= 80 && el.children.length <= 2) {
        const text = el.textContent?.trim() || '';
        const m = text.match(/\b([A-Za-zА-Яа-яЁё]{2,8}-\d{1,4})\b/);
        if (m && !text.includes('build') && !text.includes('version')) return m[1];
      }
    }
    return 'ПВЗ';
  }

  function detectLiveDpKey() {
    try {
      const resources = performance.getEntriesByType('resource');
      for (let i = resources.length - 1; i >= 0; i--) {
        const url = resources[i].name;
        if (url.includes('/delivery-point/') && url.includes('dpKey=')) {
          const m = url.match(/[?&]dpKey=([^&]+)/);
          if (m && m[1]) return m[1];
        }
      }
    } catch (e) {}
    if (window.__lastLiveDpKey) return window.__lastLiveDpKey;
    try {
      const appEl = document.querySelector('#app') || document.querySelector('#root');
      const pinia = appEl?.__vue_app__?.config?.globalProperties?.$pinia;
      if (pinia) {
        const str = JSON.stringify(pinia.state.value);
        const m = str.match(/"(?:dpKey|deliveryPointKey)":\s*"?(\d+)"?/);
        if (m && m[1]) return m[1];
      }
    } catch (e) {}
    return null;
  }

  // Sahifadagi "К выдаче N" sonini qat'iy o'qiydi (aniq topilmasa null)
  function readIssueCountStrict() {
    const all = Array.from(document.querySelectorAll('*'));
    for (const el of all) {
      if (insideOurUi(el)) continue;
      if (el.children.length > 2) continue;
      const m = (el.textContent || '').match(/К выдаче\s*(\d+)/i);
      if (m) return Number(m[1]);
    }
    return null;
  }

  function detectOrdersCount() {
    const c = readIssueCountStrict();
    return c == null ? '—' : String(c);
  }

  // ---------------------------------------------------------------- dpKey tekshiruvi
  // Yuklashdan oldin ID WMS'da tekshiriladi. Xato bo'lsa, qoldiqlar yuklanmaydi.
  async function validateDpKey(dpKey) {
    if (!dpKey) throw new Error(t('dpEmpty'));
    if (!DPKEY_RE.test(dpKey)) throw new Error(t('dpFormat'));

    const headers = authHeaders();
    let data;
    try {
      data = await fetchJson(
        `${API_V3}/orders?page=0&size=1&status=DELIVERED%2CTEMP_DELIVERED&deliveryPointKey=${encodeURIComponent(dpKey)}&sortField=accepted_date&direction=ASC`,
        headers, 15000, 1
      );
    } catch (e) {
      if (e && e.isAccessError) throw e;
      throw new Error(`${t('dpInvalid', dpKey)}\n${e && e.message ? e.message : ''}`.trim());
    }

    const isList = Array.isArray(data) || (data && Array.isArray(data.content));
    if (!isList) throw new Error(t('dpInvalid', dpKey));

    const total = !Array.isArray(data) && typeof data.totalElements === 'number' ? data.totalElements : null;
    return { total };
  }

  // ---------------------------------------------------------------- ish vaqti
  // Faol vaqt: ketma-ket harakatlar orasidagi vaqt hisoblanadi. Tanaffus
  // IDLE_LIMIT_MS dan uzoq bo'lsa, faqat shu chegarasi hisobga olinadi.
  const timing = { startedAt: 0, endedAt: 0, activeMs: 0, lastActivityAt: 0, running: false };

  function resetTiming() {
    timing.startedAt = 0;
    timing.endedAt = 0;
    timing.activeMs = 0;
    timing.lastActivityAt = 0;
    timing.running = false;
  }

  function startTiming() {
    const now = Date.now();
    timing.startedAt = now;
    timing.endedAt = 0;
    timing.activeMs = 0;
    timing.lastActivityAt = now;
    timing.running = true;
  }

  function markActivity() {
    if (!timing.running) return;
    const now = Date.now();
    timing.activeMs += Math.min(Math.max(0, now - timing.lastActivityAt), IDLE_LIMIT_MS);
    timing.lastActivityAt = now;
  }

  function finishTiming() {
    if (!timing.running) return;
    const now = Date.now();
    timing.activeMs += Math.min(Math.max(0, now - timing.lastActivityAt), IDLE_LIMIT_MS);
    timing.lastActivityAt = now;
    timing.endedAt = now;
    timing.running = false;
  }

  function timingSnapshot() {
    const now = Date.now();
    const endedAt = timing.running ? now : timing.endedAt;
    const tail = timing.running ? Math.min(Math.max(0, now - timing.lastActivityAt), IDLE_LIMIT_MS) : 0;
    return {
      startedAt: timing.startedAt,
      endedAt,
      totalMs: timing.startedAt && endedAt ? Math.max(0, endedAt - timing.startedAt) : 0,
      activeMs: timing.activeMs + tail
    };
  }

  function timeMetaHtml(scanned) {
    const tm = timingSnapshot();
    const hrs = tm.activeMs / 3600000;
    const speed = tm.activeMs >= 60000 && scanned > 0 ? t('speedVal', Math.round(scanned / hrs)) : t('speedNA');
    const items = [
      [t('timeStart'), fmtClock(tm.startedAt)],
      [t('timeTotal'), fmtDur(tm.totalMs)],
      [t('timeActive'), fmtDur(tm.activeMs)],
      [t('timeSpeed'), speed]
    ];
    return items.map(([k, v]) => `<span class="uz-chip">${escHtml(k)} <b>${escHtml(v)}</b></span>`).join('');
  }

  function updateTimerUI() {
    const a = $('u-timer');
    if (!a) return;
    const tm = timingSnapshot();
    a.textContent = fmtDur(tm.activeMs);
    const tot = $('u-timer-total');
    if (tot) tot.textContent = `${t('timeTotal')}: ${fmtDur(tm.totalMs)}`;
  }

  // ---------------------------------------------------------------- sound
  let audioCtx = null;
  function beep(freq = 800, duration = 0.08, type = 'sine') {
    if (!SOUND_ON) return;
    try {
      if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
      if (audioCtx.state === 'suspended') audioCtx.resume();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = type;
      osc.frequency.value = freq;
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      const t0 = audioCtx.currentTime;
      gain.gain.setValueAtTime(0.3, t0);
      gain.gain.exponentialRampToValueAtTime(0.00001, t0 + duration);
      osc.start(t0);
      osc.stop(t0 + duration);
    } catch (e) {}
  }

  const soundOk = () => beep(1046, 0.07);
  const soundCell = () => { beep(587, 0.08); setTimeout(() => beep(880, 0.08), 90); };
  const soundDuplicate = () => { beep(880, 0.05); setTimeout(() => beep(880, 0.05), 80); };
  const soundAcceptance = () => { beep(659, 0.08); setTimeout(() => beep(659, 0.08), 100); };
  const soundReturn = () => { beep(440, 0.12); setTimeout(() => beep(349, 0.15), 130); };
  const soundSurplus = () => beep(220, 0.25, 'sawtooth');
  const soundSuccess = () => {
    beep(523, 0.08);
    setTimeout(() => beep(659, 0.08), 90);
    setTimeout(() => beep(784, 0.08), 180);
    setTimeout(() => beep(1046, 0.16), 270);
  };

  // ---------------------------------------------------------------- state
  const state = {
    activePvzName: 'ПВЗ',
    dpKey: '',
    totalOrdersCount: 0,
    totalPlanUnits: 0,
    totalScannedUnits: 0,
    ordersMap: {},
    codeToOrderId: {},
    skuIndex: {},
    factGlobal: {},
    planCellUnits: {},
    factCellUnits: {},
    acceptanceScans: {},
    returnScans: {},
    foreignScans: {},
    scannedHistory: [],
    lookupCache: new Map(),
    activeCell: null,
    screen: 'start',
    workScreen: 'start',
    lastReport: null,
    savedInSession: false,
    loading: false
  };

  function resetInventoryState() {
    state.totalOrdersCount = 0;
    state.totalPlanUnits = 0;
    state.totalScannedUnits = 0;
    state.ordersMap = {};
    state.codeToOrderId = {};
    state.skuIndex = {};
    state.factGlobal = {};
    state.planCellUnits = {};
    state.factCellUnits = {};
    state.acceptanceScans = {};
    state.returnScans = {};
    state.foreignScans = {};
    state.scannedHistory = [];
    state.lookupCache = new Map();
    state.activeCell = null;
    state.lastReport = null;
    state.savedInSession = false;
  }

  // ---------------------------------------------------------------- deep lookup
  async function deepLookupShipmentStatus(code) {
    const cacheKey = code.trim();
    if (state.lookupCache.has(cacheKey)) return state.lookupCache.get(cacheKey);

    let headers;
    try { headers = authHeaders(); } catch (e) { return null; }
    const dpKey = state.dpKey;
    if (!dpKey) return null;

    const cleanCode = encodeURIComponent(cacheKey);
    const needle = cacheKey.toLowerCase();
    const candidateUrls = [
      `${API}/shipments/search?code=${cleanCode}&dpKey=${dpKey}`,
      `${API}/search-shipments?code=${cleanCode}&dpKey=${dpKey}`,
      `${API}/shipments/search?query=${cleanCode}&dpKey=${dpKey}`,
      `${API}/b2c-orders?search=${cleanCode}&dpKey=${dpKey}`
    ];

    let result = null;
    for (const url of candidateUrls) {
      try {
        const data = await fetchJson(url, headers, 8000, 0);
        const str = JSON.stringify(data).toLowerCase();
        if (!str.includes(needle)) continue;

        if (str.includes('приемк') || str.includes('acceptance') || str.includes('к размещению') || str.includes('размещен')) {
          const rawObj = Array.isArray(data) ? data[0] : (data.acceptance?.[0] || data.content?.[0] || data);
          result = {
            type: 'ACCEPTANCE',
            orderId: normOrderId(rawObj) || rawObj?.externalId || '',
            statusText: t('statusAcceptance')
          };
          break;
        }

        if (str.includes('отгрузк') || str.includes('завершен') || str.includes('возврат') || str.includes('return') || str.includes('cancel')) {
          const rawObj = Array.isArray(data) ? data[0] : (data.shipment?.[0] || data.returns?.[0] || data.content?.[0] || data);
          result = {
            type: 'RETURN',
            orderId: normOrderId(rawObj) || '',
            statusText: t('statusReturn')
          };
          break;
        }
      } catch (e) {
        if (e && e.isAccessError) break;
      }
    }

    state.lookupCache.set(cacheKey, result);
    return result;
  }

  // ---------------------------------------------------------------- load data
  async function loadInventoryData(onProgress) {
    onProgress(t('loadAuth'));
    const dpKey = state.dpKey;
    if (!dpKey) throw new Error(t('errNoDpKey'));
    authHeaders();

    // 1) Buyurtmalar ro'yxati (UZUM BANK va SX/RX ham shu yerda)
    let page = 0;
    const size = 100;
    const orderHeaders = [];
    const seenKeys = new Set();

    onProgress(t('loadList', dpKey));

    while (page < 500) {
      const data = await fetchJson(
        `${API_V3}/orders?page=${page}&size=${size}&status=DELIVERED%2CTEMP_DELIVERED&deliveryPointKey=${dpKey}&sortField=accepted_date&direction=ASC`,
        authHeaders(), 30000, 2
      );
      const content = Array.isArray(data) ? data : (data.content || []);
      content.forEach(o => {
        const key = normOrderId(o);
        if (!key || seenKeys.has(key)) return;
        seenKeys.add(key);
        orderHeaders.push(o);
      });

      let last;
      if (Array.isArray(data)) last = content.length < size;
      else if (typeof data.last === 'boolean') last = data.last;
      else last = content.length < size;

      page++;
      onProgress(t('loadCollected', orderHeaders.length));
      if (last || content.length === 0) break;
    }

    resetInventoryState();
    state.totalOrdersCount = orderHeaders.length;

    const headerMap = {};
    orderHeaders.forEach(h => { headerMap[normOrderId(h)] = h; });
    const orderIds = orderHeaders.map(h => normOrderId(h));

    // 2) Mahsulot tarkibi faqat oddiy buyurtmalar uchun so'raladi
    const itemIds = orderIds.filter(id => isItemable(headerMap[id]));
    const chunkSize = 25;
    const itemsMap = {};
    let failedOrders = 0;

    for (let i = 0; i < itemIds.length; i += chunkSize) {
      const slice = itemIds.slice(i, i + chunkSize);
      onProgress(t('loadItems', Math.min(i + chunkSize, itemIds.length), itemIds.length));
      try {
        const r = await fetchJson(`${API}/b2c-orders/items?b2cOrderId=${slice.join(',')}`, authHeaders(), 30000, 2);
        const arr = Array.isArray(r) ? r : (r.content || []);
        arr.forEach(o => {
          const id = normOrderId(o);
          if (id) itemsMap[id] = o;
        });
      } catch (e) {
        if (e && e.isAccessError) throw e;
        failedOrders += slice.length;
      }
    }

    // 3) Har bir buyurtmani yig'ish
    orderIds.forEach(orderIdStr => {
      const header = headerMap[orderIdStr] || {};
      const withItems = itemsMap[orderIdStr] || {};
      const isBank = isBankOrder(header);

      const pidStr = normPid(header) || normPid(withItems);
      const barcodeStr = normBarcode(header) || normBarcode(withItems);
      const pidShown = pidStr || (isBank ? barcodeStr : orderIdStr);
      const customer = header.customer || withItems.customer || {};
      const kinfo = detectKind(header, withItems, pidStr, barcodeStr);

      // Skaner o'qishi mumkin bo'lgan barcha variantlar
      const uniq = new Set();
      uniq.add(orderIdStr);
      if (!isBank) {
        uniq.add(orderIdStr.padStart(10, '0'));
        uniq.add(`10-${orderIdStr.padStart(10, '0')}-1`);
      }
      if (pidStr) uniq.add(pidStr);
      if (barcodeStr) {
        uniq.add(barcodeStr);
        const m = barcodeStr.match(/^10-(\d{10})-\d+$/);
        if (m) {
          uniq.add(m[1]);
          uniq.add(String(parseInt(m[1], 10)));
        }
      }

      const addCargo = (arr) => {
        if (!Array.isArray(arr)) return;
        arr.forEach(c => {
          const code = typeof c === 'object' ? c?.barcode : c;
          if (!code) return;
          const s = String(code).trim();
          uniq.add(s);
          const m = s.match(/^10-(\d{10})-\d+$/);
          if (m) uniq.add(String(parseInt(m[1], 10)));
        });
      };
      addCargo(header.cargoPlaceBarcodes);
      addCargo(header.cargoPlaces);
      addCargo(withItems.cargoPlaceBarcodes);
      addCargo(withItems.cargoPlaces);

      const skuCodes = new Set();
      const items = withItems.items || header.items || [];
      const cells = {};
      let itemsQty = 0;
      const itemSkus = [];

      items.forEach(it => {
        const qty = Number(it.amount) > 0 ? Number(it.amount) : 1;
        const cell = extractCell(it, withItems, header);
        itemsQty += qty;
        cells[cell] = (cells[cell] || 0) + qty;

        const sku = String(it.skuBarcode || it.barcode || '').trim();
        if (sku) {
          skuCodes.add(sku);
          itemSkus.push({ sku, qty, title: it.title || it.name || '' });
        }
        if (it.barcode) skuCodes.add(String(it.barcode).trim());
        if (it.sellerBarcode) skuCodes.add(String(it.sellerBarcode).trim());
        if (Array.isArray(it.barcodes)) it.barcodes.forEach(b => b && skuCodes.add(String(b).trim()));
      });

      // UZUM BANK yoki items bo'sh: reja 1 dona, katak buyurtma darajasidan
      if (itemsQty === 0) itemsQty = 1;
      if (Object.keys(cells).length === 0) {
        cells[extractCell(header, withItems)] = itemsQty;
      }

      const realCells = Object.keys(cells).filter(c => !isNoCell(c));

      state.ordersMap[orderIdStr] = {
        orderId: orderIdStr,
        pid: pidShown,
        kind: kinfo.kind,
        tag: kinfo.tag,
        cells,
        expectedCell: realCells.join(', ') || NO_CELL,
        customer,
        planQty: itemsQty,
        itemSkus,
        scannedCodes: new Set()
      };
      state.factGlobal[orderIdStr] = 0;

      uniq.forEach(a => {
        skuCodes.delete(a);
        state.codeToOrderId[a] = orderIdStr;
      });
      skuCodes.forEach(s => {
        if (!state.skuIndex[s]) state.skuIndex[s] = [];
        state.skuIndex[s].push(orderIdStr);
      });

      Object.entries(cells).forEach(([cell, qty]) => {
        if (!state.planCellUnits[cell]) state.planCellUnits[cell] = {};
        state.planCellUnits[cell][orderIdStr] = (state.planCellUnits[cell][orderIdStr] || 0) + qty;
      });
      state.totalPlanUnits += itemsQty;
    });

    Object.keys(state.planCellUnits).forEach(c => { state.factCellUnits[c] = {}; });

    if (failedOrders > 0) {
      const ok = confirm(t('confirmFailed', failedOrders));
      if (!ok) {
        resetInventoryState();
        throw new Error(t('errLoadCancelled'));
      }
    }
  }

  // ---------------------------------------------------------------- scanning
  function resolveBarcode(code, cell) {
    code = code.trim();
    if (!code) return null;

    if (state.codeToOrderId[code]) return { orderId: state.codeToOrderId[code], viaSku: false };

    const mCargo = code.match(/^10-(\d{10})-\d+$/);
    if (mCargo) {
      const extractedId = String(parseInt(mCargo[1], 10));
      if (state.ordersMap[extractedId]) return { orderId: extractedId, viaSku: false };
    }

    if (/^\d{8,12}$/.test(code)) {
      const cleanNum = String(parseInt(code, 10));
      if (state.ordersMap[cleanNum]) return { orderId: cleanNum, viaSku: false };
    }

    const candidates = state.skuIndex[code];
    if (candidates && candidates.length) {
      const remaining = candidates.filter(id => (state.factGlobal[id] || 0) < state.ordersMap[id].planQty);
      if (remaining.length === 0) return { orderId: candidates[0], viaSku: true, full: true };
      const onCell = remaining.find(id => state.ordersMap[id].cells[cell]);
      return { orderId: onCell || remaining[0], viaSku: true };
    }

    for (const [knownCode, orderId] of Object.entries(state.codeToOrderId)) {
      if (knownCode.length >= 8 && code.startsWith(knownCode)) {
        return { orderId, viaSku: false };
      }
    }
    return null;
  }

  function canonicalCell(code) {
    const k = Object.keys(state.planCellUnits).find(c => normCell(c) === normCell(code));
    return k || code;
  }

  // SX / RX kodlari katak sifatida tanilmaydi
  function looksLikeCell(code) {
    if (Object.keys(state.planCellUnits).some(c => normCell(c) === normCell(code))) return true;
    if (state.codeToOrderId[code] || state.skuIndex[code]) return false;
    if (PARTNER_RE.test(code)) return false;
    if (/^10-\d{10}-\d+$/.test(code)) return false;
    if (/^UZUM-BANK-/i.test(code)) return false;
    if (/^\d{8,}$/.test(code)) return false;
    return code.length <= 12;
  }

  let scanChain = Promise.resolve();
  function enqueueScan(code) {
    scanChain = scanChain.then(() => processScan(code)).catch(e => console.error('scan error', e));
  }

  async function processScan(raw) {
    const code = String(raw || '').trim();
    if (!code) return;

    if (looksLikeCell(code)) {
      state.activeCell = canonicalCell(code);
      if (!state.factCellUnits[state.activeCell]) state.factCellUnits[state.activeCell] = {};
      soundCell();
      updateScannerUI(code, 'CELL', t('hintScanItems'));
      return;
    }

    const cell = state.activeCell;
    if (!cell) {
      soundSurplus();
      updateScannerUI(code, 'INFO', t('msgNeedCell'));
      return;
    }

    const resolved = resolveBarcode(code, cell);
    if (resolved) {
      const { orderId, viaSku, full } = resolved;
      const order = state.ordersMap[orderId];
      const alreadyFull = full || (state.factGlobal[orderId] || 0) >= order.planQty;

      if ((!viaSku && order.scannedCodes.has(code)) || alreadyFull) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', t('msgAlreadyFull', order.pid || orderId));
        return;
      }

      if (!state.factCellUnits[cell]) state.factCellUnits[cell] = {};
      if (!viaSku) order.scannedCodes.add(code);
      state.factCellUnits[cell][orderId] = (state.factCellUnits[cell][orderId] || 0) + 1;
      state.factGlobal[orderId] = (state.factGlobal[orderId] || 0) + 1;
      state.totalScannedUnits++;
      state.scannedHistory.push({ cell, code, type: 'ORDER', orderId });

      soundOk();
      updateScannerUI(code, 'OK', t('msgAccepted', order.pid || orderId, order.customer?.firstName || ''));
      return;
    }

    updateScannerUI(code, 'SEARCH', t('msgSearching'));
    const deep = await deepLookupShipmentStatus(code);

    if (deep && deep.type === 'ACCEPTANCE') {
      if (!state.acceptanceScans[cell]) state.acceptanceScans[cell] = {};
      if (state.acceptanceScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', t('msgAcceptanceDup'));
        return;
      }
      state.acceptanceScans[cell][code] = { orderId: deep.orderId, statusText: deep.statusText, count: 1 };
      state.totalScannedUnits++;
      state.scannedHistory.push({ cell, code, type: 'ACCEPTANCE' });
      soundAcceptance();
      updateScannerUI(code, 'ACCEPTANCE', t('msgAcceptanceHint'));
      return;
    }

    if (deep && deep.type === 'RETURN') {
      if (!state.returnScans[cell]) state.returnScans[cell] = {};
      if (state.returnScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', t('msgReturnDup'));
        return;
      }
      state.returnScans[cell][code] = { orderId: deep.orderId, statusText: deep.statusText, count: 1 };
      state.totalScannedUnits++;
      state.scannedHistory.push({ cell, code, type: 'RETURN' });
      soundReturn();
      updateScannerUI(code, 'RETURN', t('msgReturnHint'));
      return;
    }

    if (!state.foreignScans[cell]) state.foreignScans[cell] = {};
    state.foreignScans[cell][code] = (state.foreignScans[cell][code] || 0) + 1;
    state.totalScannedUnits++;
    state.scannedHistory.push({ cell, code, type: 'FOREIGN' });
    soundSurplus();
    updateScannerUI(code, 'FOREIGN', t('msgForeignHint'));
  }

  function undoLastScan() {
    if (state.scannedHistory.length === 0) {
      updateScannerUI(null, 'INFO', t('msgNothingToUndo'));
      return;
    }
    const last = state.scannedHistory.pop();
    state.totalScannedUnits = Math.max(0, state.totalScannedUnits - 1);

    if (last.type === 'ORDER') {
      const order = state.ordersMap[last.orderId];
      if (order) order.scannedCodes.delete(last.code);
      state.factGlobal[last.orderId] = Math.max(0, (state.factGlobal[last.orderId] || 0) - 1);
      const cf = state.factCellUnits[last.cell];
      if (cf && cf[last.orderId]) {
        cf[last.orderId]--;
        if (cf[last.orderId] <= 0) delete cf[last.orderId];
      }
    } else if (last.type === 'ACCEPTANCE') {
      if (state.acceptanceScans[last.cell]) delete state.acceptanceScans[last.cell][last.code];
    } else if (last.type === 'RETURN') {
      if (state.returnScans[last.cell]) delete state.returnScans[last.cell][last.code];
    } else if (last.type === 'FOREIGN') {
      const f = state.foreignScans[last.cell];
      if (f && f[last.code]) {
        f[last.code]--;
        if (f[last.code] <= 0) delete f[last.code];
      }
    }
    updateScannerUI(null, 'INFO', t('msgUndone', last.code));
  }

  // ---------------------------------------------------------------- report
  function foundCellsOf(orderId) {
    const res = [];
    Object.keys(state.factCellUnits).forEach(cell => {
      const q = state.factCellUnits[cell][orderId] || 0;
      if (q > 0) res.push({ cell, qty: q });
    });
    return res;
  }

  function buildReport() {
    const report = {
      pureShortages: [],
      mismatchedCells: [],
      inAcceptance: [],
      canceledReturns: [],
      pureSurplus: [],
      partnerOrders: [],
      totalShortageUnits: 0,
      totalMismatchUnits: 0,
      totalAcceptanceUnits: 0,
      totalReturnUnits: 0,
      totalSurplusUnits: 0,
      totalPartnerUnits: 0
    };

    const skuOf = order => (order.itemSkus.length ? order.itemSkus[0].sku : `10-${order.orderId.padStart(10, '0')}-1`);

    Object.entries(state.ordersMap).forEach(([orderId, order]) => {
      const fact = state.factGlobal[orderId] || 0;
      const short = Math.max(0, order.planQty - fact);
      const found = foundCellsOf(orderId);
      const hasCells = !Object.keys(order.cells).every(isNoCell);

      const misByCell = [];
      if (hasCells) {
        found.forEach(f => {
          const extra = f.qty - (order.cells[f.cell] || 0);
          if (extra > 0) misByCell.push({ cell: f.cell, count: extra });
        });
      }
      const misQty = misByCell.reduce((a, m) => a + m.count, 0);

      const base = {
        orderId,
        pid: order.pid,
        kind: order.kind,
        tag: order.tag,
        sku: skuOf(order),
        expectedCell: order.expectedCell,
        customer: safeCustomer(order.customer)
      };

      // Hamkor (SX/RX) buyurtmalari alohida bo'limda
      if (order.kind === 'PARTNER') {
        let status = 'OK';
        if (short > 0 && misQty > 0) status = 'BOTH';
        else if (short > 0) status = 'SHORT';
        else if (misQty > 0) status = 'MIS';
        if (status !== 'OK') report.totalPartnerUnits += short + misQty;
        report.partnerOrders.push({ ...base, planQty: order.planQty, factQty: fact, short, misQty, status, found });
        return;
      }

      if (short > 0) {
        report.totalShortageUnits += short;
        report.pureShortages.push({ ...base, missingCount: short });
      }
      misByCell.forEach(m => {
        report.totalMismatchUnits += m.count;
        report.mismatchedCells.push({ ...base, count: m.count, actualCell: m.cell });
      });
    });

    const prio = { BOTH: 0, SHORT: 1, MIS: 2, OK: 3 };
    report.partnerOrders.sort((a, b) => (prio[a.status] - prio[b.status]) || String(a.pid).localeCompare(String(b.pid)));

    Object.entries(state.acceptanceScans).forEach(([cell, items]) => {
      Object.entries(items).forEach(([code, d]) => {
        report.totalAcceptanceUnits += d.count;
        report.inAcceptance.push({ code, orderId: d.orderId, foundCell: cell, count: d.count, statusText: d.statusText });
      });
    });

    Object.entries(state.returnScans).forEach(([cell, items]) => {
      Object.entries(items).forEach(([code, d]) => {
        report.totalReturnUnits += d.count;
        report.canceledReturns.push({ code, orderId: d.orderId, foundCell: cell, count: d.count, statusText: d.statusText });
      });
    });

    Object.entries(state.foreignScans).forEach(([cell, items]) => {
      Object.entries(items).forEach(([code, qty]) => {
        report.totalSurplusUnits += qty;
        report.pureSurplus.push({ code, qty, foundCell: cell });
      });
    });

    state.lastReport = report;
    return report;
  }

  // ---------------------------------------------------------------- telegram (fon)
  function buildTelegramText(entry) {
    const s = entry.stats || {};
    const tm = entry.timing || { totalMs: 0, activeMs: 0 };
    return [
      t('tgTitle'),
      t('tgPvz', entry.pvzName, entry.dpKey),
      t('tgDate', entry.date, entry.time),
      t('tgAuditor', entry.auditor),
      '',
      t('tgOrders', s.totalOrders ?? 0),
      t('tgPlan', s.totalPlanUnits ?? 0),
      t('tgFact', s.totalScanned ?? 0),
      t('tgTime', fmtDur(tm.totalMs), fmtDur(tm.activeMs)),
      '',
      t('tgShort', s.shortageCount ?? 0, s.shortageUnits ?? 0),
      t('tgMis', s.mismatchCount ?? 0, s.mismatchUnits ?? 0),
      t('tgAcc', s.acceptanceCount ?? 0, s.acceptanceUnits ?? 0),
      t('tgRet', s.returnCount ?? 0, s.returnUnits ?? 0),
      t('tgSur', s.surplusCount ?? 0, s.surplusUnits ?? 0),
      t('tgPartner', s.partnerCount ?? 0, s.partnerUnits ?? 0)
    ].join('\n');
  }

  async function sendTgTo(token, chatId, text) {
    const url = `https://api.telegram.org/bot${token}/sendMessage`;
    const body = JSON.stringify({
      chat_id: chatId,
      text: String(text).slice(0, 4000),
      disable_web_page_preview: true
    });
    for (let attempt = 0; attempt < 2; attempt++) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 15000);
      try {
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body,
          signal: ctrl.signal
        });
        clearTimeout(timer);
        if (res.ok) return true;
      } catch (e) {
        clearTimeout(timer);
        console.warn('[uzum-inv] telegram attempt failed', e && e.message);
      }
    }
    return false;
  }

  // Xabar barcha sozlangan chat ID larga yuboriladi. Hech bo'lmasa bittasiga yetsa — true.
  async function sendTelegramSilently(text) {
    const token = getTgToken();
    const chats = getTgChats();
    if (!TG_TOKEN_RE.test(token) || !chats.length) return false;
    const results = await Promise.all(chats.map(id => sendTgTo(token, id, text)));
    return results.some(Boolean);
  }

  // ---------------------------------------------------------------- xlsx (kutubxonasiz)
  function createRealXlsxBlob(sheets) {
    const encoder = new TextEncoder();
    const crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      crcTable[i] = c;
    }
    function getCrc(bytes) {
      let c = 0xFFFFFFFF;
      for (let i = 0; i < bytes.length; i++) c = (c >>> 8) ^ crcTable[(c ^ bytes[i]) & 0xFF];
      return (c ^ 0xFFFFFFFF) >>> 0;
    }
    function colName(n) {
      let s = '';
      while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
      }
      return s;
    }
    function esc(str) {
      if (str == null) return '';
      return String(str)
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    const files = [];
    let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;
    sheets.forEach((_, idx) => { ct += `<Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`; });
    ct += `</Types>`;
    files.push({ name: '[Content_Types].xml', data: ct });

    files.push({
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    });

    let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    sheets.forEach((_, idx) => { wbRels += `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`; });
    wbRels += `</Relationships>`;
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: wbRels });

    let wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`;
    sheets.forEach((s, idx) => { wb += `<sheet name="${esc(s.name).slice(0, 31)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`; });
    wb += `</sheets></workbook>`;
    files.push({ name: 'xl/workbook.xml', data: wb });

    sheets.forEach((s, idx) => {
      let ws = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n`;
      if (s.cols && Array.isArray(s.cols)) {
        ws += `<cols>\n`;
        s.cols.forEach((w, cIdx) => {
          ws += `<col min="${cIdx + 1}" max="${cIdx + 1}" width="${w}" customWidth="1"/>\n`;
        });
        ws += `</cols>\n`;
      }
      ws += `<sheetData>\n`;
      s.rows.forEach((row, rIdx) => {
        const rowNum = rIdx + 1;
        ws += `<row r="${rowNum}">`;
        row.forEach((val, cIdx) => {
          if (val === '' || val === null || val === undefined) return;
          const ref = colName(cIdx) + rowNum;
          if (typeof val === 'number' && isFinite(val)) {
            ws += `<c r="${ref}"><v>${val}</v></c>`;
          } else {
            ws += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
          }
        });
        ws += `</row>\n`;
      });
      ws += `</sheetData>\n</worksheet>`;
      files.push({ name: `xl/worksheets/sheet${idx + 1}.xml`, data: ws });
    });

    const DOS_TIME = 0;
    const DOS_DATE = 0x21;
    const fileEntries = [];
    let offset = 0;
    const parts = [];

    files.forEach(f => {
      const nameBytes = encoder.encode(f.name);
      const dataBytes = encoder.encode(f.data);
      const crc = getCrc(dataBytes);
      const size = dataBytes.length;

      const lh = new Uint8Array(30 + nameBytes.length);
      const dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0x0800, true);
      dv.setUint16(8, 0, true);
      dv.setUint16(10, DOS_TIME, true);
      dv.setUint16(12, DOS_DATE, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, size, true);
      dv.setUint32(22, size, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);

      parts.push(lh);
      parts.push(dataBytes);
      fileEntries.push({ nameBytes, crc, size, offset });
      offset += lh.length + size;
    });

    const centralDirOffset = offset;
    let centralDirSize = 0;

    fileEntries.forEach(fe => {
      const cdh = new Uint8Array(46 + fe.nameBytes.length);
      const dv = new DataView(cdh.buffer);
      dv.setUint32(0, 0x02014b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 20, true);
      dv.setUint16(8, 0x0800, true);
      dv.setUint16(10, 0, true);
      dv.setUint16(12, DOS_TIME, true);
      dv.setUint16(14, DOS_DATE, true);
      dv.setUint32(16, fe.crc, true);
      dv.setUint32(20, fe.size, true);
      dv.setUint32(24, fe.size, true);
      dv.setUint16(28, fe.nameBytes.length, true);
      dv.setUint16(30, 0, true);
      dv.setUint16(32, 0, true);
      dv.setUint16(34, 0, true);
      dv.setUint16(36, 0, true);
      dv.setUint32(38, 0, true);
      dv.setUint32(42, fe.offset, true);
      cdh.set(fe.nameBytes, 46);
      parts.push(cdh);
      centralDirSize += cdh.length;
    });

    const eocd = new Uint8Array(22);
    const dv = new DataView(eocd.buffer);
    dv.setUint32(0, 0x06054b50, true);
    dv.setUint16(4, 0, true);
    dv.setUint16(6, 0, true);
    dv.setUint16(8, fileEntries.length, true);
    dv.setUint16(10, fileEntries.length, true);
    dv.setUint32(12, centralDirSize, true);
    dv.setUint32(16, centralDirOffset, true);
    dv.setUint16(20, 0, true);
    parts.push(eocd);

    return new Blob(parts, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function exportCleanXlsx(reportData, auditMeta) {
    refreshLocale();
    const report = reportData || state.lastReport || buildReport();
    const m = auditMeta || {};
    const now = new Date();
    const dateStr = m.date || now.toLocaleDateString('ru-RU');
    const timeStr = m.time || now.toLocaleTimeString('ru-RU');
    const pvzName = m.pvzName || state.activePvzName;
    const dpKey = m.dpKey || state.dpKey;
    const auditor = m.auditor || getAuditorName();
    const totalOrders = m.totalOrders ?? state.totalOrdersCount;
    const totalPlanUnits = m.totalPlanUnits ?? state.totalPlanUnits;
    const totalScanned = m.totalScanned ?? state.totalScannedUnits;
    const partners = report.partnerOrders || [];
    const partnerIssues = partners.filter(p => p.status !== 'OK');

    // Vaqt: tarixdan olinsa o'sha yozuv, joriy ko'rinishda hozirgi sessiya
    const tm = auditMeta ? (auditMeta.timing || null) : (timing.startedAt ? timingSnapshot() : null);

    const rows = [];
    rows.push([t('xTitle')]);
    rows.push([t('xPvz'), `${pvzName} (ID: ${dpKey})`, '', t('xDate'), `${dateStr} ${timeStr}`]);
    rows.push([t('xMgr'), t('mgrName'), '', t('xStatus'), t('xStatusDone')]);
    rows.push([t('xAuditor'), auditor]);
    rows.push([]);

    rows.push([t('xSummary')]);
    rows.push([t('xHdrInd'), t('xHdrQty'), t('xHdrStatus')]);
    rows.push([t('xSumOrders'), totalOrders, t('xStPlanToIssue')]);
    rows.push([t('xSumPlan'), totalPlanUnits, t('xStPlanPcs')]);
    rows.push([t('xSumScanned'), totalScanned, t('xStFact')]);
    rows.push([t('xSumShort'), report.pureShortages.length, t('xStShort', report.totalShortageUnits)]);
    rows.push([t('xSumMis'), report.mismatchedCells.length, t('xStMis', report.totalMismatchUnits)]);
    rows.push([t('xSumAcc'), report.inAcceptance.length, t('xStAcc', report.totalAcceptanceUnits)]);
    rows.push([t('xSumRet'), report.canceledReturns.length, t('xStRet', report.totalReturnUnits)]);
    rows.push([t('xSumSur'), report.pureSurplus.length, t('xStSur', report.totalSurplusUnits)]);
    rows.push([t('xSumPartner'), partnerIssues.length, t('xStPartner', report.totalPartnerUnits || 0)]);
    rows.push([
      t('xSumTime'),
      tm ? fmtDur(tm.totalMs) : t('noName'),
      tm ? t('xStTime', fmtDur(tm.activeMs), fmtClock(tm.startedAt), fmtClock(tm.endedAt)) : t('noName')
    ]);
    rows.push([]);

    rows.push([t('xS1')]);
    if (report.pureShortages.length === 0) {
      rows.push([t('xOkShort')]);
    } else {
      rows.push(t('xHdrS1'));
      report.pureShortages.forEach((r, idx) => {
        rows.push([idx + 1, pidText(r), r.orderId, r.sku, `-${t('pcs', r.missingCount)}`, cellLabel(r.expectedCell), t('xNotFound'), fullName(r.customer), r.customer?.phone || t('noName')]);
      });
    }
    rows.push([]);

    rows.push([t('xS2')]);
    if (report.mismatchedCells.length === 0) {
      rows.push([t('xOkMis')]);
    } else {
      rows.push(t('xHdrS2'));
      report.mismatchedCells.forEach((r, idx) => {
        rows.push([idx + 1, pidText(r), r.orderId, r.sku, t('pcs', r.count), cellLabel(r.expectedCell), r.actualCell, fullName(r.customer), t('xActS2', r.actualCell)]);
      });
    }
    rows.push([]);

    rows.push([t('xS3')]);
    if (report.inAcceptance.length === 0) {
      rows.push([t('xOkAcc')]);
    } else {
      rows.push(t('xHdrS3'));
      report.inAcceptance.forEach((r, idx) => {
        rows.push([idx + 1, r.code, r.orderId || t('noName'), `+${t('pcs', r.count)}`, `${t('cellPrefix')}${r.foundCell}`, r.statusText || t('statusAcceptanceShort'), '', '', t('xActS3')]);
      });
    }
    rows.push([]);

    rows.push([t('xS4')]);
    if (report.canceledReturns.length === 0) {
      rows.push([t('xOkRet')]);
    } else {
      rows.push(t('xHdrS3'));
      report.canceledReturns.forEach((r, idx) => {
        rows.push([idx + 1, r.code, r.orderId || t('noName'), `+${t('pcs', r.count)}`, `${t('cellPrefix')}${r.foundCell}`, r.statusText || t('xStatusRetShort'), '', '', t('xActS4')]);
      });
    }
    rows.push([]);

    rows.push([t('xS5')]);
    if (report.pureSurplus.length === 0) {
      rows.push([t('xOkSur')]);
    } else {
      rows.push(t('xHdrS5'));
      report.pureSurplus.forEach((r, idx) => {
        rows.push([idx + 1, r.code, `+${t('pcs', r.qty)}`, `${t('cellPrefix')}${r.foundCell}`, t('xForeignStatus'), '', '', '', t('xActS5')]);
      });
    }
    rows.push([]);

    rows.push([t('xS6')]);
    if (partners.length === 0) {
      rows.push([t('xOkPartner')]);
    } else {
      rows.push(t('xHdrS6'));
      partners.forEach((p, idx) => {
        rows.push([idx + 1, pidText(p), p.orderId, statusText(p.status), p.planQty, p.factQty, plainFound(p.found), fullName(p.customer), partnerAction(p.status)]);
      });
    }

    const colWidths = [6, 30, 14, 24, 16, 18, 24, 28, 44];
    const blob = createRealXlsxBlob([{ name: t('sheetName'), rows, cols: colWidths }]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeFileName(`${t('sheetName')}_${pvzName}_${String(dateStr).replace(/\./g, '-')}.xlsx`);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ---------------------------------------------------------------- history
  function getHistoryList() {
    try {
      const raw = localStorage.getItem(HISTORY_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) { return []; }
  }

  function writeHistory(list) {
    for (let attempt = 0; attempt < 10; attempt++) {
      try {
        localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
        return true;
      } catch (e) {
        if (list.length <= 1) return false;
        list.pop();
      }
    }
    return false;
  }

  function saveAuditPermanently() {
    refreshLocale(true);
    if (state.savedInSession) {
      alert(t('alreadySaved'));
      return;
    }

    const report = buildReport();
    const now = new Date();
    const partnerIssues = report.partnerOrders.filter(p => p.status !== 'OK');
    const tm = timingSnapshot();

    const auditEntry = {
      id: `INV-${Date.now()}`,
      date: now.toLocaleDateString('ru-RU'),
      time: now.toLocaleTimeString('ru-RU'),
      pvzName: state.activePvzName,
      dpKey: state.dpKey,
      auditor: getAuditorName(),
      timing: {
        startedAt: tm.startedAt,
        endedAt: tm.endedAt,
        totalMs: tm.totalMs,
        activeMs: tm.activeMs
      },
      stats: {
        totalOrders: state.totalOrdersCount,
        totalPlanUnits: state.totalPlanUnits,
        totalScanned: state.totalScannedUnits,
        shortageCount: report.pureShortages.length,
        shortageUnits: report.totalShortageUnits,
        mismatchCount: report.mismatchedCells.length,
        mismatchUnits: report.totalMismatchUnits,
        acceptanceCount: report.inAcceptance.length,
        acceptanceUnits: report.totalAcceptanceUnits,
        returnCount: report.canceledReturns.length,
        returnUnits: report.totalReturnUnits,
        surplusCount: report.pureSurplus.length,
        surplusUnits: report.totalSurplusUnits,
        partnerCount: partnerIssues.length,
        partnerUnits: report.totalPartnerUnits
      },
      report
    };

    const list = getHistoryList();
    list.unshift(auditEntry);
    while (list.length > HISTORY_LIMIT) list.pop();
    const saved = writeHistory(list);

    // Telegram va vaqt tugashi faqat muvaffaqiyatli saqlangandan keyin
    if (saved) {
      state.savedInSession = true;
      finishTiming();
      sendTelegramSilently(buildTelegramText(auditEntry)).then(ok => {
        if (!ok) console.warn('[uzum-inv] telegram: xabar yuborilmadi');
      });
    }

    soundSuccess();
    exportCleanXlsx(report, {
      date: auditEntry.date,
      time: auditEntry.time,
      pvzName: auditEntry.pvzName,
      dpKey: auditEntry.dpKey,
      auditor: auditEntry.auditor,
      totalOrders: auditEntry.stats.totalOrders,
      totalPlanUnits: auditEntry.stats.totalPlanUnits,
      totalScanned: auditEntry.stats.totalScanned,
      timing: auditEntry.timing
    });
    alert(saved ? t('savedOk', state.activePvzName) : t('saveFail'));
    renderHeadActions('report');
  }

  // ---------------------------------------------------------------- CSS
  const CSS = `
#uzum-inv-btn{position:fixed;top:12px;left:0;z-index:999999;display:inline-flex;align-items:center;gap:8px;height:38px;padding:0 16px;margin:0;border:0;border-radius:12px;cursor:pointer;color:#fff;font:700 13px/1 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;white-space:nowrap;background:linear-gradient(135deg,#a78bfa 0%,#7c3aed 100%);box-shadow:0 8px 20px rgba(124,58,237,.35),inset 0 1px 0 rgba(255,255,255,.2);transition:transform .15s ease,box-shadow .2s ease,filter .2s ease;outline:none}
#uzum-inv-btn:hover{transform:translateY(-1px);filter:brightness(1.06);box-shadow:0 12px 26px rgba(124,58,237,.45),inset 0 1px 0 rgba(255,255,255,.2)}
#uzum-inv-btn:active{transform:translateY(0)}
#uzum-inv-btn:focus-visible{box-shadow:0 0 0 3px rgba(167,139,250,.5)}

#uzum-inv-modal{--bg:#f5f3fc;--side:#ffffff;--surface:#ffffff;--surface-2:#f2eefb;--surface-3:#e9e2f8;--line:#e6dff4;--line-2:#d5c9ea;--text:#1d1538;--muted:#6e6792;--primary:#7c3aed;--primary-2:#a855f7;--primary-ink:#ffffff;--ok:#059669;--bad:#e11d48;--warn:#d97706;--acc:#0284c7;--ret:#ea580c;--info:#4f46e5;--par:#c026d3;--mut:#64748b;--shadow:0 30px 80px rgba(46,16,101,.22);display:none;position:fixed;inset:0;z-index:1000000;align-items:center;justify-content:center;padding:20px;background:rgba(30,20,60,.45);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);font-family:system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;font-size:14px;line-height:1.45;color:var(--text)}
@media (prefers-color-scheme:dark){#uzum-inv-modal{--bg:#0f0b1d;--side:#171130;--surface:#1c1538;--surface-2:#241b47;--surface-3:#2f2459;--line:#33295a;--line-2:#463a74;--text:#f3efff;--muted:#a79cd0;--primary:#a78bfa;--primary-2:#c084fc;--primary-ink:#1b1038;--ok:#34d399;--bad:#fb7185;--warn:#fbbf24;--acc:#38bdf8;--ret:#fb923c;--info:#818cf8;--par:#e879f9;--mut:#94a3b8;--shadow:0 30px 80px rgba(0,0,0,.6)}}
#uzum-inv-modal *{box-sizing:border-box}
#uzum-inv-modal button,#uzum-inv-modal input{font:inherit}
#uzum-inv-modal .s-ok{--c:var(--ok);--cs:color-mix(in srgb,var(--ok) 14%,transparent)}
#uzum-inv-modal .s-bad{--c:var(--bad);--cs:color-mix(in srgb,var(--bad) 14%,transparent)}
#uzum-inv-modal .s-warn{--c:var(--warn);--cs:color-mix(in srgb,var(--warn) 14%,transparent)}
#uzum-inv-modal .s-acc{--c:var(--acc);--cs:color-mix(in srgb,var(--acc) 14%,transparent)}
#uzum-inv-modal .s-ret{--c:var(--ret);--cs:color-mix(in srgb,var(--ret) 14%,transparent)}
#uzum-inv-modal .s-info{--c:var(--info);--cs:color-mix(in srgb,var(--info) 14%,transparent)}
#uzum-inv-modal .s-par{--c:var(--par);--cs:color-mix(in srgb,var(--par) 14%,transparent)}
#uzum-inv-modal .s-mut{--c:var(--mut);--cs:color-mix(in srgb,var(--mut) 14%,transparent)}
#uzum-inv-modal .s-search{--c:var(--primary);--cs:color-mix(in srgb,var(--primary) 14%,transparent)}
#uzum-inv-modal .uz-c{color:var(--c)}

#uzum-inv-modal .uz-app{width:min(1120px,100%);height:min(92vh,900px);display:grid;grid-template-columns:232px minmax(0,1fr);background:var(--bg);border:1px solid var(--line);border-radius:22px;overflow:hidden;box-shadow:var(--shadow)}
#uzum-inv-modal .uz-side{display:flex;flex-direction:column;gap:18px;padding:20px 16px;background:var(--side);border-right:1px solid var(--line);min-height:0;overflow:auto}
#uzum-inv-modal .uz-brand{display:flex;align-items:center;gap:12px;padding:4px 6px}
#uzum-inv-modal .uz-logo{width:42px;height:42px;flex-shrink:0;border-radius:14px;display:grid;place-items:center;font-weight:900;font-size:19px;color:#fff;background:linear-gradient(135deg,var(--primary-2),var(--primary));box-shadow:0 10px 24px color-mix(in srgb,var(--primary) 35%,transparent)}
#uzum-inv-modal .uz-brand-t{font-weight:800;font-size:15px;letter-spacing:-.2px}
#uzum-inv-modal .uz-brand-s{font-size:11.5px;color:var(--muted);margin-top:2px}
#uzum-inv-modal .uz-nav{display:flex;flex-direction:column;gap:6px}
#uzum-inv-modal .uz-navbtn{display:flex;align-items:center;gap:10px;width:100%;padding:11px 12px;border:0;border-radius:12px;background:transparent;color:var(--muted);font-weight:700;font-size:13.5px;text-align:left;cursor:pointer;transition:background .15s,color .15s}
#uzum-inv-modal .uz-navbtn:hover{background:var(--surface-2);color:var(--text)}
#uzum-inv-modal .uz-navbtn.active{background:var(--primary);color:var(--primary-ink);box-shadow:0 8px 20px color-mix(in srgb,var(--primary) 30%,transparent)}
#uzum-inv-modal .uz-navic{width:18px;text-align:center;font-size:15px}
#uzum-inv-modal .uz-navbottom{margin-top:auto;display:flex;flex-direction:column;gap:6px;padding-top:12px;border-top:1px solid var(--line)}
#uzum-inv-modal .uz-pvzcard{margin-top:6px;padding:12px 14px;border-radius:14px;background:var(--surface-2);border:1px solid var(--line)}
#uzum-inv-modal .uz-pvzcard .uz-lbl{display:block;margin-bottom:4px}
#uzum-inv-modal .uz-pvzcard b{display:block;font-size:14px;font-weight:800}

#uzum-inv-modal .uz-main{display:flex;flex-direction:column;min-width:0;min-height:0}
#uzum-inv-modal .uz-head{display:flex;align-items:flex-start;flex-wrap:wrap;gap:12px 14px;padding:22px 26px 12px;flex-shrink:0}
#uzum-inv-modal .uz-head-l{flex:1;min-width:0}
#uzum-inv-modal .uz-h1{margin:0;font-size:21px;font-weight:800;letter-spacing:-.4px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#uzum-inv-modal .uz-sub{margin:4px 0 0;font-size:13px;color:var(--muted)}
#uzum-inv-modal .uz-head-a{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}
#uzum-inv-modal .uz-x{width:36px;height:36px;flex-shrink:0;display:grid;place-items:center;border-radius:11px;border:1px solid var(--line);background:var(--surface);color:var(--muted);cursor:pointer;font-size:13px;transition:all .15s ease}
#uzum-inv-modal .uz-x:hover{color:var(--text);border-color:var(--line-2)}
#uzum-inv-modal .uz-steps{display:flex;gap:8px;flex-wrap:wrap;padding:0 26px 14px;flex-shrink:0}
#uzum-inv-modal .uz-step{display:flex;align-items:center;gap:8px;padding:5px 12px 5px 5px;border-radius:999px;background:var(--surface);border:1px solid var(--line);font-size:12.5px;font-weight:700;color:var(--muted)}
#uzum-inv-modal .uz-dot{width:24px;height:24px;border-radius:50%;display:grid;place-items:center;font-size:11.5px;background:var(--surface-3);color:var(--muted)}
#uzum-inv-modal .uz-step.done{color:var(--text)}
#uzum-inv-modal .uz-step.done .uz-dot{background:var(--ok);color:#fff}
#uzum-inv-modal .uz-step.on{color:var(--primary);border-color:color-mix(in srgb,var(--primary) 45%,transparent)}
#uzum-inv-modal .uz-step.on .uz-dot{background:var(--primary);color:var(--primary-ink)}
#uzum-inv-modal .uz-body{flex:1;min-height:0;overflow-y:auto;padding:4px 26px 26px;-webkit-overflow-scrolling:touch}

#uzum-inv-modal .uz-btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:40px;padding:0 16px;border-radius:12px;border:1px solid var(--line-2);background:var(--surface);color:var(--text);font-weight:700;font-size:13px;cursor:pointer;white-space:nowrap;transition:transform .12s,filter .15s,border-color .15s,color .15s}
#uzum-inv-modal .uz-btn:hover{border-color:var(--primary);color:var(--primary)}
#uzum-inv-modal .uz-btn:active{transform:translateY(1px)}
#uzum-inv-modal .uz-btn:focus-visible{outline:2px solid var(--primary);outline-offset:2px}
#uzum-inv-modal .uz-btn:disabled{opacity:.55;cursor:default;transform:none}
#uzum-inv-modal .uz-btn.pri{background:linear-gradient(135deg,var(--primary-2),var(--primary));color:#fff;border-color:transparent;box-shadow:0 10px 22px color-mix(in srgb,var(--primary) 30%,transparent)}
#uzum-inv-modal .uz-btn.pri:hover{color:#fff;filter:brightness(1.07)}
#uzum-inv-modal .uz-btn.ok{background:linear-gradient(135deg,#34d399,var(--ok));color:#fff;border-color:transparent}
#uzum-inv-modal .uz-btn.ok:hover{color:#fff;filter:brightness(1.05)}
#uzum-inv-modal .uz-btn.ghost{background:transparent}
#uzum-inv-modal .uz-btn.danger{background:color-mix(in srgb,var(--bad) 10%,transparent);color:var(--bad);border-color:transparent}
#uzum-inv-modal .uz-btn.danger:hover{color:var(--bad);filter:brightness(1.05)}
#uzum-inv-modal .uz-btn.big{width:100%;min-height:52px;font-size:15px;border-radius:14px}

#uzum-inv-modal .uz-card{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:18px 20px;min-width:0}
#uzum-inv-modal .uz-lbl{font-size:11px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted)}
#uzum-inv-modal .uz-muted{color:var(--muted);font-size:12.5px}
#uzum-inv-modal .uz-empty{padding:26px;text-align:center;color:var(--muted);background:var(--surface);border:1px dashed var(--line-2);border-radius:16px}

#uzum-inv-modal .uz-setup{max-width:580px;margin:6px auto 0;text-align:center}
#uzum-inv-modal .uz-hero{width:68px;height:68px;margin:6px auto 16px;border-radius:22px;display:grid;place-items:center;font-size:30px;background:linear-gradient(135deg,color-mix(in srgb,var(--primary-2) 25%,transparent),color-mix(in srgb,var(--primary) 18%,transparent))}
#uzum-inv-modal .uz-form{display:grid;grid-template-columns:1fr 1fr;gap:14px;text-align:left}
#uzum-inv-modal .uz-field{display:flex;flex-direction:column;gap:8px;min-width:0}
#uzum-inv-modal .uz-in{width:100%;min-height:46px;padding:0 14px;border-radius:12px;border:1px solid var(--line-2);background:var(--surface-2);color:var(--text);font-size:15px;font-weight:700;outline:none;transition:border-color .15s,box-shadow .15s}
#uzum-inv-modal .uz-in:focus{border-color:var(--primary);box-shadow:0 0 0 4px color-mix(in srgb,var(--primary) 18%,transparent)}
#uzum-inv-modal .uz-issue{margin-top:16px;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 16px;border-radius:14px;background:var(--surface-2);border:1px dashed var(--line-2)}
#uzum-inv-modal .uz-issue b{font-size:20px;font-weight:900;color:var(--primary)}
#uzum-inv-modal .uz-note{margin-top:8px;font-size:12px;color:var(--muted)}
#uzum-inv-modal .uz-dpst{margin-top:14px;padding:10px 14px;border-radius:12px;font-weight:700;font-size:13px;text-align:left;color:var(--c,var(--muted));background:var(--cs,var(--surface-2));border:1px solid color-mix(in srgb,var(--c,var(--muted)) 30%,transparent);border-left:4px solid var(--c,var(--muted))}
#uzum-inv-modal .uz-prog{min-height:20px;margin-top:12px;text-align:center;font-size:13px;font-weight:700;color:var(--primary)}

#uzum-inv-modal .uz-stats{display:grid;grid-template-columns:repeat(3,minmax(0,1fr)) 112px;gap:12px;margin-bottom:12px}
#uzum-inv-modal .uz-stat{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:14px 16px;min-width:0}
#uzum-inv-modal .uz-stat b{display:block;margin-top:6px;font-size:24px;font-weight:900;letter-spacing:-.5px;font-variant-numeric:tabular-nums;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#uzum-inv-modal .uz-stat.hl b{color:var(--primary)}
#uzum-inv-modal .uz-stat-ring{display:flex;align-items:center;justify-content:center;padding:10px}
#uzum-inv-modal .uz-ring{position:relative;width:76px;height:76px}
#uzum-inv-modal .uz-ring svg{width:100%;height:100%;transform:rotate(-90deg)}
#uzum-inv-modal .uz-ring circle{fill:none;stroke-width:7}
#uzum-inv-modal .uz-ring .bg{stroke:var(--surface-3)}
#uzum-inv-modal .uz-ring .fg{stroke:var(--primary);stroke-linecap:round;transition:stroke-dashoffset .4s ease}
#uzum-inv-modal .uz-ring-t{position:absolute;inset:0;display:grid;place-items:center;font-weight:900;font-size:15px;font-variant-numeric:tabular-nums}
#uzum-inv-modal .uz-meter{height:8px;border-radius:999px;background:var(--surface-3);overflow:hidden}
#uzum-inv-modal .uz-meter>div{height:100%;width:0;border-radius:999px;background:linear-gradient(90deg,var(--primary-2),var(--primary));transition:width .35s ease}
#uzum-inv-modal .uz-barrow{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px 14px;margin:8px 2px 16px}
#uzum-inv-modal .uz-barlbl{font-size:12px;font-weight:700;color:var(--muted);font-variant-numeric:tabular-nums}
#uzum-inv-modal .uz-timer{display:inline-flex;align-items:center;flex-wrap:wrap;gap:8px;font-size:12px;font-weight:700;color:var(--muted)}
#uzum-inv-modal .uz-timer b{font-size:18px;font-weight:900;color:var(--primary);font-variant-numeric:tabular-nums}

#uzum-inv-modal .uz-work{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.3fr);gap:12px;margin-bottom:14px}
#uzum-inv-modal .uz-cellcard{display:flex;flex-direction:column;justify-content:center;min-height:150px;border-radius:18px;padding:18px 20px;color:#fff;background:linear-gradient(150deg,#8b5cf6 0%,#5b21b6 100%);box-shadow:0 14px 30px rgba(91,33,182,.3);min-width:0}
#uzum-inv-modal .uz-cellcard .uz-lbl{color:rgba(255,255,255,.78)}
#uzum-inv-modal .uz-cell{margin-top:8px;font-size:clamp(28px,4vw,42px);font-weight:900;letter-spacing:-.5px;line-height:1.05;word-break:break-all}
#uzum-inv-modal .uz-cell.empty{font-size:15px;font-weight:700;letter-spacing:0;opacity:.85}
#uzum-inv-modal .uz-lastcard{display:flex;flex-direction:column;gap:10px;justify-content:center}
#uzum-inv-modal .uz-code{font:700 15px/1.3 ui-monospace,Consolas,monospace;word-break:break-all}
#uzum-inv-modal .uz-hint{padding:11px 14px;border-radius:12px;font-weight:700;font-size:13px;color:var(--c,var(--muted));background:var(--cs,var(--surface-2));border:1px solid color-mix(in srgb,var(--c,var(--muted)) 30%,transparent);border-left:4px solid var(--c,var(--muted))}

#uzum-inv-modal .uz-scanrow{display:flex;gap:10px;margin-bottom:16px}
#uzum-inv-modal .uz-scan{flex:1;min-width:0;min-height:52px;padding:0 18px;border-radius:14px;border:2px solid var(--primary);background:var(--surface);color:var(--text);font-size:15px;font-weight:600;outline:none;box-shadow:0 0 0 4px color-mix(in srgb,var(--primary) 12%,transparent)}
#uzum-inv-modal .uz-scan:focus{box-shadow:0 0 0 5px color-mix(in srgb,var(--primary) 25%,transparent)}
#uzum-inv-modal .uz-listh{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 2px 10px;flex-wrap:wrap}
#uzum-inv-modal .uz-listh span{font-weight:800;font-size:14px}
#uzum-inv-modal .uz-list{max-height:300px;overflow:auto;background:var(--surface);border:1px solid var(--line);border-radius:16px}
#uzum-inv-modal .uz-list .uz-empty{border:0;border-radius:0}

#uzum-inv-modal .uz-tw{overflow-x:auto;-webkit-overflow-scrolling:touch}
#uzum-inv-modal .uz-tbl{width:100%;min-width:560px;border-collapse:separate;border-spacing:0;font-size:13px}
#uzum-inv-modal .uz-tbl th{position:sticky;top:0;z-index:1;padding:11px 14px;text-align:left;background:var(--surface-2);color:var(--muted);font-size:10.5px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;white-space:nowrap}
#uzum-inv-modal .uz-tbl td{padding:12px 14px;border-top:1px solid var(--line);vertical-align:middle}
#uzum-inv-modal .uz-tbl tbody tr:hover td{background:var(--surface-2)}
#uzum-inv-modal .uz-tbl .r{text-align:right;font-weight:800;white-space:nowrap}
#uzum-inv-modal .uz-mono{font-family:ui-monospace,Consolas,monospace;font-weight:700}

#uzum-inv-modal .uz-tag{display:inline-flex;align-items:center;padding:3px 9px;margin-left:6px;border-radius:999px;font-size:10.5px;font-weight:800;letter-spacing:.3px;white-space:nowrap;vertical-align:middle;color:var(--c,var(--muted));background:var(--cs,var(--surface-2))}
#uzum-inv-modal .uz-st{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;font-weight:800;white-space:nowrap;color:var(--c,var(--muted))}
#uzum-inv-modal .uz-st::before{content:"";width:7px;height:7px;border-radius:50%;background:currentColor}

#uzum-inv-modal .uz-tiles{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin-bottom:16px}
#uzum-inv-modal .uz-tile{min-width:0;padding:14px;border-radius:16px;color:var(--c,var(--muted));background:var(--cs,var(--surface));border:1px solid color-mix(in srgb,var(--c,var(--muted)) 28%,transparent)}
#uzum-inv-modal .uz-tile .n{font-size:26px;font-weight:900;line-height:1;font-variant-numeric:tabular-nums}
#uzum-inv-modal .uz-tile .t{margin-top:7px;font-size:12px;font-weight:700;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#uzum-inv-modal .uz-tile .u{margin-top:4px;font-size:12px;font-weight:800;color:var(--c,var(--muted));white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#uzum-inv-modal .uz-meta{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
#uzum-inv-modal .uz-chip{display:inline-flex;align-items:center;gap:6px;padding:7px 12px;border-radius:999px;background:var(--surface);border:1px solid var(--line);font-size:12.5px;font-weight:700;color:var(--muted)}
#uzum-inv-modal .uz-chip b{color:var(--text);font-weight:900;font-variant-numeric:tabular-nums}

#uzum-inv-modal .uz-sec{background:var(--surface);border:1px solid var(--line);border-radius:18px;overflow:hidden;margin-bottom:14px}
#uzum-inv-modal .uz-sec-h{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;padding:13px 18px;border-bottom:1px solid var(--line);background:var(--cs,transparent)}
#uzum-inv-modal .uz-sec-t{display:flex;align-items:center;gap:10px;font-weight:800;font-size:14px;color:var(--c,var(--text))}
#uzum-inv-modal .uz-sec-t::before{content:"";width:10px;height:10px;border-radius:4px;background:var(--c,var(--muted))}
#uzum-inv-modal .uz-pill{font-size:12px;font-weight:700;color:var(--muted);background:var(--surface);border:1px solid var(--line);padding:4px 11px;border-radius:999px;white-space:nowrap}
#uzum-inv-modal .uz-none{padding:18px;color:var(--muted);font-size:13px}
#uzum-inv-modal .uz-none.ok{color:var(--ok);font-weight:700}

#uzum-inv-modal .uz-hsearch{width:min(340px,100%);min-height:44px;padding:0 14px;margin-bottom:14px;border-radius:12px;border:1px solid var(--line-2);background:var(--surface);color:var(--text);outline:none}
#uzum-inv-modal .uz-hsearch:focus{border-color:var(--primary);box-shadow:0 0 0 4px color-mix(in srgb,var(--primary) 18%,transparent)}
#uzum-inv-modal .uz-hcard{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:14px;padding:16px 18px;margin-bottom:12px;background:var(--surface);border:1px solid var(--line);border-radius:18px;min-width:0}
#uzum-inv-modal .uz-hcard:hover{border-color:var(--line-2)}
#uzum-inv-modal .uz-hname{font-weight:800;font-size:15px;min-width:0}
#uzum-inv-modal .uz-hname small{display:block;margin-top:3px;font-size:12px;color:var(--muted);font-weight:600}
#uzum-inv-modal .uz-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
#uzum-inv-modal .uz-tc{font-size:11.5px;font-weight:800;padding:4px 10px;border-radius:8px;color:var(--c,var(--muted));background:var(--cs,var(--surface-2))}
#uzum-inv-modal .uz-acts{display:flex;gap:8px;flex-wrap:wrap}

#uzum-inv-modal .uz-setcard{padding:6px 20px}
#uzum-inv-modal .uz-in-wide{width:min(340px,60vw);min-height:40px;font-size:13.5px;font-weight:600}
#uzum-inv-modal .uz-setctl .uz-dpst{margin-top:8px;padding:7px 10px;font-size:12px}
#uzum-inv-modal .uz-setrow{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:16px 0;border-bottom:1px solid var(--line)}
#uzum-inv-modal .uz-setrow:last-child{border-bottom:0}
#uzum-inv-modal .uz-setinfo{display:flex;flex-direction:column;gap:3px;min-width:0}
#uzum-inv-modal .uz-setinfo b{font-size:14px;font-weight:800}
#uzum-inv-modal .uz-setinfo .uz-muted{font-size:12.5px}
#uzum-inv-modal .uz-setctl{flex-shrink:0}
#uzum-inv-modal .uz-in-sm{width:auto;min-width:120px;min-height:40px;font-size:14px}
#uzum-inv-modal .uz-toggle{display:inline-flex;align-items:center;gap:10px;min-height:34px;padding:4px 12px 4px 4px;border-radius:999px;border:1px solid var(--line-2);background:var(--surface-2);color:var(--muted);font-weight:700;font-size:12.5px;cursor:pointer}
#uzum-inv-modal .uz-toggle-k{width:26px;height:26px;border-radius:50%;background:var(--surface-3);transition:transform .2s ease,background .2s ease}
#uzum-inv-modal .uz-toggle.on{color:var(--ok);border-color:color-mix(in srgb,var(--ok) 40%,transparent)}
#uzum-inv-modal .uz-toggle.on .uz-toggle-k{background:var(--ok);transform:translateX(0)}
#uzum-inv-modal .uz-toggle:not(.on) .uz-toggle-k{background:var(--line-2)}
#uzum-inv-modal .uz-setfoot{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-top:14px}

@media (max-width:980px){
  #uzum-inv-modal .uz-app{grid-template-columns:minmax(0,1fr);grid-template-rows:auto minmax(0,1fr)}
  #uzum-inv-modal .uz-side{flex-direction:row;flex-wrap:wrap;align-items:center;gap:10px;padding:12px 16px;border-right:0;border-bottom:1px solid var(--line);overflow:visible}
  #uzum-inv-modal .uz-brand-s{display:none}
  #uzum-inv-modal .uz-nav{flex-direction:row;flex:1;min-width:0}
  #uzum-inv-modal .uz-navbtn{width:auto;padding:9px 12px}
  #uzum-inv-modal .uz-pvzcard{display:none}
  #uzum-inv-modal .uz-tiles{grid-template-columns:repeat(3,minmax(0,1fr))}
  #uzum-inv-modal .uz-stats{grid-template-columns:repeat(2,minmax(0,1fr))}
  #uzum-inv-modal .uz-stat-ring{grid-column:span 2}
}
@media (max-width:640px){
  #uzum-inv-modal{padding:0}
  #uzum-inv-modal .uz-app{width:100%;height:100%;border-radius:0;border:0}
  #uzum-inv-modal .uz-side{padding:10px 12px}
  #uzum-inv-modal .uz-brand-t{font-size:14px}
  #uzum-inv-modal .uz-navbtn{padding:8px 10px;font-size:12.5px}
  #uzum-inv-modal .uz-head{padding:16px 16px 10px}
  #uzum-inv-modal .uz-h1{font-size:18px;white-space:normal}
  #uzum-inv-modal .uz-steps{padding:0 16px 10px}
  #uzum-inv-modal .uz-step .lbl{display:none}
  #uzum-inv-modal .uz-body{padding:4px 14px 20px}
  #uzum-inv-modal .uz-form{grid-template-columns:1fr}
  #uzum-inv-modal .uz-stats{gap:8px}
  #uzum-inv-modal .uz-stat{padding:12px}
  #uzum-inv-modal .uz-stat b{font-size:19px}
  #uzum-inv-modal .uz-work{grid-template-columns:1fr}
  #uzum-inv-modal .uz-scanrow{gap:8px}
  #uzum-inv-modal .uz-scanrow .uz-btn{padding:0 12px}
  #uzum-inv-modal .uz-tiles{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}
  #uzum-inv-modal .uz-tile{padding:12px}
  #uzum-inv-modal .uz-tile .n{font-size:22px}
  #uzum-inv-modal .uz-tbl{min-width:0}
  #uzum-inv-modal .uz-tbl thead{display:none}
  #uzum-inv-modal .uz-tbl,#uzum-inv-modal .uz-tbl tbody,#uzum-inv-modal .uz-tbl tr,#uzum-inv-modal .uz-tbl td{display:block;width:100%}
  #uzum-inv-modal .uz-tbl tr{padding:10px 14px;border-top:1px solid var(--line)}
  #uzum-inv-modal .uz-tbl td{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:5px 0;border:0;text-align:right}
  #uzum-inv-modal .uz-tbl td::before{content:attr(data-l);flex-shrink:0;font-size:11px;font-weight:800;letter-spacing:.4px;text-transform:uppercase;color:var(--muted);text-align:left}
  #uzum-inv-modal .uz-tbl td:empty{display:none}
  #uzum-inv-modal .uz-tbl tbody tr:hover td{background:none}
  #uzum-inv-modal .uz-head-a{width:100%;justify-content:flex-start}
  #uzum-inv-modal .uz-hcard{flex-direction:column;align-items:stretch}
  #uzum-inv-modal .uz-acts .uz-btn{flex:1}
  #uzum-inv-modal .uz-hsearch{width:100%}
}
@media (max-width:420px){
  #uzum-inv-modal .uz-stats{grid-template-columns:1fr 1fr}
  #uzum-inv-modal .uz-stat-ring{grid-column:span 2}
  #uzum-inv-modal .uz-tiles{grid-template-columns:1fr 1fr}
}
`;

  function ensureStyle() {
    if ($('uzum-inv-style')) return;
    const st = document.createElement('style');
    st.id = 'uzum-inv-style';
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  // ---------------------------------------------------------------- UI helpers
  function tbl(heads, rows) {
    const head = heads.map(h => `<th>${escHtml(h)}</th>`).join('');
    const body = rows.map(cells => `<tr>${cells.map((c, i) => `<td data-l="${escHtml(heads[i] || '')}">${c}</td>`).join('')}</tr>`).join('');
    return `<div class="uz-tw"><table class="uz-tbl"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  }

  function setDpStatus(kind, text) {
    const el = $('u-dp-status');
    if (!el) return;
    const cls = { idle: 's-mut', busy: 's-search', ok: 's-ok', bad: 's-bad' }[kind] || 's-mut';
    el.className = 'uz-dpst ' + cls;
    el.textContent = text;
  }

  function isModalOpen() {
    const m = $('uzum-inv-modal');
    return !!m && m.style.display === 'flex';
  }

  function isScanningActive() {
    return isModalOpen() && state.screen === 'scan';
  }

  function isTypingElsewhere() {
    const a = document.activeElement;
    if (!a || a === $('u-barcode-input')) return false;
    return a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT' || a.isContentEditable;
  }

  function focusScanInput() {
    const inp = $('u-barcode-input');
    if (inp && isScanningActive()) inp.focus({ preventScroll: true });
  }

  function closeModal() {
    const m = $('uzum-inv-modal');
    if (m) m.style.display = 'none';
  }

  function applyStaticTexts() {
    const set = (id, txt) => { const el = $(id); if (el) el.textContent = txt; };
    set('u-app-name', t('appName'));
    set('u-tab-new-t', t('navNew'));
    set('u-tab-history-t', t('navHist'));
    set('u-tab-settings-t', t('navSet'));
    set('u-side-lbl', t('pvz'));
    set('u-step-1', t('stepParams'));
    set('u-step-2', t('stepScan'));
    set('u-step-3', t('stepReport'));
    set('u-lbl-pvz', t('pvz'));
    set('u-lbl-dpkey', t('dpKeyLabel'));
    set('u-lbl-issue', t('issueLabel'));
    set('u-ordersnote', t('ordersFromPage'));
    set('u-load-btn', t('loadBtn'));
    set('u-lbl-stat-orders', t('statOrders'));
    set('u-lbl-stat-plan', t('statPlan'));
    set('u-lbl-stat-scanned', t('statScanned'));
    set('u-lbl-cell', t('cellActive'));
    set('u-lbl-last', t('lastBarcode'));
    set('u-undo-btn', t('undo'));
    set('u-lbl-cell-items', t('cellItems'));
    set('u-lbl-timer', t('timerLabel'));
    const close = $('u-close');
    if (close) { close.title = t('close'); close.setAttribute('aria-label', t('close')); }
    const scanInp = $('u-barcode-input');
    if (scanInp) scanInp.placeholder = t('scanPlaceholder');
    const hs = $('u-hist-search');
    if (hs) hs.placeholder = t('histSearch');
  }

  function updateSteps(name) {
    const order = ['start', 'scan', 'report'];
    const cur = order.indexOf(name);
    const modal = $('uzum-inv-modal');
    if (!modal) return;
    modal.querySelectorAll('.uz-step').forEach(el => {
      const idx = order.indexOf(el.dataset.step);
      el.classList.toggle('done', cur > idx);
      el.classList.toggle('on', cur === idx);
    });
  }

  function renderHeadActions(name) {
    const box = $('u-head-actions');
    if (!box) return;
    let html = '';
    if (name === 'scan') {
      html = `<button type="button" class="uz-btn ghost" id="u-restart-btn">${escHtml(t('restartBtn'))}</button>` +
             `<button type="button" class="uz-btn pri" id="u-finish-btn">${escHtml(t('finish'))}</button>`;
    } else if (name === 'report') {
      html = `<button type="button" class="uz-btn ghost" id="u-back-btn">${escHtml(t('btnBack'))}</button>` +
             `<button type="button" class="uz-btn ghost" id="u-excel-btn">${escHtml(t('btnExcel'))}</button>` +
             `<button type="button" class="uz-btn ok" id="u-save-btn" ${state.savedInSession ? 'disabled' : ''}>${escHtml(state.savedInSession ? t('btnSaved') : t('btnSave'))}</button>`;
    }
    box.innerHTML = html;
    const bind = (id, fn) => { const el = $(id); if (el) el.onclick = fn; };
    bind('u-restart-btn', restartRevision);
    bind('u-finish-btn', () => showScreen('report'));
    bind('u-back-btn', () => showScreen('scan'));
    bind('u-excel-btn', () => exportCleanXlsx());
    bind('u-save-btn', saveAuditPermanently);
  }

  function updatePageHead(name) {
    const pvz = state.activePvzName;
    const map = {
      start: [t('setupTitle'), t('setupHint')],
      scan: [t('scanTitle', pvz), t('scanSub')],
      report: [t('reportTitle', pvz), t('reportSub')],
      history: [t('histTitle'), t('histSub')],
      settings: [t('setTitle'), t('setSub')]
    };
    const pair = map[name] || map.start;
    $('u-h1').textContent = pair[0];
    $('u-sub').textContent = pair[1];
    const steps = $('u-steps');
    if (steps) steps.style.display = (name === 'history' || name === 'settings') ? 'none' : '';
    updateSteps(name);
    renderHeadActions(name);
  }

  function updateSidePvz() {
    const el = $('u-side-pvz');
    if (!el) return;
    el.textContent = state.activePvzName || '—';
  }

  function showScreen(name) {
    if (name === 'report') renderReportScreen();
    if (name === 'history') renderHistoryScreen();
    if (name === 'settings') renderSettingsScreen();
    state.screen = name;
    if (name === 'scan' || name === 'report' || name === 'start') state.workScreen = name;
    ['start', 'scan', 'report', 'history', 'settings'].forEach(n => {
      const el = $(`u-${n}-screen`);
      if (el) el.style.display = n === name ? '' : 'none';
    });
    const isWork = !['history', 'settings'].includes(name);
    $('u-tab-new').classList.toggle('active', isWork);
    $('u-tab-history').classList.toggle('active', name === 'history');
    $('u-tab-settings').classList.toggle('active', name === 'settings');
    updateSidePvz();
    updatePageHead(name);
    if (name === 'scan') focusScanInput();
  }

  // ---------------------------------------------------------------- scan screen
  function renderCellItems() {
    const box = $('u-cell-items');
    if (!box) return;
    const cell = state.activeCell;
    const orderItems = (cell && state.factCellUnits[cell]) || {};
    const accItems = (cell && state.acceptanceScans[cell]) || {};
    const retItems = (cell && state.returnScans[cell]) || {};
    const forItems = (cell && state.foreignScans[cell]) || {};
    const rows = [];

    Object.entries(orderItems).forEach(([oId, qty]) => {
      const o = state.ordersMap[oId] || {};
      rows.push([
        `<b>${escHtml(t('orderLabel', o.pid || oId))}</b>${kindBadgeHtml(o)} <span class="uz-muted">${safeCustomerName(o.customer)}</span>`,
        `<span class="r uz-c s-ok">${escHtml(t('pcs', qty))}</span>`
      ]);
    });
    Object.entries(accItems).forEach(([code, d]) => {
      rows.push([
        `<span class="uz-mono">${escHtml(code)}</span><span class="uz-tag s-acc">${escHtml(t('badgeAcc'))}</span>`,
        `<span class="r uz-c s-acc">+${escHtml(t('pcs', d.count))}</span>`
      ]);
    });
    Object.entries(retItems).forEach(([code, d]) => {
      rows.push([
        `<span class="uz-mono">${escHtml(code)}</span><span class="uz-tag s-ret">${escHtml(t('badgeRet'))}</span>`,
        `<span class="r uz-c s-ret">+${escHtml(t('pcs', d.count))}</span>`
      ]);
    });
    Object.entries(forItems).forEach(([code, qty]) => {
      rows.push([
        `<span class="uz-mono">${escHtml(code)}</span><span class="uz-tag s-bad">${escHtml(t('badgeForeign'))}</span>`,
        `<span class="r uz-c s-bad">+${escHtml(t('pcs', qty))}</span>`
      ]);
    });

    if (!rows.length) {
      box.innerHTML = `<div class="uz-empty">${escHtml(cell ? t('emptyCell') : t('hintNeedCell'))}</div>`;
      return;
    }
    box.innerHTML = tbl([t('colOrderBarcode'), t('colQty')], rows);
  }

  function updateScannerUI(lastBarcode, statusType = 'NORMAL', hintText = '') {
    const cellEl = $('u-active-cell');
    if (!cellEl) return;

    cellEl.textContent = state.activeCell || t('cellPlaceholder');
    cellEl.classList.toggle('empty', !state.activeCell);
    if (lastBarcode != null) $('u-last-barcode').textContent = lastBarcode;

    $('u-sc-orders').textContent = state.totalOrdersCount;
    $('u-sc-plan').textContent = t('pcs', state.totalPlanUnits);
    $('u-sc-total').textContent = t('pcs', state.totalScannedUnits);

    const found = Object.values(state.factGlobal).reduce((a, b) => a + b, 0);
    const pct = state.totalPlanUnits ? Math.min(100, Math.round((found / state.totalPlanUnits) * 100)) : 0;
    const fg = $('u-ring-fg');
    fg.style.strokeDasharray = String(RING_C);
    fg.style.strokeDashoffset = String(RING_C * (1 - pct / 100));
    $('u-ring-t').textContent = pct + '%';
    $('u-bar-fill').style.width = pct + '%';
    $('u-bar-label').textContent = t('barLabel', found, state.totalPlanUnits);

    const hintCls = { OK: 's-ok', INFO: 's-warn', ACCEPTANCE: 's-acc', RETURN: 's-ret', FOREIGN: 's-bad', SEARCH: 's-search', CELL: 's-search' };
    const hint = $('u-scan-hint');
    hint.className = 'uz-hint ' + (hintCls[statusType] || '');
    hint.textContent = hintText || (state.activeCell ? t('hintScanItems') : t('hintNeedCell'));

    renderCellItems();
    updateTimerUI();
  }

  // ---------------------------------------------------------------- report screen
  function renderReportScreen() {
    refreshLocale();
    const report = buildReport();
    const partnerAll = report.partnerOrders;
    const partnerIssues = partnerAll.filter(p => p.status !== 'OK');
    const issuesTotal = report.pureShortages.length + report.mismatchedCells.length + report.inAcceptance.length +
      report.canceledReturns.length + report.pureSurplus.length + partnerIssues.length;

    const tiles = [
      ['s-bad', t('tileShort'), report.pureShortages.length, `-${t('pcs', report.totalShortageUnits)}`],
      ['s-warn', t('tileMis'), report.mismatchedCells.length, t('pcs', report.totalMismatchUnits)],
      ['s-acc', t('tileAcc'), report.inAcceptance.length, t('pcs', report.totalAcceptanceUnits)],
      ['s-ret', t('tileRet'), report.canceledReturns.length, t('pcs', report.totalReturnUnits)],
      ['s-mut', t('tileSur'), report.pureSurplus.length, `+${t('pcs', report.totalSurplusUnits)}`],
      ['s-par', t('tilePartner'), partnerIssues.length, t('ofTotal', partnerAll.length)]
    ].map(([c, label, n, u]) =>
      `<div class="uz-tile ${c}"><div class="n">${n}</div><div class="t">${escHtml(label)}</div><div class="u">${escHtml(u)}</div></div>`
    ).join('');

    const meta = `<div class="uz-meta">
      <span class="uz-chip">${escHtml(t('metaOrders'))} <b>${state.totalOrdersCount}</b></span>
      <span class="uz-chip">${escHtml(t('metaPlan'))} <b>${escHtml(t('pcs', state.totalPlanUnits))}</b></span>
      <span class="uz-chip">${escHtml(t('metaScanned'))} <b>${escHtml(t('pcs', state.totalScannedUnits))}</b></span>
      <span class="uz-chip">${escHtml(t('issuesPill', issuesTotal))}</span>
      ${timeMetaHtml(state.totalScannedUnits)}
    </div>`;

    const sec = (cls, title, cnt, units, body) => `
      <section class="uz-sec ${cls}">
        <div class="uz-sec-h"><span class="uz-sec-t">${escHtml(title)}</span><span class="uz-pill">${escHtml(t('pos', cnt))} · ${escHtml(units)}</span></div>
        ${body}
      </section>`;
    const none = (txt, ok) => `<div class="uz-none ${ok ? 'ok' : ''}">${escHtml(txt)}</div>`;

    let html = meta + `<div class="uz-tiles">${tiles}</div>`;

    html += sec('s-bad', t('secShort'), report.pureShortages.length, `-${t('pcs', report.totalShortageUnits)}`,
      report.pureShortages.length === 0 ? none(t('noShort'), true) :
      tbl([t('thPid'), t('thMissing'), t('thShelf'), t('thClient')], report.pureShortages.map(r => [
        pidHtml(r),
        `<span class="r uz-c s-bad">-${escHtml(t('pcs', r.missingCount))}</span>`,
        escHtml(cellLabel(r.expectedCell)),
        `${safeCustomerName(r.customer)} <span class="uz-muted">${escHtml(r.customer?.phone || '')}</span>`
      ])));

    html += sec('s-warn', t('secMis'), report.mismatchedCells.length, t('pcs', report.totalMismatchUnits),
      report.mismatchedCells.length === 0 ? none(t('noMis'), true) :
      tbl([t('thPid'), t('thQty'), t('thPlanCell'), t('thActualCell')], report.mismatchedCells.map(r => [
        pidHtml(r),
        `<span class="r">${escHtml(t('pcs', r.count))}</span>`,
        `<span class="uz-c s-bad">${escHtml(cellLabel(r.expectedCell))}</span>`,
        `<span class="uz-c s-ok">${escHtml(r.actualCell)}</span>`
      ])));

    html += sec('s-acc', t('secAcc'), report.inAcceptance.length, t('pcs', report.totalAcceptanceUnits),
      report.inAcceptance.length === 0 ? none(t('noAcc'), false) :
      tbl([t('thBarcodeLoc'), t('thQty'), t('thShelf'), t('thRec')], report.inAcceptance.map(r => [
        `<span class="uz-mono">${escHtml(r.code)}</span>`,
        `<span class="r uz-c s-acc">+${escHtml(t('pcs', r.count))}</span>`,
        escHtml(`${t('cellWord')} ${r.foundCell}`),
        `<b>${escHtml(t('recAcc'))}</b>`
      ])));

    html += sec('s-ret', t('secRet'), report.canceledReturns.length, t('pcs', report.totalReturnUnits),
      report.canceledReturns.length === 0 ? none(t('noRet'), false) :
      tbl([t('thBarcodeLoc'), t('thQty'), t('thShelf'), t('thRec')], report.canceledReturns.map(r => [
        `<span class="uz-mono">${escHtml(r.code)}</span>`,
        `<span class="r uz-c s-ret">+${escHtml(t('pcs', r.count))}</span>`,
        escHtml(`${t('cellWord')} ${r.foundCell}`),
        `<b>${escHtml(t('recRet'))}</b>`
      ])));

    html += sec('s-mut', t('secSur'), report.pureSurplus.length, `+${t('pcs', report.totalSurplusUnits)}`,
      report.pureSurplus.length === 0 ? none(t('noSur'), false) :
      tbl([t('thBarcode'), t('thExtra'), t('thFoundAt')], report.pureSurplus.map(r => [
        `<span class="uz-mono">${escHtml(r.code)}</span>`,
        `<span class="r uz-c s-bad">+${escHtml(t('pcs', r.qty))}</span>`,
        escHtml(`${t('cellWord')} ${r.foundCell}`)
      ])));

    html += sec('s-par', t('secPartner'), partnerAll.length, t('pcs', report.totalPartnerUnits),
      partnerAll.length === 0 ? none(t('noPartner'), false) :
      tbl([t('thPid'), t('thStatus'), t('thPlanFact'), t('thLocation'), t('thClient'), t('thRec')], partnerAll.map(p => [
        pidHtml(p),
        `<span class="uz-st ${statusCls(p.status)}">${escHtml(statusText(p.status))}</span>`,
        `<span class="uz-mono">${p.planQty} / ${p.factQty}</span>`,
        escHtml(plainFound(p.found)),
        `${safeCustomerName(p.customer)} <span class="uz-muted">${escHtml(p.customer?.phone || '')}</span>`,
        escHtml(partnerAction(p.status))
      ])));

    $('u-report-screen').innerHTML = html;
  }

  // ---------------------------------------------------------------- history
  function renderHistoryScreen(filter = '') {
    refreshLocale();
    const box = $('u-hist-container');
    if (!box) return;
    const list = getHistoryList();

    if (!list.length) {
      box.innerHTML = `<div class="uz-empty">${escHtml(t('histEmpty'))}</div>`;
      return;
    }

    const q = String(filter || '').trim().toLowerCase();
    const items = q ? list.filter(x => JSON.stringify(x).toLowerCase().includes(q)) : list;
    if (!items.length) {
      box.innerHTML = `<div class="uz-empty">${escHtml(t('histNotFound'))}</div>`;
      return;
    }

    box.innerHTML = items.map(item => {
      const s = item.stats || {};
      const tm = item.timing;
      const chips = [
        `<span class="uz-tc">${escHtml(t('chipOrders', s.totalOrders ?? 0))}</span>`,
        `<span class="uz-tc">${escHtml(t('chipPlan', s.totalPlanUnits ?? 0))}</span>`,
        `<span class="uz-tc">${escHtml(t('chipFact', s.totalScanned ?? 0))}</span>`,
        tm ? `<span class="uz-tc s-info">${escHtml(t('chipTime', fmtDur(tm.totalMs || 0), fmtDur(tm.activeMs || 0)))}</span>` : '',
        `<span class="uz-tc s-bad">${escHtml(t('chipShort', s.shortageCount ?? 0))}</span>`,
        `<span class="uz-tc s-warn">${escHtml(t('chipMis', s.mismatchCount ?? 0))}</span>`,
        `<span class="uz-tc s-acc">${escHtml(t('chipAcc', s.acceptanceCount ?? 0))}</span>`,
        `<span class="uz-tc s-ret">${escHtml(t('chipRet', s.returnCount ?? 0))}</span>`,
        `<span class="uz-tc s-par">${escHtml(t('chipPartner', s.partnerCount ?? 0))}</span>`
      ].join('');
      return `
        <div class="uz-hcard">
          <div class="uz-hname">
            ${escHtml(item.pvzName)}
            <small>ID ${escHtml(item.dpKey)} · ${escHtml(item.date)} ${escHtml(item.time)} · ${escHtml(item.auditor || '')}</small>
            <div class="uz-chips">${chips}</div>
          </div>
          <div class="uz-acts">
            <button type="button" class="uz-btn pri u-dl" data-id="${escHtml(item.id)}">${escHtml(t('btnXlsx'))}</button>
            <button type="button" class="uz-btn danger u-del" data-id="${escHtml(item.id)}" title="${escHtml(t('deleteTitle'))}">✕</button>
          </div>
        </div>`;
    }).join('');

    box.querySelectorAll('.u-dl').forEach(b => {
      b.onclick = () => {
        const found = getHistoryList().find(x => x.id === b.getAttribute('data-id'));
        if (!found) return;
        exportCleanXlsx(found.report, {
          date: found.date,
          time: found.time,
          pvzName: found.pvzName,
          dpKey: found.dpKey,
          auditor: found.auditor,
          totalOrders: found.stats?.totalOrders,
          totalPlanUnits: found.stats?.totalPlanUnits,
          totalScanned: found.stats?.totalScanned,
          timing: found.timing || null
        });
      };
    });

    box.querySelectorAll('.u-del').forEach(b => {
      b.onclick = () => {
        if (!confirm(t('confirmDelete'))) return;
        const id = b.getAttribute('data-id');
        writeHistory(getHistoryList().filter(x => x.id !== id));
        renderHistoryScreen($('u-hist-search').value);
      };
    });
  }

  // ---------------------------------------------------------------- telegram config check
  function tgTokenStatus(tok) {
    if (!tok) return `<div class="uz-dpst s-bad">${escHtml(t('tgNeedToken'))}</div>`;
    if (!TG_TOKEN_RE.test(tok)) return `<div class="uz-dpst s-bad">${escHtml(t('tgBadToken', TG_BOT_PREFIX))}</div>`;
    return `<div class="uz-dpst s-ok">${escHtml(t('tgTokenFormatOk'))}</div>`;
  }

  // Inventni boshlashdan oldin: token, chat ID lar va tokenning haqiqiyligi (getMe) tekshiriladi
  async function checkTelegramConfig() {
    const token = getTgToken();
    if (!token) return t('tgNeedToken');
    if (!TG_TOKEN_RE.test(token)) return t('tgBadToken', TG_BOT_PREFIX);
    const chats = getTgChats();
    if (!chats.length) return t('tgNeedChat');
    if (chats.some(c => !TG_CHAT_RE.test(c))) return t('tgBadChat');
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 10000);
      const res = await fetch(`https://api.telegram.org/bot${token}/getMe`, { signal: ctrl.signal });
      clearTimeout(timer);
      const j = await res.json().catch(() => null);
      if (!res.ok || !j || !j.ok) return t('tgCheckFail');
    } catch (e) {
      return t('tgCheckFail');
    }
    return '';
  }

  // ---------------------------------------------------------------- settings screen
  function settingRow(key, title, desc, control) {
    return `<div class="uz-setrow">
      <div class="uz-setinfo"><b>${escHtml(title)}</b><span class="uz-muted">${escHtml(desc)}</span></div>
      <div class="uz-setctl">${control}</div>
    </div>`;
  }

  function toggleCtl(key, on) {
    return `<button type="button" class="uz-toggle ${on ? 'on' : ''}" data-set="${key}" data-type="bool" aria-pressed="${on}">
      <span class="uz-toggle-k"></span><span class="uz-toggle-t">${escHtml(on ? t('setOn') : t('setOff'))}</span></button>`;
  }

  function renderSettingsScreen() {
    refreshLocale();
    const s = getSettings();
    const langOpts = [['auto', t('langAuto')], ['ru', 'Русский'], ['uz', "O'zbekcha"]]
      .map(([v, label]) => `<option value="${v}" ${s.lang === v ? 'selected' : ''}>${escHtml(label)}</option>`).join('');

    const rows = [
      settingRow('sound', t('setSound'), t('setSoundD'), toggleCtl('sound', s.sound)),
      settingRow('tgToken', t('setTgTok'), t('setTgTokD'),
        `<input type="text" class="uz-in uz-in-wide" data-set="tgToken" data-type="str" autocomplete="off" spellcheck="false" placeholder="${escHtml(TG_BOT_PREFIX)}..." value="${escHtml(s.tgToken)}">${tgTokenStatus(s.tgToken)}`),
      settingRow('tgChats', t('setTgChat'), t('setTgChatD'),
        `<input type="text" class="uz-in uz-in-wide" data-set="tgChats" data-type="str" autocomplete="off" spellcheck="false" placeholder="123456789, -1001234567890" value="${escHtml(s.tgChats)}">`),
      settingRow('autoFocus', t('setFocus'), t('setFocusD'), toggleCtl('autoFocus', s.autoFocus)),
      settingRow('hotkeys', t('setHot'), t('setHotD'), toggleCtl('hotkeys', s.hotkeys)),
      settingRow('idleMinutes', t('setIdle'), t('setIdleD'),
        `<input type="number" class="uz-in uz-in-sm" data-set="idleMinutes" data-type="num" min="1" max="10" value="${s.idleMinutes}">`),
      settingRow('lang', t('setLang'), t('setLangD'),
        `<select class="uz-in uz-in-sm" data-set="lang" data-type="enum">${langOpts}</select>`)
    ].join('');

    $('u-settings-screen').innerHTML = `
      <div class="uz-card uz-setcard">${rows}</div>
      <div class="uz-setfoot">
        <span class="uz-muted" id="u-set-msg"></span>
        <button type="button" class="uz-btn danger" id="u-set-reset">${escHtml(t('setReset'))}</button>
      </div>`;

    const root = $('u-settings-screen');
    root.querySelectorAll('[data-set]').forEach(el => {
      const key = el.getAttribute('data-set');
      const type = el.getAttribute('data-type');
      const apply = () => {
        let val;
        if (type === 'bool') val = !(el.classList.contains('on'));
        else val = el.value;
        setSetting(key, val);
        renderSettingsScreen();
        flashSettingsMsg(t('setSaved'));
      };
      if (type === 'bool') el.onclick = apply;
      else el.onchange = apply;
    });
    $('u-set-reset').onclick = () => {
      try { localStorage.removeItem(SETTINGS_KEY); } catch (e) {}
      applySettings();
      renderSettingsScreen();
      flashSettingsMsg(t('setResetOk'));
    };
  }

  function flashSettingsMsg(text) {
    const el = $('u-set-msg');
    if (!el) return;
    el.textContent = text;
    clearTimeout(flashSettingsMsg._t);
    flashSettingsMsg._t = setTimeout(() => { if (el) el.textContent = ''; }, 2000);
  }

  // ---------------------------------------------------------------- setup / flow
  function fillSetup() {
    const detectedName = detectPvzNameFromHeader();
    const liveDpKey = detectLiveDpKey();
    if (detectedName && detectedName !== 'ПВЗ') state.activePvzName = detectedName;
    if (liveDpKey) state.dpKey = liveDpKey;
    $('u-pvz-name-input').value = state.activePvzName;
    $('u-pvz-dpkey-input').value = state.dpKey || '';
    $('u-issue-count').textContent = detectOrdersCount();
    if (state.dpKey) setDpStatus('idle', t('dpNotChecked'));
    else setDpStatus('bad', t('dpEmpty'));
  }

  function restartRevision() {
    if (state.totalScannedUnits > 0 && !state.savedInSession && !confirm(t('restartConfirm'))) return;
    resetInventoryState();
    resetTiming();
    state.workScreen = 'start';
    fillSetup();
    showScreen('start');
  }

  // Tugma bosilganda: avval ID WMS'da tekshiriladi, keyin yuklash boshlanadi
  async function onLoadClick() {
    if (state.loading) return;
    const btn = $('u-load-btn');
    const prog = $('u-progress');
    const pvzName = $('u-pvz-name-input').value.trim();
    const dpKey = $('u-pvz-dpkey-input').value.trim();

    if (state.totalScannedUnits > 0 && !state.savedInSession && !confirm(t('restartConfirm'))) return;

    const tgErr = await checkTelegramConfig();
    if (tgErr) {
      setDpStatus('bad', tgErr);
      alert(tgErr);
      return;
    }

    state.loading = true;
    btn.disabled = true;
    btn.textContent = t('loadingBtn');
    prog.textContent = t('dpChecking');
    setDpStatus('busy', t('dpChecking'));

    try {
      const check = await validateDpKey(dpKey);

      // Sahifadagi "К выдаче" soni bilan solishtirish (agar WMS jami sonini bersa)
      const pageCount = readIssueCountStrict();
      if (check.total != null && pageCount != null && check.total !== pageCount) {
        if (!confirm(t('dpMismatch', pageCount, check.total))) {
          setDpStatus('idle', t('dpNotChecked'));
          prog.textContent = '';
          return;
        }
      }

      setDpStatus('ok', t('dpOk', dpKey));
      state.activePvzName = pvzName || state.activePvzName;
      state.dpKey = dpKey;

      await loadInventoryData(msg => { prog.textContent = msg; });
      startTiming();
      prog.textContent = '';
      showScreen('scan');
    } catch (err) {
      prog.textContent = '';
      setDpStatus('bad', err.message);
      alert(err.message);
    } finally {
      state.loading = false;
      btn.disabled = false;
      btn.textContent = t('loadBtn');
    }
  }

  function openModal() {
    refreshLocale(true);
    ensureModal();
    applyStaticTexts();
    $('uzum-inv-modal').style.display = 'flex';

    if (state.screen === 'history') {
      showScreen('history');
    } else if (state.screen === 'settings') {
      showScreen('settings');
    } else if (state.workScreen === 'scan' || state.workScreen === 'report') {
      showScreen(state.workScreen);
    } else {
      fillSetup();
      showScreen('start');
    }
  }

  // ---------------------------------------------------------------- modal
  function ensureModal() {
    ensureStyle();
    if ($('uzum-inv-modal')) return;

    const modal = document.createElement('div');
    modal.id = 'uzum-inv-modal';
    modal.innerHTML = `
      <div class="uz-app" role="dialog" aria-modal="true" aria-labelledby="u-h1">
        <aside class="uz-side">
          <div class="uz-brand">
            <div class="uz-logo">R</div>
            <div>
              <div class="uz-brand-t" id="u-app-name"></div>
              <div class="uz-brand-s">v${VERSION}</div>
            </div>
          </div>
          <nav class="uz-nav">
            <button type="button" class="uz-navbtn active" id="u-tab-new"><span class="uz-navic">＋</span><span id="u-tab-new-t"></span></button>
          </nav>
          <div class="uz-pvzcard">
            <span class="uz-lbl" id="u-side-lbl"></span>
            <b id="u-side-pvz">—</b>
          </div>
          <nav class="uz-navbottom">
            <button type="button" class="uz-navbtn" id="u-tab-history"><span class="uz-navic">▤</span><span id="u-tab-history-t"></span></button>
            <button type="button" class="uz-navbtn" id="u-tab-settings"><span class="uz-navic">⚙</span><span id="u-tab-settings-t"></span></button>
          </nav>
        </aside>

        <main class="uz-main">
          <header class="uz-head">
            <div class="uz-head-l">
              <h1 class="uz-h1" id="u-h1"></h1>
              <p class="uz-sub" id="u-sub"></p>
            </div>
            <div class="uz-head-a" id="u-head-actions"></div>
            <button type="button" class="uz-x" id="u-close">✕</button>
          </header>

          <div class="uz-steps" id="u-steps">
            <div class="uz-step" data-step="start"><span class="uz-dot">1</span><span class="lbl" id="u-step-1"></span></div>
            <div class="uz-step" data-step="scan"><span class="uz-dot">2</span><span class="lbl" id="u-step-2"></span></div>
            <div class="uz-step" data-step="report"><span class="uz-dot">3</span><span class="lbl" id="u-step-3"></span></div>
          </div>

          <div class="uz-body">
            <section id="u-start-screen" class="uz-setup" style="display:none">
              <div class="uz-card">
                <div class="uz-hero" aria-hidden="true">📦</div>
                <div class="uz-form">
                  <div class="uz-field">
                    <label class="uz-lbl" for="u-pvz-name-input" id="u-lbl-pvz"></label>
                    <input type="text" id="u-pvz-name-input" class="uz-in" autocomplete="off">
                  </div>
                  <div class="uz-field">
                    <label class="uz-lbl" for="u-pvz-dpkey-input" id="u-lbl-dpkey"></label>
                    <input type="text" id="u-pvz-dpkey-input" class="uz-in" autocomplete="off" inputmode="numeric">
                  </div>
                </div>
                <div class="uz-issue">
                  <span class="uz-lbl" id="u-lbl-issue"></span>
                  <b id="u-issue-count">—</b>
                </div>
                <div class="uz-note" id="u-ordersnote"></div>
                <div id="u-dp-status" class="uz-dpst s-mut"></div>
                <div style="margin-top:18px"><button type="button" id="u-load-btn" class="uz-btn pri big"></button></div>
                <div id="u-progress" class="uz-prog"></div>
              </div>
            </section>

            <section id="u-scan-screen" style="display:none">
              <div class="uz-stats">
                <div class="uz-stat"><span class="uz-lbl" id="u-lbl-stat-orders"></span><b id="u-sc-orders">0</b></div>
                <div class="uz-stat"><span class="uz-lbl" id="u-lbl-stat-plan"></span><b id="u-sc-plan">0</b></div>
                <div class="uz-stat hl"><span class="uz-lbl" id="u-lbl-stat-scanned"></span><b id="u-sc-total">0</b></div>
                <div class="uz-stat uz-stat-ring">
                  <div class="uz-ring">
                    <svg viewBox="0 0 80 80"><circle class="bg" cx="40" cy="40" r="${RING_R}"></circle><circle class="fg" id="u-ring-fg" cx="40" cy="40" r="${RING_R}"></circle></svg>
                    <div class="uz-ring-t" id="u-ring-t">0%</div>
                  </div>
                </div>
              </div>
              <div class="uz-meter"><div id="u-bar-fill"></div></div>
              <div class="uz-barrow">
                <span class="uz-barlbl" id="u-bar-label"></span>
                <span class="uz-timer">
                  <span id="u-lbl-timer"></span>
                  <b id="u-timer">00:00:00</b>
                  <span id="u-timer-total"></span>
                </span>
              </div>

              <div class="uz-work">
                <div class="uz-cellcard">
                  <span class="uz-lbl" id="u-lbl-cell"></span>
                  <div id="u-active-cell" class="uz-cell empty"></div>
                </div>
                <div class="uz-card uz-lastcard">
                  <span class="uz-lbl" id="u-lbl-last"></span>
                  <div id="u-last-barcode" class="uz-code">—</div>
                  <div id="u-scan-hint" class="uz-hint"></div>
                </div>
              </div>

              <div class="uz-scanrow">
                <input type="text" id="u-barcode-input" class="uz-scan" autocomplete="off" spellcheck="false">
                <button type="button" id="u-undo-btn" class="uz-btn"></button>
              </div>

              <div class="uz-listh"><span id="u-lbl-cell-items"></span></div>
              <div id="u-cell-items" class="uz-list"></div>
            </section>

            <section id="u-report-screen" style="display:none"></section>

            <section id="u-history-screen" style="display:none">
              <input type="search" id="u-hist-search" class="uz-hsearch">
              <div id="u-hist-container"></div>
            </section>

            <section id="u-settings-screen" style="display:none"></section>
          </div>
        </main>
      </div>
    `;
    document.body.appendChild(modal);

    // Faollikni kuzatish: ish vaqti shu hodisalar orqali hisoblanadi
    modal.addEventListener('keydown', () => markActivity(), true);
    modal.addEventListener('pointerdown', () => markActivity(), true);

    $('u-close').onclick = closeModal;
    $('u-tab-new').onclick = () => showScreen(state.workScreen || 'start');
    $('u-tab-history').onclick = () => showScreen('history');
    $('u-tab-settings').onclick = () => showScreen('settings');
    $('u-hist-search').oninput = e => renderHistoryScreen(e.target.value);
    $('u-load-btn').onclick = onLoadClick;
    $('u-undo-btn').onclick = () => { undoLastScan(); focusScanInput(); };

    $('u-pvz-dpkey-input').addEventListener('input', e => {
      if (e.target.value.trim()) setDpStatus('idle', t('dpNotChecked'));
      else setDpStatus('bad', t('dpEmpty'));
    });

    const scanInp = $('u-barcode-input');
    scanInp.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const code = scanInp.value.trim();
        scanInp.value = '';
        if (code) enqueueScan(code);
      }
    });
    scanInp.addEventListener('blur', () => {
      setTimeout(() => {
        if (AUTO_FOCUS && isScanningActive() && !isTypingElsewhere()) focusScanInput();
      }, 60);
    });
    $('u-scan-screen').addEventListener('click', e => {
      if (!e.target.closest('button,input')) focusScanInput();
    });
  }

  // ---------------------------------------------------------------- button position
  // Tugma body'ga qo'shiladi va PVZ tanlash maydonining chap tomonida turadi.
  let anchorEl = null;
  let anchorSearchAt = 0;
  let lastPos = '';

  function getAnchor() {
    if (anchorEl && anchorEl.isConnected) {
      const r = anchorEl.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) return anchorEl;
    }
    const now = Date.now();
    if (now - anchorSearchAt < 2000) return null;
    anchorSearchAt = now;
    anchorEl = findPvzDropdown();
    if (anchorEl && anchorEl.getBoundingClientRect().width > 0) return anchorEl;
    return null;
  }

  function positionButton() {
    const btn = $('uzum-inv-btn');
    if (!btn) return;

    const anchor = getAnchor();
    const bw = btn.offsetWidth || 170;
    const bh = btn.offsetHeight || 38;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    let left;
    let top;
    let show = true;

    if (anchor) {
      const r = anchor.getBoundingClientRect();
      if (r.bottom < 0 || r.top > vh) show = false;
      left = r.left - bw - BTN_GAP;
      top = r.top + (r.height - bh) / 2;
      if (left < 8) {
        left = Math.max(8, Math.min(r.right - bw, vw - bw - 8));
        top = r.bottom + 8;
      }
    } else {
      left = Math.max(8, vw - 420 - bw);
      top = 12;
    }

    const key = `${show}|${Math.round(left)}|${Math.round(top)}`;
    if (key === lastPos) return;
    lastPos = key;
    btn.style.display = show ? '' : 'none';
    btn.style.left = Math.round(left) + 'px';
    btn.style.top = Math.round(top) + 'px';
  }

  function ensureButton() {
    ensureStyle();
    let btn = $('uzum-inv-btn');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'uzum-inv-btn';
      btn.type = 'button';
      btn.addEventListener('click', openModal);
      document.body.appendChild(btn);
    }
    const label = `📦 ${t('appName')}`;
    if (btn.textContent !== label) btn.textContent = label;
  }

  function closeOnEscape(e) {
    if (e.key === 'Escape' && isModalOpen()) closeModal();
    // Yangi funksiya: Ctrl+Z — oxirgi skanni bekor qilish (sozlamada yoqilgan bo'lsa)
    if (HOTKEYS_ON && isScanningActive() && (e.ctrlKey || e.metaKey) &&
        !e.shiftKey && e.key.toLowerCase() === 'z' && !isTypingElsewhere()) {
      e.preventDefault();
      undoLastScan();
      focusScanInput();
    }
  }

  // ---------------------------------------------------------------- boot
  window.__uzumState = state;
  refreshLocale(true);
  ensureStyle();
  ensureModal();
  ensureButton();
  positionButton();

  window.addEventListener('resize', positionButton);
  window.addEventListener('scroll', positionButton, true);
  document.addEventListener('keydown', closeOnEscape);
  setInterval(() => {
    refreshLocale();
    ensureButton();
    positionButton();
  }, 400);
  setInterval(() => {
    if (isModalOpen() && state.screen === 'scan') updateTimerUI();
  }, 1000);
})();
