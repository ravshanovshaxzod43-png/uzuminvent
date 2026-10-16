
// ==UserScript==
// @name         Uzum Ревизор ПВЗ
// @namespace    https://dp.uzum.uz/
// @version      1.3.0
// @description  Ревизия склада ПВЗ Uzum Market: потоварная инвентаризация, приёмка, возвраты, пересорт, история и Excel-акт
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

  const VERSION = '1.3.0';
  const API = 'https://api-wms.uzum.uz/de/delivery-point';
  const HISTORY_KEY = 'uzum_inventory_history';
  const jwtRegex = /eyJ[a-zA-Z0-9_-]{25,}\.eyJ[a-zA-Z0-9_-]{25,}\.[a-zA-Z0-9_-]{25,}/g;

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

  function sanitizeFileName(s) {
    return String(s).replace(/[\\/:*?"<>|]+/g, '_');
  }

  function $(id) { return document.getElementById(id); }

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
        if (matches) for (const t of matches) if (isTokenValid(t)) return t;
      }
    } catch (e) {}

    for (const s of [sessionStorage, localStorage]) {
      try {
        for (let i = 0; i < s.length; i++) {
          const matches = s.getItem(s.key(i))?.match(jwtRegex);
          if (matches) for (const t of matches) if (isTokenValid(t)) return t;
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
    if (url && url.includes('/delivery-point/b2c-orders')) {
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

  // перехват fetch
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

  // перехват XMLHttpRequest (axios и др.)
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
        if (res.status === 401 || res.status === 403) throw new Error(`Нет доступа к WMS API (${res.status}). Обновите страницу (F5) и повторите.`);
        if (!res.ok) throw new Error(`Ошибка WMS API (${res.status}).`);
        return await res.json();
      } catch (e) {
        clearTimeout(timer);
        lastErr = e;
        if (/Нет доступа/.test(e.message)) throw e;
      }
    }
    throw lastErr;
  }

  function authHeaders() {
    const token = findTokenAnywhere();
    if (!token) throw new Error("Токен не найден. Обновите вкладку 'Заказы' (F5) и повторите.");
    return { 'accept': 'application/json', 'authorization': `Bearer ${token}` };
  }

  // ---------------------------------------------------------------- header detection
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
        if (url.includes('/delivery-point/b2c-orders') && url.includes('dpKey=')) {
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

  function detectOrdersCount() {
    const all = Array.from(document.querySelectorAll('*'));
    for (const el of all) {
      if (insideOurUi(el)) continue;
      if (el.children.length <= 2 && el.textContent && el.textContent.includes('К выдаче')) {
        const m = el.textContent.match(/К выдаче\s*(\d+)/i) || el.textContent.match(/(\d+)/);
        if (m && m[1]) return m[1];
      }
    }
    return '—';
  }

  // ---------------------------------------------------------------- sound
  let audioCtx = null;
  function beep(freq = 800, duration = 0.08, type = 'sine') {
    try {
      if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
      if (audioCtx.state === 'suspended') audioCtx.resume();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = type;
      osc.frequency.value = freq;
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      const t = audioCtx.currentTime;
      gain.gain.setValueAtTime(0.3, t);
      gain.gain.exponentialRampToValueAtTime(0.00001, t + duration);
      osc.start(t);
      osc.stop(t + duration);
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
    dpKey: null,
    totalOrdersCount: 0,
    totalPlanUnits: 0,
    totalScannedUnits: 0,
    ordersMap: {},
    codeToOrderId: {},   // уникальные идентификаторы (ID, PID, грузоместа)
    skuIndex: {},        // SKU -> [orderId] (SKU могут повторяться между заказами)
    factGlobal: {},      // orderId -> учтено штук
    planCellUnits: {},
    factCellUnits: {},
    acceptanceScans: {},
    returnScans: {},
    foreignScans: {},
    scannedHistory: [],
    lookupCache: new Map(),
    activeCell: null,
    screen: 'start',     // start | scan | report | history
    workScreen: 'start', // последний рабочий экран (для вкладки «Новая ревизия»)
    lastReport: null
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

        // защита от ложных срабатываний: ответ обязан содержать сам отсканированный код
        if (!str.includes(needle)) continue;

        if (str.includes('приемк') || str.includes('acceptance') || str.includes('к размещению') || str.includes('размещен')) {
          const rawObj = Array.isArray(data) ? data[0] : (data.acceptance?.[0] || data.content?.[0] || data);
          result = {
            type: 'ACCEPTANCE',
            label: 'В приёмке (Не принят полностью)',
            orderId: rawObj?.b2cOrderId || rawObj?.orderId || rawObj?.externalId || '',
            cargoPlace: rawObj?.cargoPlaceBarcode || rawObj?.barcode || '',
            statusText: 'Размещен / В приёмке'
          };
          break;
        }

        if (str.includes('отгрузк') || str.includes('завершен') || str.includes('возврат') || str.includes('return') || str.includes('cancel')) {
          const rawObj = Array.isArray(data) ? data[0] : (data.shipment?.[0] || data.returns?.[0] || data.content?.[0] || data);
          result = {
            type: 'RETURN',
            label: 'Отменен / В отгрузке на склад',
            orderId: rawObj?.b2cOrderId || rawObj?.orderId || '',
            cargoPlace: rawObj?.cargoPlaceBarcode || rawObj?.barcode || '',
            statusText: 'Завершен (К возврату)'
          };
          break;
        }
      } catch (e) {
        if (/Нет доступа/.test(e.message)) break;
      }
    }

    state.lookupCache.set(cacheKey, result);
    return result;
  }

  // ---------------------------------------------------------------- load data
  async function loadInventoryData(onProgress) {
    onProgress('Авторизация в WMS...');
    const dpKey = state.dpKey;
    if (!dpKey) throw new Error('ID ПВЗ не указан. Введите ID точки (dpKey) вручную.');
    authHeaders(); // проверка токена

    let page = 0;
    const size = 100;
    const orderHeaders = [];
    const seenHeaderIds = new Set();

    onProgress(`Выгрузка реестра заказов (dpKey: ${dpKey})...`);

    while (page < 500) {
      const data = await fetchJson(
        `${API}/b2c-orders?page=${page}&size=${size}&status=DELIVERED%2CTEMP_DELIVERED&dpKey=${dpKey}`,
        authHeaders(), 30000, 2
      );
      const content = Array.isArray(data) ? data : (data.content || []);
      content.forEach(o => {
        const id = String(o.b2cOrderId);
        if (!seenHeaderIds.has(id)) { seenHeaderIds.add(id); orderHeaders.push(o); }
      });

      let last;
      if (Array.isArray(data)) last = content.length < size;
      else if (typeof data.last === 'boolean') last = data.last;
      else if (typeof data.totalPages === 'number') last = page + 1 >= data.totalPages;
      else last = content.length < size;

      page++;
      onProgress(`Собрано заказов к выдаче: ${orderHeaders.length}...`);
      if (last || content.length === 0) break;
    }

    resetInventoryState();
    state.totalOrdersCount = orderHeaders.length;

    const headerMap = {};
    orderHeaders.forEach(h => { headerMap[String(h.b2cOrderId)] = h; });

    const chunkSize = 25;
    const orderIds = orderHeaders.map(o => String(o.b2cOrderId));
    let failedOrders = 0;

    for (let i = 0; i < orderIds.length; i += chunkSize) {
      const slice = orderIds.slice(i, i + chunkSize);
      onProgress(`Загрузка состава товаров: ${Math.min(i + chunkSize, orderIds.length)} из ${orderIds.length}`);

      let ordersWithItems = [];
      try {
        const r = await fetchJson(`${API}/b2c-orders/items?b2cOrderId=${slice.join(',')}`, authHeaders(), 30000, 2);
        ordersWithItems = Array.isArray(r) ? r : (r.content || []);
      } catch (e) {
        if (/Нет доступа/.test(e.message)) throw e;
        failedOrders += slice.length;
      }

      const itemsMap = {};
      ordersWithItems.forEach(o => { itemsMap[String(o.b2cOrderId)] = o; });

      slice.forEach(orderIdStr => {
        const header = headerMap[orderIdStr] || {};
        const withItems = itemsMap[orderIdStr] || {};

        const pidStr = String(header.b2cOrderPublicId || withItems.b2cOrderPublicId || '').trim();
        const expectedCell = String(header.cellBarcode || withItems.cellBarcode || 'БЕЗ_ЯЧЕЙКИ').trim();
        const customer = header.customer || withItems.customer || {};

        // уникальные коды заказа
        const uniq = new Set();
        uniq.add(orderIdStr);
        uniq.add(orderIdStr.padStart(10, '0'));
        uniq.add(`10-${orderIdStr.padStart(10, '0')}-1`);
        if (pidStr) uniq.add(pidStr);

        const addCargo = (arr) => {
          if (Array.isArray(arr)) {
            arr.forEach(c => {
              const code = typeof c === 'object' ? c?.barcode : c;
              if (code) {
                const s = String(code).trim();
                uniq.add(s);
                const m = s.match(/^10-(\d{10})-\d+$/);
                if (m) uniq.add(String(parseInt(m[1], 10)));
              }
            });
          }
        };
        addCargo(header.cargoPlaceBarcodes);
        addCargo(header.cargoPlaces);
        addCargo(withItems.cargoPlaceBarcodes);
        addCargo(withItems.cargoPlaces);

        // коды товаров (SKU) могут совпадать у разных заказов
        const skuCodes = new Set();
        const items = withItems.items || header.items || [];
        let itemsQty = 0;
        const itemSkus = [];

        items.forEach(it => {
          const qty = Number(it.amount) > 0 ? Number(it.amount) : 1;
          itemsQty += qty;
          const sku = String(it.skuBarcode || it.barcode || '').trim();
          if (sku) {
            skuCodes.add(sku);
            itemSkus.push({ sku, qty, title: it.title || it.name || '' });
          }
          if (it.barcode) skuCodes.add(String(it.barcode).trim());
          if (it.sellerBarcode) skuCodes.add(String(it.sellerBarcode).trim());
          if (Array.isArray(it.barcodes)) it.barcodes.forEach(b => b && skuCodes.add(String(b).trim()));
        });

        if (itemsQty === 0) itemsQty = 1;

        state.ordersMap[orderIdStr] = {
          orderId: orderIdStr,
          pid: pidStr,
          expectedCell,
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

        if (!state.planCellUnits[expectedCell]) state.planCellUnits[expectedCell] = {};
        state.planCellUnits[expectedCell][orderIdStr] = itemsQty;
        state.totalPlanUnits += itemsQty;
      });
    }

    Object.keys(state.planCellUnits).forEach(c => { state.factCellUnits[c] = {}; });

    if (failedOrders > 0) {
      const ok = confirm(`Не удалось загрузить состав ${failedOrders} заказов (для них план принят как 1 шт).\nПродолжить ревизию?`);
      if (!ok) throw new Error('Загрузка отменена. Повторите попытку.');
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

    // SKU: выбираем заказ, у которого ещё есть недосканированные штуки (приоритет — текущая полка)
    const candidates = state.skuIndex[code];
    if (candidates && candidates.length) {
      const remaining = candidates.filter(id => (state.factGlobal[id] || 0) < state.ordersMap[id].planQty);
      if (remaining.length === 0) return { orderId: candidates[0], viaSku: true, full: true };
      const onCell = remaining.find(id => state.ordersMap[id].expectedCell === cell);
      return { orderId: onCell || remaining[0], viaSku: true };
    }

    for (const [knownCode, orderId] of Object.entries(state.codeToOrderId)) {
      if (knownCode.length >= 8 && code.startsWith(knownCode)) {
        return { orderId, viaSku: false };
      }
    }

    return null;
  }

  function looksLikeCell(code) {
    if (state.planCellUnits.hasOwnProperty(code)) return true;
    if (state.codeToOrderId[code] || state.skuIndex[code]) return false;
    return code.length <= 5 && !code.startsWith('1000') && !code.includes('-');
  }

  // очередь: сканы обрабатываются строго по порядку (сеть + быстрый сканер)
  let scanChain = Promise.resolve();
  function enqueueScan(code) {
    scanChain = scanChain.then(() => processScan(code)).catch(e => console.error('scan error', e));
  }

  async function processScan(code) {
    code = String(code || '').trim();
    if (!code) return;

    if (looksLikeCell(code)) {
      state.activeCell = code;
      if (!state.factCellUnits[code]) state.factCellUnits[code] = {};
      soundCell();
      updateScannerUI();
      return;
    }

    const cell = state.activeCell;
    if (!cell) {
      soundSurplus();
      alert('Сначала отсканируйте Штрихкод Ячейки (полки)!');
      return;
    }

    const resolved = resolveBarcode(code, cell);

    if (resolved) {
      const { orderId, viaSku, full } = resolved;
      const order = state.ordersMap[orderId];
      const alreadyFull = full || (state.factGlobal[orderId] || 0) >= order.planQty;

      if ((!viaSku && order.scannedCodes.has(code)) || alreadyFull) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', `ℹ️ Заказ №${order.pid || orderId} уже учтен полностью.`);
        return;
      }

      if (!state.factCellUnits[cell]) state.factCellUnits[cell] = {};
      if (!viaSku) order.scannedCodes.add(code);
      state.factCellUnits[cell][orderId] = (state.factCellUnits[cell][orderId] || 0) + 1;
      state.factGlobal[orderId] = (state.factGlobal[orderId] || 0) + 1;
      state.totalScannedUnits++;

      state.scannedHistory.push({ cell, code, type: 'ORDER', orderId });

      soundOk();
      updateScannerUI(code, 'OK', `✓ Заказ №${order.pid || orderId} (${order.customer?.firstName || ''}) принят.`);
      return;
    }

    updateScannerUI(code, 'SEARCH', '🔍 Проверка статуса в WMS...');
    const deepStatus = await deepLookupShipmentStatus(code);

    if (deepStatus && deepStatus.type === 'ACCEPTANCE') {
      if (!state.acceptanceScans[cell]) state.acceptanceScans[cell] = {};
      if (state.acceptanceScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', 'ℹ️ Эта позиция приёмки уже учтена.');
        return;
      }
      state.acceptanceScans[cell][code] = {
        orderId: deepStatus.orderId,
        cargoPlace: deepStatus.cargoPlace,
        statusText: deepStatus.statusText,
        count: 1
      };
      state.totalScannedUnits++;
      state.scannedHistory.push({ cell, code, type: 'ACCEPTANCE', meta: deepStatus });

      soundAcceptance();
      updateScannerUI(code, 'ACCEPTANCE', '⚠️ В ПРИЁМКЕ (Заказ не завершен). Оставить в ячейке! Завершить приёмку в WMS.');
      return;
    }

    if (deepStatus && deepStatus.type === 'RETURN') {
      if (!state.returnScans[cell]) state.returnScans[cell] = {};
      if (state.returnScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', 'ℹ️ Эта позиция возврата уже учтена.');
        return;
      }
      state.returnScans[cell][code] = {
        orderId: deepStatus.orderId,
        cargoPlace: deepStatus.cargoPlace,
        statusText: deepStatus.statusText,
        count: 1
      };
      state.totalScannedUnits++;
      state.scannedHistory.push({ cell, code, type: 'RETURN', meta: deepStatus });

      soundReturn();
      updateScannerUI(code, 'RETURN', '🔴 ОТМЕНЕН / ВОЗВРАТ! Срочно убрать с полки и отправить на склад.');
      return;
    }

    if (!state.foreignScans[cell]) state.foreignScans[cell] = {};
    state.foreignScans[cell][code] = (state.foreignScans[cell][code] || 0) + 1;
    state.totalScannedUnits++;
    state.scannedHistory.push({ cell, code, type: 'FOREIGN' });

    soundSurplus();
    updateScannerUI(code, 'FOREIGN', '⛔ Чужой товар/посылка! Не числится на этом ПВЗ. Отложить в возврат!');
  }

  function undoLastScan() {
    if (state.scannedHistory.length === 0) return;
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

    updateScannerUI('ОТМЕНЕН', 'INFO', `↩️ Скан "${last.code}" успешно отменен.`);
  }

  // ---------------------------------------------------------------- report
  function buildReport() {
    const report = {
      pureShortages: [],
      mismatchedCells: [],
      inAcceptance: [],
      canceledReturns: [],
      pureSurplus: [],
      totalShortageUnits: 0,
      totalMismatchUnits: 0,
      totalAcceptanceUnits: 0,
      totalReturnUnits: 0,
      totalSurplusUnits: 0
    };

    const skuOf = (orderId, order) => order.itemSkus.length > 0 ? order.itemSkus[0].sku : `10-${orderId.padStart(10, '0')}-1`;

    Object.entries(state.ordersMap).forEach(([orderId, order]) => {
      const factQty = state.factGlobal[orderId] || 0;
      if (factQty < order.planQty) {
        const diff = order.planQty - factQty;
        report.totalShortageUnits += diff;
        report.pureShortages.push({
          orderId,
          pid: order.pid,
          sku: skuOf(orderId, order),
          missingCount: diff,
          expectedCell: order.expectedCell,
          customer: safeCustomer(order.customer)
        });
      }
    });

    Object.entries(state.ordersMap).forEach(([orderId, order]) => {
      const expCell = order.expectedCell;
      if (!expCell || expCell === 'БЕЗ_ЯЧЕЙКИ' || expCell === 'null') return;

      Object.keys(state.factCellUnits).forEach(actualCell => {
        const qty = state.factCellUnits[actualCell][orderId];
        if (actualCell !== expCell && qty > 0) {
          report.totalMismatchUnits += qty;
          report.mismatchedCells.push({
            orderId,
            pid: order.pid,
            sku: skuOf(orderId, order),
            count: qty,
            expectedCell: expCell,
            actualCell,
            customer: safeCustomer(order.customer)
          });
        }
      });
    });

    Object.entries(state.acceptanceScans).forEach(([cell, items]) => {
      Object.entries(items).forEach(([code, data]) => {
        report.totalAcceptanceUnits += data.count;
        report.inAcceptance.push({ code, orderId: data.orderId, cargoPlace: data.cargoPlace, foundCell: cell, count: data.count, statusText: data.statusText });
      });
    });

    Object.entries(state.returnScans).forEach(([cell, items]) => {
      Object.entries(items).forEach(([code, data]) => {
        report.totalReturnUnits += data.count;
        report.canceledReturns.push({ code, orderId: data.orderId, cargoPlace: data.cargoPlace, foundCell: cell, count: data.count, statusText: data.statusText });
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

  // ---------------------------------------------------------------- xlsx
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
    sheets.forEach((_, idx) => ct += `<Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    ct += `</Types>`;
    files.push({ name: '[Content_Types].xml', data: ct });

    files.push({
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    });

    let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    sheets.forEach((_, idx) => wbRels += `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`);
    wbRels += `</Relationships>`;
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: wbRels });

    let wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`;
    sheets.forEach((s, idx) => wb += `<sheet name="${esc(s.name).slice(0, 31)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`);
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
    const DOS_DATE = 0x21; // 1980-01-01 (валидная дата)
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
      dv.setUint16(6, 0x0800, true); // UTF-8 имена
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

  function exportCleanXlsx(reportData = null, auditMeta = null) {
    const report = reportData || state.lastReport || buildReport();
    const now = new Date();
    const dateStr = auditMeta?.date || now.toLocaleDateString('ru-RU');
    const timeStr = auditMeta?.time || now.toLocaleTimeString('ru-RU');
    const pvzName = auditMeta?.pvzName || state.activePvzName;
    const dpKey = auditMeta?.dpKey || state.dpKey;
    const totalOrders = auditMeta?.totalOrders ?? state.totalOrdersCount;
    const totalPlanUnits = auditMeta?.totalPlanUnits ?? state.totalPlanUnits;
    const totalScanned = auditMeta?.totalScanned ?? state.totalScannedUnits;

    const fullName = c => `${c?.firstName || ''} ${c?.lastName || ''}`.trim() || '—';

    const allRows = [];

    allRows.push(['АКТ ИНВЕНТАРИЗАЦИИ СКЛАДА ПВЗ']);
    allRows.push(['Пункт выдачи:', `${pvzName} (ID: ${dpKey})`, '', 'Дата и время:', `${dateStr} ${timeStr}`]);
    allRows.push(['Территориальный управляющий:', 'Хожиев Сухроб', '', 'Статус проверки:', 'Завершена целиком (Сохранено в базу)']);
    allRows.push([]);

    allRows.push(['=== ИТОГОВАЯ СВОДКА РЕВИЗИИ ===']);
    allRows.push(['Показатель', 'Количество', 'Статус']);
    allRows.push(['Всего клиентских заказов в плане WMS:', totalOrders, 'План к выдаче']);
    allRows.push(['Товаров по плану:', totalPlanUnits, 'План штук']);
    allRows.push(['Всего отсканировано товаров:', totalScanned, 'Факт на складе']);
    allRows.push(['🔴 Чистая недостача (не найдены на складе):', report.pureShortages.length, `Не хватает: -${report.totalShortageUnits} шт`]);
    allRows.push(['🟡 Пересорт ячеек (лежат не на своих полках):', report.mismatchedCells.length, `${report.totalMismatchUnits} шт требуют перемещения`]);
    allRows.push(['⚠️ Заказы в приёмке (не приняты полностью):', report.inAcceptance.length, `${report.totalAcceptanceUnits} шт (Оставить в ячейке, провести приемку)`]);
    allRows.push(['🔴 Отмененные / Возвраты (к отгрузке на склад):', report.canceledReturns.length, `${report.totalReturnUnits} шт (Срочно снять с полки на склад)`]);
    allRows.push(['⛔ Чужие товары (не числятся за ПВЗ):', report.pureSurplus.length, `+${report.totalSurplusUnits} шт в коробку возвратов`]);
    allRows.push([]);

    allRows.push(['🔴 РАЗДЕЛ 1. НЕДОСТАЧИ (ЗАКАЗЫ, КОТОРЫХ НЕТ НА СКЛАДЕ)']);
    if (report.pureShortages.length === 0) {
      allRows.push(['✓ Недостач не обнаружено. Все клиентские заказы на месте!']);
    } else {
      allRows.push(['№', 'Номер заказа (PID)', 'ID в WMS', 'Штрихкод (SKU / Место)', 'Не хватает', 'Полка по плану', 'Где найден', 'Клиент', 'Телефон']);
      report.pureShortages.forEach((r, idx) => {
        allRows.push([
          idx + 1,
          r.pid || r.orderId,
          r.orderId,
          r.sku,
          `-${r.missingCount} шт`,
          r.expectedCell || 'БЕЗ_ЯЧЕЙКИ',
          'НЕ НАЙДЕН',
          fullName(r.customer),
          r.customer?.phone || '—'
        ]);
      });
    }
    allRows.push([]);

    allRows.push(['🟡 РАЗДЕЛ 2. ПЕРЕСОРТ ЯЧЕЕК (ТОВАР НА СКЛАДЕ, НО НЕ НА СВОЕЙ ПОЛКЕ)']);
    if (report.mismatchedCells.length === 0) {
      allRows.push(['✓ Все найденные заказы лежат строго на своих местах.']);
    } else {
      allRows.push(['№', 'Номер заказа (PID)', 'ID в WMS', 'Штрихкод (SKU / Место)', 'Кол-во', 'Полка по плану', 'Где лежит по факту', 'Клиент', 'Действие']);
      report.mismatchedCells.forEach((r, idx) => {
        allRows.push([
          idx + 1,
          r.pid || r.orderId,
          r.orderId,
          r.sku,
          `${r.count} шт`,
          r.expectedCell,
          r.actualCell,
          fullName(r.customer),
          `Переложить в ячейку ${r.expectedCell}`
        ]);
      });
    }
    allRows.push([]);

    allRows.push(['⚠️ РАЗДЕЛ 3. ТОВАРЫ В ПРИЁМКЕ (НЕ ПРИНЯТЫ ПОЛНОСТЬЮ — НЕ ОТПРАВЛЯТЬ НА СКЛАД!)']);
    if (report.inAcceptance.length === 0) {
      allRows.push(['✓ Товаров в статусе незавершенной приёмки не обнаружено.']);
    } else {
      allRows.push(['№', 'Считанный ШК / Место', 'ID заказа', 'Кол-во', 'Полка', 'Статус в WMS', '', '', 'Рекомендуемое действие']);
      report.inAcceptance.forEach((r, idx) => {
        allRows.push([idx + 1, r.code, r.orderId || '—', `+${r.count} шт`, `ячейка ${r.foundCell}`, r.statusText || 'В приёмке', '', '', 'Оставить в ячейке! Завершить приёмку в WMS']);
      });
    }
    allRows.push([]);

    allRows.push(['🔴 РАЗДЕЛ 4. ОТМЕНЕННЫЕ ЗАКАЗЫ И ВОЗВРАТЫ (СРОЧНО УБРАТЬ С ПОЛКИ И ОТПРАВИТЬ НА СКЛАД)']);
    if (report.canceledReturns.length === 0) {
      allRows.push(['✓ Забытых отмененных заказов или возвратов на полках не обнаружено.']);
    } else {
      allRows.push(['№', 'Считанный ШК / Место', 'ID заказа', 'Кол-во', 'Полка', 'Статус в WMS', '', '', 'Рекомендуемое действие']);
      report.canceledReturns.forEach((r, idx) => {
        allRows.push([idx + 1, r.code, r.orderId || '—', `+${r.count} шт`, `ячейка ${r.foundCell}`, r.statusText || 'Завершен (Отгрузка)', '', '', 'Снять с полки! Оформить отгрузку на возвратный склад']);
      });
    }
    allRows.push([]);

    allRows.push(['⛔ РАЗДЕЛ 5. ЧУЖИЕ ТОВАРЫ И НЕОПОЗНАННЫЕ ИЗЛИШКИ']);
    if (report.pureSurplus.length === 0) {
      allRows.push(['✓ Лишних или чужих посылок на складе не обнаружено.']);
    } else {
      allRows.push(['№', 'Считанный ШК / Место', 'Количество', 'Где найден (Полка)', 'Статус', '', '', '', 'Рекомендуемое действие']);
      report.pureSurplus.forEach((r, idx) => {
        allRows.push([idx + 1, r.code, `+${r.qty} шт`, `ячейка ${r.foundCell}`, 'Чужой ПВЗ / Не числится', '', '', '', 'Отложить в коробку возвратов']);
      });
    }

    const colWidths = [6, 22, 14, 24, 14, 18, 18, 28, 22];
    const blob = createRealXlsxBlob([{ name: 'Акт ревизии', rows: allRows, cols: colWidths }]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeFileName(`Акт_ревизии_${pvzName}_${String(dateStr).replace(/\./g, '-')}.xlsx`);
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
    // при нехватке места удаляем самые старые записи и пробуем снова
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
    const report = buildReport();
    const now = new Date();

    const auditEntry = {
      id: `INV-${Date.now()}`,
      date: now.toLocaleDateString('ru-RU'),
      time: now.toLocaleTimeString('ru-RU'),
      pvzName: state.activePvzName,
      dpKey: state.dpKey,
      auditor: 'Хожиев Сухроб',
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
        surplusUnits: report.totalSurplusUnits
      },
      report
    };

    const list = getHistoryList();
    list.unshift(auditEntry);
    while (list.length > 100) list.pop();
    const saved = writeHistory(list);

    soundSuccess();
    if (saved) {
      alert(`✓ Ревизия филиала ${state.activePvzName} успешно зафиксирована и сохранена в базу!\nАкт доступен во вкладке "История ревизий".`);
    } else {
      alert('⚠️ Не удалось сохранить ревизию в базу браузера (нет места). Акт .xlsx будет скачан.');
    }
    exportCleanXlsx();
  }

  // ---------------------------------------------------------------- UI
  const CSS = `
#uzum-inv-btn{display:inline-flex;align-items:center;gap:8px;height:38px;padding:0 18px;margin-right:12px;background:linear-gradient(135deg,#7c3aed,#5b21b6);color:#fff;font:700 13px/1 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;border:0;border-radius:999px;cursor:pointer;box-shadow:0 4px 14px rgba(109,40,217,.35);white-space:nowrap;transition:transform .15s,box-shadow .15s;vertical-align:middle;outline:none}
#uzum-inv-btn:hover{transform:translateY(-1px);box-shadow:0 6px 18px rgba(109,40,217,.45)}
#uzum-inv-btn.uz-fixed{position:fixed;top:12px;right:400px;z-index:999999;margin:0}
#uzum-inv-modal{--p:#7c3aed;--pd:#5b21b6;--pl:#f3edff;--ink:#1f1a33;--mut:#6f6a85;--line:#e9e4f5;--bg:#faf8ff;--ok:#059669;--okl:#e7f8f1;--warn:#d97706;--warnl:#fff5e0;--acc:#7c3aed;--accl:#f3edff;--ret:#ea580c;--retl:#fff0e6;--bad:#dc2626;--badl:#fdecec;display:none;position:fixed;inset:0;background:rgba(24,16,48,.72);backdrop-filter:blur(4px);z-index:1000000;align-items:center;justify-content:center;font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:var(--ink);font-size:14px}
#uzum-inv-modal *{box-sizing:border-box}
#uzum-inv-modal .uz-shell{width:980px;max-width:96vw;height:92vh;background:var(--bg);border-radius:22px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 30px 70px rgba(20,8,50,.55)}
#uzum-inv-modal .uz-head{background:linear-gradient(120deg,#3b1480,#7c3aed);color:#fff;padding:16px 24px;display:flex;align-items:center;gap:14px}
#uzum-inv-modal .uz-logo{width:40px;height:40px;border-radius:12px;background:#fff;color:var(--pd);display:flex;align-items:center;justify-content:center;font-weight:900;font-size:22px}
#uzum-inv-modal .uz-title{font-size:18px;font-weight:800;letter-spacing:.2px}
#uzum-inv-modal .uz-sub{font-size:11.5px;opacity:.85;margin-top:2px}
#uzum-inv-modal .uz-ver{margin-left:auto;background:rgba(255,255,255,.18);padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700}
#uzum-inv-modal .uz-x{background:rgba(255,255,255,.14);border:0;color:#fff;width:34px;height:34px;border-radius:50%;font-size:16px;cursor:pointer}
#uzum-inv-modal .uz-x:hover{background:rgba(255,255,255,.28)}
#uzum-inv-modal .uz-tabs{display:flex;gap:6px;padding:12px 24px 0;background:#fff;border-bottom:1px solid var(--line)}
#uzum-inv-modal .uz-tab{padding:10px 18px;border:0;background:transparent;font-weight:700;font-size:13px;color:var(--mut);cursor:pointer;border-bottom:3px solid transparent;border-radius:8px 8px 0 0}
#uzum-inv-modal .uz-tab:hover{color:var(--pd)}
#uzum-inv-modal .uz-tab.active{color:var(--pd);border-bottom-color:var(--p);background:var(--pl)}
#uzum-inv-modal .uz-body{padding:22px 24px;flex:1;overflow-y:auto}
#uzum-inv-modal .uz-card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px 18px}
#uzum-inv-modal .uz-lbl{font-size:11px;font-weight:800;color:var(--mut);letter-spacing:.6px;text-transform:uppercase}
#uzum-inv-modal .uz-btn{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border:1px solid var(--line);background:#fff;color:var(--ink);font-weight:700;font-size:13px;border-radius:12px;cursor:pointer;transition:all .15s}
#uzum-inv-modal .uz-btn:hover{border-color:var(--p);color:var(--pd)}
#uzum-inv-modal .uz-btn.pri{background:linear-gradient(135deg,#7c3aed,#5b21b6);color:#fff;border-color:transparent;box-shadow:0 6px 16px rgba(109,40,217,.3)}
#uzum-inv-modal .uz-btn.pri:hover{filter:brightness(1.08);color:#fff}
#uzum-inv-modal .uz-btn.ok{background:var(--ok);color:#fff;border-color:transparent}
#uzum-inv-modal .uz-btn.ok:hover{filter:brightness(1.08);color:#fff}
#uzum-inv-modal .uz-btn.del{background:var(--badl);color:var(--bad);border-color:transparent;padding:10px 12px}
#uzum-inv-modal .uz-btn.big{padding:15px 46px;font-size:16px;border-radius:14px}
#uzum-inv-modal .uz-btn:disabled{opacity:.6;cursor:default}
#uzum-inv-modal .uz-in{padding:10px 14px;border:2px solid var(--line);border-radius:12px;font-size:16px;font-weight:700;text-align:center;outline:none;background:#fff;color:var(--ink)}
#uzum-inv-modal .uz-in:focus{border-color:var(--p);box-shadow:0 0 0 4px rgba(124,58,237,.14)}
#uzum-inv-modal .uz-start{max-width:520px;margin:20px auto;text-align:center}
#uzum-inv-modal .uz-start .uz-card{padding:26px}
#uzum-inv-modal .uz-row{display:flex;gap:14px;justify-content:center;margin:14px 0}
#uzum-inv-modal .uz-orders{font-weight:700;color:var(--pd);background:var(--pl);display:inline-block;padding:7px 16px;border-radius:999px;font-size:13px}
#uzum-inv-modal .uz-prog{margin-top:16px;color:var(--pd);font-weight:600;min-height:20px}
#uzum-inv-modal .uz-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px}
#uzum-inv-modal .uz-stat{background:#fff;border:1px solid var(--line);border-radius:14px;padding:12px 16px}
#uzum-inv-modal .uz-stat b{display:block;font-size:22px;font-weight:800;margin-top:2px}
#uzum-inv-modal .uz-bar{height:10px;background:var(--line);border-radius:999px;overflow:hidden;margin-bottom:6px}
#uzum-inv-modal .uz-bar>div{height:100%;width:0;background:linear-gradient(90deg,#a78bfa,#7c3aed);border-radius:999px;transition:width .25s}
#uzum-inv-modal .uz-barlbl{font-size:12px;color:var(--mut);margin-bottom:14px;font-weight:600}
#uzum-inv-modal .uz-duo{display:grid;grid-template-columns:1fr 1.3fr;gap:12px;margin-bottom:12px}
#uzum-inv-modal .uz-cell{font-size:42px;font-weight:900;color:var(--pd);line-height:1.1;margin-top:6px;word-break:break-all}
#uzum-inv-modal .uz-cell.empty{font-size:20px;color:var(--mut);font-weight:700;padding:10px 0}
#uzum-inv-modal .uz-last{font-size:16px;font-weight:800;margin:4px 0 8px;word-break:break-all;font-family:ui-monospace,Consolas,monospace}
#uzum-inv-modal .uz-banner{padding:10px 14px;border-radius:12px;font-weight:700;font-size:13.5px;border-left:6px solid var(--mut);background:#f1eff7;color:var(--mut)}
#uzum-inv-modal .uz-banner.ok{background:var(--okl);color:var(--ok);border-color:var(--ok)}
#uzum-inv-modal .uz-banner.info{background:var(--warnl);color:var(--warn);border-color:var(--warn)}
#uzum-inv-modal .uz-banner.acc{background:var(--accl);color:var(--acc);border-color:var(--acc)}
#uzum-inv-modal .uz-banner.ret{background:var(--retl);color:var(--ret);border-color:var(--ret)}
#uzum-inv-modal .uz-banner.bad{background:var(--badl);color:var(--bad);border-color:var(--bad)}
#uzum-inv-modal .uz-banner.search{background:var(--pl);color:var(--pd);border-color:var(--p)}
#uzum-inv-modal .uz-scanrow{display:flex;gap:10px;margin-bottom:14px}
#uzum-inv-modal .uz-scanrow input{flex:1;padding:14px 16px;border:2px solid var(--p);border-radius:14px;font-size:15px;outline:none;background:#fff}
#uzum-inv-modal .uz-scanrow input:focus{box-shadow:0 0 0 4px rgba(124,58,237,.16)}
#uzum-inv-modal .uz-listhead{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
#uzum-inv-modal .uz-listhead span{font-weight:800;font-size:14px}
#uzum-inv-modal .uz-items{background:#fff;border:1px solid var(--line);border-radius:14px;max-height:230px;overflow-y:auto}
#uzum-inv-modal .uz-empty{padding:20px;text-align:center;color:var(--mut)}
#uzum-inv-modal .uz-tbl{width:100%;border-collapse:collapse;font-size:12.5px}
#uzum-inv-modal .uz-tbl th{text-align:left;padding:9px 12px;background:#f6f3fd;color:var(--mut);font-size:11px;text-transform:uppercase;letter-spacing:.5px}
#uzum-inv-modal .uz-tbl td{padding:9px 12px;border-top:1px solid var(--line)}
#uzum-inv-modal .uz-tbl .r{text-align:right;font-weight:800}
#uzum-inv-modal .uz-tbl .mono{font-family:ui-monospace,Consolas,monospace;font-weight:700}
#uzum-inv-modal .uz-tbl tr.acc td{background:#faf7ff}
#uzum-inv-modal .uz-tbl tr.ret td{background:#fffaf5}
#uzum-inv-modal .uz-tbl tr.bad td{background:#fff8f8}
#uzum-inv-modal .uz-tag{font-size:10.5px;font-weight:800;padding:2px 8px;border-radius:999px;margin-left:6px}
#uzum-inv-modal .uz-tag.acc{background:var(--accl);color:var(--acc)}
#uzum-inv-modal .uz-tag.ret{background:var(--retl);color:var(--ret)}
#uzum-inv-modal .uz-tag.bad{background:var(--badl);color:var(--bad)}
#uzum-inv-modal .uz-rhead{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap}
#uzum-inv-modal .uz-rhead h3{margin:0;font-size:19px}
#uzum-inv-modal .uz-rmeta{font-size:12.5px;color:var(--mut);margin-top:3px}
#uzum-inv-modal .uz-acts{display:flex;gap:8px;flex-wrap:wrap}
#uzum-inv-modal .uz-tiles{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:18px}
#uzum-inv-modal .uz-tile{border-radius:14px;padding:12px 14px;border:1px solid var(--line);background:#fff}
#uzum-inv-modal .uz-tile .n{font-size:26px;font-weight:900;line-height:1.1}
#uzum-inv-modal .uz-tile .t{font-size:11.5px;font-weight:700;color:var(--mut);margin-top:2px}
#uzum-inv-modal .uz-tile .u{font-size:11.5px;font-weight:800;margin-top:4px}
#uzum-inv-modal .uz-tile.bad{background:var(--badl);color:var(--bad)}
#uzum-inv-modal .uz-tile.warn{background:var(--warnl);color:var(--warn)}
#uzum-inv-modal .uz-tile.acc{background:var(--accl);color:var(--acc)}
#uzum-inv-modal .uz-tile.ret{background:var(--retl);color:var(--ret)}
#uzum-inv-modal .uz-tile.mut{background:#f1eff7;color:var(--mut)}
#uzum-inv-modal .uz-sec{background:#fff;border:1px solid var(--line);border-radius:16px;overflow:hidden;margin-bottom:14px}
#uzum-inv-modal .uz-sec-h{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;font-weight:800;font-size:14px}
#uzum-inv-modal .uz-sec.bad .uz-sec-h{background:var(--badl);color:var(--bad)}
#uzum-inv-modal .uz-sec.warn .uz-sec-h{background:var(--warnl);color:var(--warn)}
#uzum-inv-modal .uz-sec.acc .uz-sec-h{background:var(--accl);color:var(--acc)}
#uzum-inv-modal .uz-sec.ret .uz-sec-h{background:var(--retl);color:var(--ret)}
#uzum-inv-modal .uz-sec.mut .uz-sec-h{background:#f1eff7;color:var(--mut)}
#uzum-inv-modal .uz-pill{background:rgba(255,255,255,.75);padding:3px 11px;border-radius:999px;font-size:12px}
#uzum-inv-modal .uz-none{padding:14px 16px;color:var(--mut);font-size:13px}
#uzum-inv-modal .uz-none.ok{color:var(--ok);font-weight:700}
#uzum-inv-modal .uz-hhead{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;gap:12px}
#uzum-inv-modal .uz-hhead h3{margin:0;font-size:18px}
#uzum-inv-modal .uz-hhead input{padding:10px 14px;border:2px solid var(--line);border-radius:12px;width:290px;font-size:13px;outline:none}
#uzum-inv-modal .uz-hhead input:focus{border-color:var(--p)}
#uzum-inv-modal .uz-hist{display:flex;justify-content:space-between;align-items:center;gap:14px;background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px 18px;margin-bottom:12px}
#uzum-inv-modal .uz-hist .h1{font-weight:800;font-size:15px}
#uzum-inv-modal .uz-hist .h1 span{color:var(--mut);font-weight:500;font-size:13px}
#uzum-inv-modal .uz-chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
#uzum-inv-modal .uz-chip{font-size:11.5px;font-weight:700;padding:3px 10px;border-radius:999px;background:#f1eff7;color:var(--mut)}
#uzum-inv-modal .uz-chip.bad{background:var(--badl);color:var(--bad)}
#uzum-inv-modal .uz-chip.warn{background:var(--warnl);color:var(--warn)}
#uzum-inv-modal .uz-chip.acc{background:var(--accl);color:var(--acc)}
#uzum-inv-modal .uz-chip.ret{background:var(--retl);color:var(--ret)}
@media (max-width:760px){#uzum-inv-modal .uz-tiles{grid-template-columns:repeat(2,1fr)}#uzum-inv-modal .uz-duo{grid-template-columns:1fr}}
`;

  function ensureStyle() {
    if ($('uzum-inv-style')) return;
    const st = document.createElement('style');
    st.id = 'uzum-inv-style';
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  function showScreen(name) {
    state.screen = name;
    if (name !== 'history') state.workScreen = name;
    ['start', 'scan', 'report', 'history'].forEach(n => {
      const el = $(`u-${n}-screen`);
      if (el) el.style.display = n === name ? 'block' : 'none';
    });
    $('u-tab-new').classList.toggle('active', name !== 'history');
    $('u-tab-history').classList.toggle('active', name === 'history');
    if (name === 'scan') {
      const inp = $('u-barcode-input');
      if (inp) inp.focus();
    }
    if (name === 'history') renderHistoryScreen();
  }

  function isModalOpen() {
    const m = $('uzum-inv-modal');
    return !!(m && m.style.display !== 'none');
  }

  function isScanningActive() {
    return isModalOpen() && state.screen === 'scan';
  }

  function ensureModal() {
    ensureStyle();
    if ($('uzum-inv-modal')) return;

    const modal = document.createElement('div');
    modal.id = 'uzum-inv-modal';
    modal.innerHTML = `
      <div class="uz-shell">
        <div class="uz-head">
          <div class="uz-logo">U</div>
          <div>
            <div class="uz-title">Ревизор ПВЗ</div>
            <div class="uz-sub" id="u-token-badge">Приёмка · Возвраты · Пересорт · История ревизий</div>
          </div>
          <span class="uz-ver">v${VERSION}</span>
          <button id="u-close" class="uz-x" title="Закрыть">✕</button>
        </div>

        <div class="uz-tabs">
          <button id="u-tab-new" class="uz-tab active">📋 Новая ревизия</button>
          <button id="u-tab-history" class="uz-tab">🗄️ История ревизий</button>
        </div>

        <div class="uz-body">
          <div id="u-start-screen" class="uz-start">
            <div class="uz-card">
              <div class="uz-lbl">Параметры ревизии</div>
              <div class="uz-row">
                <div>
                  <div class="uz-lbl" style="margin-bottom:6px">ПВЗ</div>
                  <input type="text" id="u-pvz-name-input" class="uz-in" style="width:140px">
                </div>
                <div>
                  <div class="uz-lbl" style="margin-bottom:6px">ID (dpKey)</div>
                  <input type="text" id="u-pvz-dpkey-input" class="uz-in" style="width:120px">
                </div>
              </div>
              <div id="u-pvz-display-orders" class="uz-orders">Заказов к выдаче: —</div>
              <div style="margin-top:22px"><button id="u-load-btn" class="uz-btn pri big">Выгрузить остатки и начать</button></div>
              <div id="u-progress" class="uz-prog"></div>
            </div>
          </div>

          <div id="u-scan-screen" style="display:none;">
            <div class="uz-stats">
              <div class="uz-stat"><div class="uz-lbl">Заказов к выдаче</div><b id="u-sc-orders">0</b></div>
              <div class="uz-stat"><div class="uz-lbl">Товаров по плану</div><b id="u-sc-plan">0 шт</b></div>
              <div class="uz-stat"><div class="uz-lbl">Отсканировано</div><b id="u-sc-total" style="color:var(--p)">0 шт</b></div>
            </div>
            <div class="uz-bar"><div id="u-bar-fill"></div></div>
            <div class="uz-barlbl" id="u-bar-label">Найдено 0 из 0 шт · 0%</div>

            <div class="uz-duo">
              <div class="uz-card">
                <div class="uz-lbl">Активная ячейка</div>
                <div id="u-active-cell" class="uz-cell empty">Отсканируйте ячейку</div>
              </div>
              <div class="uz-card">
                <div class="uz-lbl">Последний штрихкод</div>
                <div id="u-last-barcode" class="uz-last">—</div>
                <div id="u-scan-hint" class="uz-banner">Сначала отсканируйте штрихкод ячейки</div>
              </div>
            </div>

            <div class="uz-scanrow">
              <input type="text" id="u-barcode-input" autocomplete="off" placeholder="Сканер активен — направьте луч на штрихкод">
              <button id="u-undo-btn" class="uz-btn">↩️ Отменить скан</button>
            </div>

            <div class="uz-listhead">
              <span>Учтено в текущей ячейке</span>
              <button id="u-finish-btn" class="uz-btn ok">Итоги ревизии →</button>
            </div>
            <div id="u-cell-items" class="uz-items"></div>
          </div>

          <div id="u-report-screen" style="display:none;"></div>

          <div id="u-history-screen" style="display:none;">
            <div class="uz-hhead">
              <h3>Архив ревизий</h3>
              <input type="text" id="u-hist-search" placeholder="🔍 Поиск по заказу (PID) или штрихкоду">
            </div>
            <div id="u-hist-container"></div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(modal);

    $('u-close').onclick = () => { modal.style.display = 'none'; };
    $('u-undo-btn').onclick = () => { undoLastScan(); $('u-barcode-input').focus(); };
    $('u-tab-new').onclick = () => showScreen(state.workScreen || 'start');
    $('u-tab-history').onclick = () => showScreen('history');
    $('u-hist-search').oninput = (e) => renderHistoryScreen(e.target.value.trim());

    $('u-load-btn').onclick = async () => {
      const btn = $('u-load-btn');
      if (btn.disabled) return;
      state.activePvzName = $('u-pvz-name-input').value.trim() || state.activePvzName;
      state.dpKey = $('u-pvz-dpkey-input').value.trim() || state.dpKey;

      const prog = $('u-progress');
      btn.disabled = true;
      try {
        await loadInventoryData(msg => { prog.textContent = msg; });
        prog.textContent = '';
        showScreen('scan');
        updateScannerUI();
      } catch (err) {
        prog.textContent = '';
        alert(err.message);
      } finally {
        btn.disabled = false;
      }
    };

    $('u-finish-btn').onclick = () => renderReportScreen();

    // перехват клавиатуры только пока открыта модалка и активен экран сканирования
    let scanBuffer = '';
    let scanTimer = null;

    const captureKey = (e) => {
      if (!isScanningActive()) return;
      e.stopPropagation();
      e.stopImmediatePropagation();

      const input = $('u-barcode-input');
      if (input && document.activeElement !== input) input.focus();

      if (e.type === 'keydown') {
        if (e.key === 'Enter') {
          e.preventDefault();
          const code = ((input ? input.value : '') || scanBuffer).trim();
          if (input) input.value = '';
          scanBuffer = '';
          if (code) enqueueScan(code);
        } else if (e.key.length === 1) {
          scanBuffer += e.key;
          clearTimeout(scanTimer);
          scanTimer = setTimeout(() => { scanBuffer = ''; }, 120);
        }
      }
    };

    window.addEventListener('keydown', captureKey, true);
    window.addEventListener('keypress', captureKey, true);
    window.addEventListener('keyup', captureKey, true);
  }

  // кнопка в шапке: ищем место редко, не каждую секунду
  let btnFallbackTicks = 0;
  function ensureButton() {
    ensureStyle();
    let btn = $('uzum-inv-btn');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'uzum-inv-btn';
      btn.innerHTML = '📦 Ревизия ПВЗ';
      btn.onclick = openModal;
    }

    if (btn.isConnected && btn.dataset.mode === 'header') return;
    if (btn.isConnected && btn.dataset.mode === 'fallback') {
      btnFallbackTicks++;
      if (btnFallbackTicks % 5 !== 0) return;
    }

    const dropdown = findPvzDropdown();
    if (dropdown && dropdown.parentElement) {
      btn.classList.remove('uz-fixed');
      btn.dataset.mode = 'header';
      dropdown.parentElement.insertBefore(btn, dropdown);
    } else if (!btn.isConnected) {
      btn.classList.add('uz-fixed');
      btn.dataset.mode = 'fallback';
      document.body.appendChild(btn);
    }
  }

  function injectUI() {
    ensureModal();
    ensureButton();
  }

  function openModal() {
    ensureModal();
    $('uzum-inv-modal').style.display = 'flex';

    if (state.screen === 'start' || !state.workScreen || state.workScreen === 'start') {
      const detectedName = detectPvzNameFromHeader();
      const liveDpKey = detectLiveDpKey();
      const ordersCount = detectOrdersCount();

      if (detectedName !== 'ПВЗ') state.activePvzName = detectedName;
      state.dpKey = liveDpKey || state.dpKey || '';

      $('u-pvz-name-input').value = state.activePvzName;
      $('u-pvz-dpkey-input').value = state.dpKey || '';
      $('u-pvz-display-orders').textContent = `Заказов к выдаче: ${ordersCount} шт`;
      showScreen('start');
    } else {
      showScreen(state.workScreen);
    }
  }

  function updateScannerUI(lastBarcode = '—', statusType = 'NORMAL', hintText = '') {
    if (!$('u-active-cell')) return;

    const cellEl = $('u-active-cell');
    cellEl.textContent = state.activeCell || 'Отсканируйте ячейку';
    cellEl.classList.toggle('empty', !state.activeCell);
    $('u-last-barcode').textContent = lastBarcode;

    $('u-sc-orders').textContent = state.totalOrdersCount;
    $('u-sc-plan').textContent = `${state.totalPlanUnits} шт`;
    $('u-sc-total').textContent = `${state.totalScannedUnits} шт`;

    const found = Object.values(state.factGlobal).reduce((a, b) => a + b, 0);
    const pct = state.totalPlanUnits ? Math.min(100, Math.round(found / state.totalPlanUnits * 100)) : 0;
    $('u-bar-fill').style.width = pct + '%';
    $('u-bar-label').textContent = `Найдено ${found} из ${state.totalPlanUnits} шт · ${pct}%`;

    const classes = { OK: 'ok', INFO: 'info', ACCEPTANCE: 'acc', RETURN: 'ret', FOREIGN: 'bad', SEARCH: 'search' };
    const hintEl = $('u-scan-hint');
    hintEl.className = 'uz-banner ' + (classes[statusType] || '');
    hintEl.textContent = hintText || (state.activeCell ? 'Сканируйте товары на этой полке' : 'Сначала отсканируйте штрихкод ячейки');

    const container = $('u-cell-items');
    const orderItems = state.factCellUnits[state.activeCell] || {};
    const accItems = state.acceptanceScans[state.activeCell] || {};
    const retItems = state.returnScans[state.activeCell] || {};
    const forItems = state.foreignScans[state.activeCell] || {};

    const hasAny = Object.keys(orderItems).length + Object.keys(accItems).length + Object.keys(retItems).length + Object.keys(forItems).length > 0;

    if (!hasAny) {
      container.innerHTML = '<div class="uz-empty">Ячейка пуста. Сканируйте заказы на этой полке.</div>';
      return;
    }

    let rows = '';
    Object.entries(orderItems).forEach(([oId, qty]) => {
      const order = state.ordersMap[oId];
      rows += `<tr><td><b>Заказ №${escHtml(order?.pid || oId)}</b> <span style="color:var(--mut);font-size:11.5px">${escHtml(order?.customer?.firstName || '')}</span></td><td class="r" style="color:var(--ok)">${qty} шт</td></tr>`;
    });
    Object.entries(accItems).forEach(([code, data]) => {
      rows += `<tr class="acc"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag acc">В ПРИЁМКЕ</span></td><td class="r" style="color:var(--acc)">+${data.count} шт</td></tr>`;
    });
    Object.entries(retItems).forEach(([code, data]) => {
      rows += `<tr class="ret"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag ret">ОТМЕНЁН / ВОЗВРАТ</span></td><td class="r" style="color:var(--ret)">+${data.count} шт</td></tr>`;
    });
    Object.entries(forItems).forEach(([code, qty]) => {
      rows += `<tr class="bad"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag bad">ЧУЖОЙ ПВЗ</span></td><td class="r" style="color:var(--bad)">+${qty} шт</td></tr>`;
    });

    container.innerHTML = `<table class="uz-tbl"><thead><tr><th>Заказ / Штрихкод</th><th style="text-align:right">Количество</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  const uzTable = (heads, rows) =>
    `<table class="uz-tbl"><thead><tr>${heads.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>`;
  const uzNone = (t, ok) => `<div class="uz-none ${ok ? 'ok' : ''}">${t}</div>`;
  const uzSection = (cls, title, count, units, body) =>
    `<div class="uz-sec ${cls}"><div class="uz-sec-h"><span>${title}</span><span class="uz-pill">${count} поз. · ${units}</span></div>${body}</div>`;

  function renderReportScreen() {
    const report = buildReport();
    const container = $('u-report-screen');

    const tiles = [
      { cls: 'bad', t: 'Недостачи', n: report.pureShortages.length, u: `-${report.totalShortageUnits} шт` },
      { cls: 'warn', t: 'Пересорт', n: report.mismatchedCells.length, u: `${report.totalMismatchUnits} шт` },
      { cls: 'acc', t: 'В приёмке', n: report.inAcceptance.length, u: `${report.totalAcceptanceUnits} шт` },
      { cls: 'ret', t: 'Возвраты', n: report.canceledReturns.length, u: `${report.totalReturnUnits} шт` },
      { cls: 'mut', t: 'Чужие', n: report.pureSurplus.length, u: `+${report.totalSurplusUnits} шт` }
    ];

    let html = `
      <div class="uz-rhead">
        <div>
          <h3>Итоги ревизии · ${escHtml(state.activePvzName)}</h3>
          <div class="uz-rmeta">Заказов: <b>${state.totalOrdersCount}</b> · По плану: <b>${state.totalPlanUnits} шт</b> · Отсканировано: <b>${state.totalScannedUnits} шт</b></div>
        </div>
        <div class="uz-acts">
          <button id="u-save-audit-btn" class="uz-btn ok">💾 Завершить целиком</button>
          <button id="u-excel-btn" class="uz-btn pri">📊 Скачать .xlsx</button>
          <button id="u-resume-btn" class="uz-btn">← Вернуться</button>
        </div>
      </div>
      <div class="uz-tiles">${tiles.map(x => `<div class="uz-tile ${x.cls}"><div class="n">${x.n}</div><div class="t">${x.t}</div><div class="u">${x.u}</div></div>`).join('')}</div>`;

    // 1. Недостачи
    html += uzSection('bad', '🔴 Недостачи', report.pureShortages.length, `-${report.totalShortageUnits} шт`,
      report.pureShortages.length === 0
        ? uzNone('✓ Недостач не обнаружено. Все клиентские заказы на месте!', true)
        : uzTable(['Заказ (PID)', 'Не хватает', 'Полка', 'Клиент'], report.pureShortages.map(r => `<tr>
            <td><b>№${escHtml(r.pid || r.orderId)}</b></td>
            <td style="color:var(--bad);font-weight:800">-${r.missingCount} шт</td>
            <td>${escHtml(r.expectedCell)}</td>
            <td>${escHtml(r.customer?.firstName || '')} ${r.customer?.phone ? '(' + escHtml(r.customer.phone) + ')' : ''}</td></tr>`).join('')));

    // 2. Пересорт
    html += uzSection('warn', '🟡 Пересорт ячеек', report.mismatchedCells.length, `${report.totalMismatchUnits} шт`,
      report.mismatchedCells.length === 0
        ? uzNone('✓ Все заказы лежат строго на своих местах.', true)
        : uzTable(['Заказ (PID)', 'Кол-во', 'По плану', 'Лежит фактически'], report.mismatchedCells.map(r => `<tr>
            <td><b>${escHtml(r.pid || r.orderId)}</b></td>
            <td style="font-weight:800">${r.count} шт</td>
            <td style="color:var(--bad);font-weight:800">${escHtml(r.expectedCell)}</td>
            <td style="color:var(--ok);font-weight:800">${escHtml(r.actualCell)}</td></tr>`).join('')));

    // 3. Приёмка
    html += uzSection('acc', '⚠️ В приёмке (не приняты полностью)', report.inAcceptance.length, `${report.totalAcceptanceUnits} шт`,
      report.inAcceptance.length === 0
        ? uzNone('Позиций в статусе приёмки на полках нет.', false)
        : uzTable(['ШК / Грузоместо', 'Кол-во', 'Полка', 'Рекомендация'], report.inAcceptance.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--acc);font-weight:800">+${r.count} шт</td>
            <td>ячейка ${escHtml(r.foundCell)}</td>
            <td style="font-weight:700">Не отправлять на склад! Оставить в ячейке, провести приёмку.</td></tr>`).join('')));

    // 4. Возвраты
    html += uzSection('ret', '🔴 Отменённые / Возвраты', report.canceledReturns.length, `${report.totalReturnUnits} шт`,
      report.canceledReturns.length === 0
        ? uzNone('Отменённых заказов на полках не обнаружено.', false)
        : uzTable(['ШК / Грузоместо', 'Кол-во', 'Полка', 'Рекомендация'], report.canceledReturns.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--ret);font-weight:800">+${r.count} шт</td>
            <td>ячейка ${escHtml(r.foundCell)}</td>
            <td style="font-weight:700">Снять с полки! Оформить отгрузку на возвратный склад.</td></tr>`).join('')));

    // 5. Чужие
    html += uzSection('mut', '⛔ Чужие товары / не числятся', report.pureSurplus.length, `+${report.totalSurplusUnits} шт`,
      report.pureSurplus.length === 0
        ? uzNone('Чужих товаров на складе не обнаружено.', false)
        : uzTable(['Штрихкод', 'Лишнее кол-во', 'Где найден'], report.pureSurplus.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--bad);font-weight:800">+${r.qty} шт</td>
            <td>ячейка ${escHtml(r.foundCell)}</td></tr>`).join('')));

    container.innerHTML = html;
    showScreen('report');

    $('u-save-audit-btn').onclick = saveAuditPermanently;
    $('u-excel-btn').onclick = () => exportCleanXlsx();
    $('u-resume-btn').onclick = () => showScreen('scan');
  }

  function renderHistoryScreen(searchFilter = '') {
    const container = $('u-hist-container');
    const list = getHistoryList();

    if (list.length === 0) {
      container.innerHTML = '<div class="uz-empty">В базе пока нет сохранённых ревизий. После проверки нажмите «Завершить целиком».</div>';
      return;
    }

    let filtered = list;
    if (searchFilter) {
      const q = searchFilter.toLowerCase();
      filtered = list.filter(item => JSON.stringify(item).toLowerCase().includes(q));
    }

    if (filtered.length === 0) {
      container.innerHTML = '<div class="uz-empty">Ничего не найдено.</div>';
      return;
    }

    let html = '';
    filtered.forEach(item => {
      const s = item.stats || {};
      html += `
        <div class="uz-hist">
          <div>
            <div class="h1">${escHtml(item.pvzName)} <span>· ID ${escHtml(item.dpKey)} · ${escHtml(item.date)} ${escHtml(item.time)}</span></div>
            <div class="uz-chips">
              <span class="uz-chip">Заказов ${s.totalOrders ?? 0}</span>
              <span class="uz-chip">План ${s.totalPlanUnits ?? 0}</span>
              <span class="uz-chip">Факт ${s.totalScanned ?? 0}</span>
              <span class="uz-chip bad">Недостач ${s.shortageCount ?? 0}</span>
              <span class="uz-chip warn">Пересорт ${s.mismatchCount ?? 0}</span>
              <span class="uz-chip acc">Приёмка ${s.acceptanceCount ?? 0}</span>
              <span class="uz-chip ret">Возвраты ${s.returnCount ?? 0}</span>
            </div>
          </div>
          <div class="uz-acts">
            <button class="uz-btn pri u-download-past-btn" data-id="${escHtml(item.id)}">📥 .xlsx</button>
            <button class="uz-btn del u-delete-past-btn" data-id="${escHtml(item.id)}" title="Удалить из базы">🗑</button>
          </div>
        </div>`;
    });

    container.innerHTML = html;

    container.querySelectorAll('.u-download-past-btn').forEach(b => {
      b.onclick = () => {
        const found = list.find(x => x.id === b.getAttribute('data-id'));
        if (found) {
          exportCleanXlsx(found.report, {
            date: found.date,
            time: found.time,
            pvzName: found.pvzName,
            dpKey: found.dpKey,
            totalOrders: found.stats.totalOrders,
            totalPlanUnits: found.stats.totalPlanUnits,
            totalScanned: found.stats.totalScanned
          });
        }
      };
    });

    container.querySelectorAll('.u-delete-past-btn').forEach(b => {
      b.onclick = () => {
        if (!confirm('Удалить эту ревизию из базы?')) return;
        const id = b.getAttribute('data-id');
        writeHistory(getHistoryList().filter(x => x.id !== id));
        renderHistoryScreen($('u-hist-search').value.trim());
      };
    });
  }

  // ---------------------------------------------------------------- boot
  window.__uzumState = state;
  injectUI();

  // SPA-контроллер: возвращает кнопку в шапку, если приложение перерисовало DOM
  setInterval(injectUI, 2000);
})();
