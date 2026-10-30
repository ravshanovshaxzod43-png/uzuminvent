// ==UserScript==
// @name         Uzum Ревизор ПВЗ
// @namespace    https://dp.uzum.uz/
// @version      1.5.0
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

  const VERSION = '1.5.0';
  const API = 'https://api-wms.uzum.uz/de/delivery-point';
  const HISTORY_KEY = 'uzum_inventory_history';
  const jwtRegex = /eyJ[a-zA-Z0-9_-]{25,}\.eyJ[a-zA-Z0-9_-]{25,}\.[a-zA-Z0-9_-]{25,}/g;

  // ---------------------------------------------------------------- telegram
  // Токен и chat_id — скрипт виден всем, у кого есть доступ к нему.
  const TG = {
    enabled: true,
    token: 'PUT_BOT_TOKEN_HERE',
    chatId: 'PUT_CHAT_ID_HERE'
  };

  // ---------------------------------------------------------------- i18n
  const I18N = {
    ru: {
      appName: 'Ревизор ПВЗ',
      subtitle: 'Приёмка · Возвраты · Пересорт · История ревизий',
      tabNew: '📋 Новая ревизия',
      tabHistory: '🗄️ История ревизий',
      close: 'Закрыть',
      paramsTitle: 'Параметры ревизии',
      pvz: 'ПВЗ',
      dpKeyLabel: 'ID (dpKey)',
      ordersToIssue: n => `Заказов к выдаче: ${n}`,
      ordersToIssueEmpty: 'Заказов к выдаче: —',
      loadBtn: 'Выгрузить остатки и начать',
      statOrders: 'Заказов к выдаче',
      statPlan: 'Товаров по плану',
      statScanned: 'Отсканировано',
      barLabel: (f, t, p) => `Найдено ${f} из ${t} шт · ${p}%`,
      cellActive: 'Активная ячейка',
      cellPlaceholder: 'Отсканируйте ячейку',
      lastBarcode: 'Последний штрихкод',
      hintNeedCell: 'Сначала отсканируйте штрихкод ячейки',
      hintScanItems: 'Сканируйте товары на этой полке',
      scanPlaceholder: 'Сканер активен — направьте луч на штрихкод',
      undo: '↩️ Отменить скан',
      cellItems: 'Учтено в текущей ячейке',
      finish: 'Итоги ревизии →',
      emptyCell: 'Ячейка пуста. Сканируйте заказы на этой полке.',
      colOrderBarcode: 'Заказ / Штрихкод',
      colQty: 'Количество',
      badgeAcc: 'В ПРИЁМКЕ',
      badgeRet: 'ОТМЕНЁН / ВОЗВРАТ',
      badgeForeign: 'ЧУЖОЙ ПВЗ',
      orderLabel: n => `Заказ №${n}`,
      pcs: n => `${n} шт`,
      cellWord: 'ячейка',
      msgNeedCell: '⚠️ Сначала отсканируйте штрихкод ячейки (полки)!',
      msgAlreadyFull: pid => `ℹ️ Заказ №${pid} уже учтен полностью.`,
      msgAccepted: (pid, name) => `✓ Заказ №${pid} (${name}) принят.`,
      msgSearching: '🔍 Проверка статуса в WMS...',
      msgAcceptanceHint: '⚠️ В ПРИЁМКЕ (Заказ не завершен). Оставить в ячейке! Завершить приёмку в WMS.',
      msgAcceptanceDup: 'ℹ️ Эта позиция приёмки уже учтена.',
      msgReturnHint: '🔴 ОТМЕНЕН / ВОЗВРАТ! Срочно убрать с полки и отправить на склад.',
      msgReturnDup: 'ℹ️ Эта позиция возврата уже учтена.',
      msgForeignHint: '⛔ Чужой товар/посылка! Не числится на этом ПВЗ. Отложить в возврат!',
      msgUndone: code => `↩️ Скан "${code}" успешно отменен.`,
      statusAcceptance: 'Размещен / В приёмке',
      statusReturn: 'Завершен (К возврату)',
      statusAcceptanceShort: 'В приёмке',
      errNoToken: "Токен не найден. Обновите вкладку 'Заказы' (F5) и повторите.",
      errNoAccess: s => `Нет доступа к WMS API (${s}). Обновите страницу (F5) и повторите.`,
      errApi: s => `Ошибка WMS API (${s}).`,
      errNoDpKey: 'ID ПВЗ не указан. Введите ID точки (dpKey) вручную.',
      loadAuth: 'Авторизация в WMS...',
      loadList: k => `Выгрузка реестра заказов (dpKey: ${k})...`,
      loadCollected: n => `Собрано заказов к выдаче: ${n}...`,
      loadItems: (a, b) => `Загрузка состава товаров: ${a} из ${b}`,
      confirmFailed: n => `Не удалось загрузить состав ${n} заказов (для них план принят как 1 шт).\nПродолжить ревизию?`,
      errLoadCancelled: 'Загрузка отменена. Повторите попытку.',
      alreadySaved: 'Эта ревизия уже сохранена. Для новой ревизии загрузите остатки заново.',
      savedOk: pvz => `✓ Ревизия филиала ${pvz} успешно зафиксирована и сохранена в базу!\nАкт доступен во вкладке "История ревизий".`,
      saveFail: '⚠️ Не удалось сохранить ревизию в базу браузера (нет места). Акт .xlsx будет скачан.',
      confirmDelete: 'Удалить эту ревизию из базы?',
      histEmpty: 'В базе пока нет сохранённых ревизий. После проверки нажмите «Завершить целиком».',
      histNotFound: 'Ничего не найдено.',
      histTitle: 'Архив ревизий',
      histSearch: '🔍 Поиск по заказу (PID) или штрихкоду',
      chipOrders: n => `Заказов ${n}`,
      chipPlan: n => `План ${n}`,
      chipFact: n => `Факт ${n}`,
      chipShort: n => `Недостач ${n}`,
      chipMis: n => `Пересорт ${n}`,
      chipAcc: n => `Приёмка ${n}`,
      chipRet: n => `Возвраты ${n}`,
      btnXlsx: '📥 .xlsx',
      deleteTitle: 'Удалить из базы',
      reportTitle: pvz => `Итоги ревизии · ${pvz}`,
      reportMeta: (o, p, s) => `Заказов: <b>${o}</b> · По плану: <b>${p} шт</b> · Отсканировано: <b>${s} шт</b>`,
      btnSave: '💾 Завершить целиком',
      btnExcel: '📊 Скачать .xlsx',
      btnBack: '← Вернуться',
      tileShort: 'Недостачи',
      tileMis: 'Пересорт',
      tileAcc: 'В приёмке',
      tileRet: 'Возвраты',
      tileSur: 'Чужие',
      secShort: '🔴 Недостачи',
      secMis: '🟡 Пересорт ячеек',
      secAcc: '⚠️ В приёмке (не приняты полностью)',
      secRet: '🔴 Отменённые / Возвраты',
      secSur: '⛔ Чужие товары / не числятся',
      pos: n => `${n} поз.`,
      noShort: '✓ Недостач не обнаружено. Все клиентские заказы на месте!',
      noMis: '✓ Все заказы лежат строго на своих местах.',
      noAcc: 'Позиций в статусе приёмки на полках нет.',
      noRet: 'Отменённых заказов на полках не обнаружено.',
      noSur: 'Чужих товаров на складе не обнаружено.',
      thPid: 'Заказ (PID)',
      thMissing: 'Не хватает',
      thShelf: 'Полка',
      thClient: 'Клиент',
      thQty: 'Кол-во',
      thPlanCell: 'По плану',
      thActualCell: 'Лежит фактически',
      thBarcodeLoc: 'ШК / Грузоместо',
      thRec: 'Рекомендация',
      thBarcode: 'Штрихкод',
      thExtra: 'Лишнее кол-во',
      thFoundAt: 'Где найден',
      recAcc: 'Не отправлять на склад! Оставить в ячейке, провести приёмку.',
      recRet: 'Снять с полки! Оформить отгрузку на возвратный склад.',
      cellPrefix: 'ячейка ',
      unknown: 'Не определён',
      mgrName: 'Хожиев Сухроб',
      noName: '—',
      // telegram
      tgTitle: '📦 Ревизия ПВЗ завершена',
      tgPvz: (p, k) => `ПВЗ: ${p} (ID: ${k})`,
      tgDate: (d, t) => `Дата: ${d} ${t}`,
      tgAuditor: a => `Ревизор: ${a}`,
      tgOrders: n => `Заказов: ${n}`,
      tgPlan: n => `План: ${n} шт`,
      tgFact: n => `Факт: ${n} шт`,
      tgShort: (c, u) => `🔴 Недостачи: ${c} поз. (-${u} шт)`,
      tgMis: (c, u) => `🟡 Пересорт: ${c} поз. (${u} шт)`,
      tgAcc: (c, u) => `⚠️ В приёмке: ${c} поз. (${u} шт)`,
      tgRet: (c, u) => `🔴 Возвраты: ${c} поз. (${u} шт)`,
      tgSur: (c, u) => `⛔ Чужие: ${c} поз. (+${u} шт)`,
      // xlsx
      sheetName: 'Акт ревизии',
      xTitle: 'АКТ ИНВЕНТАРИЗАЦИИ СКЛАДА ПВЗ',
      xPvz: 'Пункт выдачи:',
      xDate: 'Дата и время:',
      xMgr: 'Территориальный управляющий:',
      xAuditor: 'Ревизор:',
      xStatus: 'Статус проверки:',
      xStatusDone: 'Завершена целиком (Сохранено в базу)',
      xSummary: '=== ИТОГОВАЯ СВОДКА РЕВИЗИИ ===',
      xHdrInd: 'Показатель',
      xHdrQty: 'Количество',
      xHdrStatus: 'Статус',
      xSumOrders: 'Всего клиентских заказов в плане WMS:',
      xSumPlan: 'Товаров по плану:',
      xSumScanned: 'Всего отсканировано товаров:',
      xSumShort: '🔴 Чистая недостача (не найдены на складе):',
      xSumMis: '🟡 Пересорт ячеек (лежат не на своих полках):',
      xSumAcc: '⚠️ Заказы в приёмке (не приняты полностью):',
      xSumRet: '🔴 Отмененные / Возвраты (к отгрузке на склад):',
      xSumSur: '⛔ Чужие товары (не числятся за ПВЗ):',
      xStPlanToIssue: 'План к выдаче',
      xStPlanPcs: 'План штук',
      xStFact: 'Факт на складе',
      xStShort: n => `Не хватает: -${n} шт`,
      xStMis: n => `${n} шт требуют перемещения`,
      xStAcc: n => `${n} шт (Оставить в ячейке, провести приемку)`,
      xStRet: n => `${n} шт (Срочно снять с полки на склад)`,
      xStSur: n => `+${n} шт в коробку возвратов`,
      xS1: '🔴 РАЗДЕЛ 1. НЕДОСТАЧИ (ЗАКАЗЫ, КОТОРЫХ НЕТ НА СКЛАДЕ)',
      xS2: '🟡 РАЗДЕЛ 2. ПЕРЕСОРТ ЯЧЕЕК (ТОВАР НА СКЛАДЕ, НО НЕ НА СВОЕЙ ПОЛКЕ)',
      xS3: '⚠️ РАЗДЕЛ 3. ТОВАРЫ В ПРИЁМКЕ (НЕ ПРИНЯТЫ ПОЛНОСТЬЮ — НЕ ОТПРАВЛЯТЬ НА СКЛАД!)',
      xS4: '🔴 РАЗДЕЛ 4. ОТМЕНЕННЫЕ ЗАКАЗЫ И ВОЗВРАТЫ (СРОЧНО УБРАТЬ С ПОЛКИ И ОТПРАВИТЬ НА СКЛАД)',
      xS5: '⛔ РАЗДЕЛ 5. ЧУЖИЕ ТОВАРЫ И НЕОПОЗНАННЫЕ ИЗЛИШКИ',
      xOkShort: '✓ Недостач не обнаружено. Все клиентские заказы на месте!',
      xOkMis: '✓ Все найденные заказы лежат строго на своих местах.',
      xOkAcc: '✓ Товаров в статусе незавершенной приёмки не обнаружено.',
      xOkRet: '✓ Забытых отмененных заказов или возвратов на полках не обнаружено.',
      xOkSur: '✓ Лишних или чужих посылок на складе не обнаружено.',
      xHdrS1: ['№', 'Номер заказа (PID)', 'ID в WMS', 'Штрихкод (SKU / Место)', 'Не хватает', 'Полка по плану', 'Где найден', 'Клиент', 'Телефон'],
      xHdrS2: ['№', 'Номер заказа (PID)', 'ID в WMS', 'Штрихкод (SKU / Место)', 'Кол-во', 'Полка по плану', 'Где лежит по факту', 'Клиент', 'Действие'],
      xHdrS3: ['№', 'Считанный ШК / Место', 'ID заказа', 'Кол-во', 'Полка', 'Статус в WMS', '', '', 'Рекомендуемое действие'],
      xHdrS5: ['№', 'Считанный ШК / Место', 'Количество', 'Где найден (Полка)', 'Статус', '', '', '', 'Рекомендуемое действие'],
      xNotFound: 'НЕ НАЙДЕН',
      xActS2: n => `Переложить в ячейку ${n}`,
      xActS3: 'Оставить в ячейке! Завершить приёмку в WMS',
      xActS4: 'Снять с полки! Оформить отгрузку на возвратный склад',
      xActS5: 'Отложить в коробку возвратов',
      xStatusRetShort: 'Завершен (Отгрузка)',
      xForeignStatus: 'Чужой ПВЗ / Не числится',
      xSaved: 'Сохранено в базу'
    },

    uz: {
      appName: 'PVZ inventarizatsiya',
      subtitle: 'Qabul · Qaytarish · Noto\'g\'ri joy · Inventarizatsiyalar tarixi',
      tabNew: '📋 Yangi inventarizatsiya',
      tabHistory: '🗄️ Inventarizatsiyalar tarixi',
      close: 'Yopish',
      paramsTitle: 'Inventarizatsiya parametrlari',
      pvz: 'PVZ',
      dpKeyLabel: 'ID (dpKey)',
      ordersToIssue: n => `Berishga buyurtmalar: ${n}`,
      ordersToIssueEmpty: 'Berishga buyurtmalar: —',
      loadBtn: 'Qoldiqlarni yuklab boshlash',
      statOrders: 'Berishga buyurtmalar',
      statPlan: 'Reja bo\'yicha mahsulotlar',
      statScanned: 'Skanerlangan',
      barLabel: (f, t, p) => `Topildi ${f} / ${t} dona · ${p}%`,
      cellActive: 'Faol katak',
      cellPlaceholder: 'Katakni skanerlang',
      lastBarcode: 'Oxirgi shtrixkod',
      hintNeedCell: 'Avval katak shtrixkodini skanerlang',
      hintScanItems: 'Ushbu javondagi mahsulotlarni skanerlang',
      scanPlaceholder: 'Skaner faol — shtrixkodga nurni yo\'naltiring',
      undo: '↩️ Skanni bekor qilish',
      cellItems: 'Joriy katakda hisobga olingan',
      finish: 'Inventarizatsiya natijalari →',
      emptyCell: 'Katak bo\'sh. Ushbu javondagi buyurtmalarni skanerlang.',
      colOrderBarcode: 'Buyurtma / Shtrixkod',
      colQty: 'Miqdor',
      badgeAcc: 'QABULDA',
      badgeRet: 'BEKOR QILINGAN / QAYTARISH',
      badgeForeign: 'BOSHQA PVZ',
      orderLabel: n => `Buyurtma №${n}`,
      pcs: n => `${n} dona`,
      cellWord: 'katak',
      msgNeedCell: '⚠️ Avval katak (javon) shtrixkodini skanerlang!',
      msgAlreadyFull: pid => `ℹ️ Buyurtma №${pid} to'liq hisobga olingan.`,
      msgAccepted: (pid, name) => `✓ Buyurtma №${pid} (${name}) qabul qilindi.`,
      msgSearching: '🔍 WMS da holat tekshirilmoqda...',
      msgAcceptanceHint: '⚠️ QABULDA (buyurtma yakunlanmagan). Katakda qoldiring! WMS da qabulni yakunlang.',
      msgAcceptanceDup: 'ℹ️ Bu qabul pozitsiyasi allaqachon hisobga olingan.',
      msgReturnHint: '🔴 BEKOR QILINGAN / QAYTARISH! Zudlik bilan javondan olib, omborga yuboring.',
      msgReturnDup: 'ℹ️ Bu qaytarish pozitsiyasi allaqachon hisobga olingan.',
      msgForeignHint: '⛔ Boshqa mahsulot/jo\'natma! Ushbu PVZ ga tegishli emas. Qaytarishga ajratib qo\'ying!',
      msgUndone: code => `↩️ "${code}" skani muvaffaqiyatli bekor qilindi.`,
      statusAcceptance: 'Joylashtirilgan / Qabulda',
      statusReturn: 'Yakunlangan (Qaytarishga)',
      statusAcceptanceShort: 'Qabulda',
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
      alreadySaved: 'Bu inventarizatsiya allaqachon saqlangan. Yangi inventarizatsiya uchun qoldiqlarni qaytadan yuklang.',
      savedOk: pvz => `✓ ${pvz} filiali inventarizatsiyasi muvaffaqiyatli qayd etildi va bazaga saqlandi!\nAkt "Inventarizatsiyalar tarixi" yorlig'ida mavjud.`,
      saveFail: '⚠️ Inventarizatsiyani brauzer bazasiga saqlab bo\'lmadi (joy yetarli emas). .xlsx akt yuklab olinadi.',
      confirmDelete: 'Bu inventarizatsiyani bazadan o\'chirasizmi?',
      histEmpty: 'Bazada hali saqlangan inventarizatsiyalar yo\'q. Tekshiruvdan keyin «Butunlay yakunlash» tugmasini bosing.',
      histNotFound: 'Hech narsa topilmadi.',
      histTitle: 'Inventarizatsiyalar arxivi',
      histSearch: '🔍 Buyurtma (PID) yoki shtrixkod bo\'yicha qidiruv',
      chipOrders: n => `Buyurtmalar ${n}`,
      chipPlan: n => `Reja ${n}`,
      chipFact: n => `Fakt ${n}`,
      chipShort: n => `Kamomad ${n}`,
      chipMis: n => `Noto'g'ri joy ${n}`,
      chipAcc: n => `Qabul ${n}`,
      chipRet: n => `Qaytarish ${n}`,
      btnXlsx: '📥 .xlsx',
      deleteTitle: 'Bazadan o\'chirish',
      reportTitle: pvz => `Inventarizatsiya natijalari · ${pvz}`,
      reportMeta: (o, p, s) => `Buyurtmalar: <b>${o}</b> · Reja bo'yicha: <b>${p} dona</b> · Skanerlangan: <b>${s} dona</b>`,
      btnSave: '💾 Butunlay yakunlash',
      btnExcel: '📊 .xlsx yuklash',
      btnBack: '← Orqaga',
      tileShort: 'Kamomad',
      tileMis: 'Noto\'g\'ri joy',
      tileAcc: 'Qabulda',
      tileRet: 'Qaytarish',
      tileSur: 'Begona',
      secShort: '🔴 Kamomad',
      secMis: '🟡 Noto\'g\'ri katakdagi buyurtmalar',
      secAcc: '⚠️ Qabulda (to\'liq qabul qilinmagan)',
      secRet: '🔴 Bekor qilingan / Qaytarishlar',
      secSur: '⛔ Begona mahsulotlar / ro\'yxatda yo\'q',
      pos: n => `${n} poz.`,
      noShort: '✓ Kamomad aniqlanmadi. Barcha mijoz buyurtmalari joyida!',
      noMis: '✓ Barcha buyurtmalar o\'z javonlarida.',
      noAcc: 'Javonlarda qabul holatidagi pozitsiyalar yo\'q.',
      noRet: 'Javonlarda bekor qilingan buyurtmalar topilmadi.',
      noSur: 'Omborda begona mahsulotlar topilmadi.',
      thPid: 'Buyurtma (PID)',
      thMissing: 'Yetishmaydi',
      thShelf: 'Javon',
      thClient: 'Mijoz',
      thQty: 'Miqdor',
      thPlanCell: 'Reja bo\'yicha',
      thActualCell: 'Haqiqatda turibdi',
      thBarcodeLoc: 'Shtrixkod / Joy',
      thRec: 'Tavsiya',
      thBarcode: 'Shtrixkod',
      thExtra: 'Ortiqcha miqdor',
      thFoundAt: 'Qayerda topildi',
      recAcc: 'Omborga yubormang! Katakda qoldiring, qabulni o\'tkazing.',
      recRet: 'Javondan oling! Qaytarish omboriga jo\'nating.',
      cellPrefix: 'katak ',
      unknown: 'Noma\'lum',
      mgrName: 'Xojiyev Suxrob',
      noName: '—',
      tgTitle: '📦 PVZ inventarizatsiyasi yakunlandi',
      tgPvz: (p, k) => `PVZ: ${p} (ID: ${k})`,
      tgDate: (d, t) => `Sana: ${d} ${t}`,
      tgAuditor: a => `Inventarizatsiya o'tkazuvchi: ${a}`,
      tgOrders: n => `Buyurtmalar: ${n}`,
      tgPlan: n => `Reja: ${n} dona`,
      tgFact: n => `Fakt: ${n} dona`,
      tgShort: (c, u) => `🔴 Kamomad: ${c} poz. (-${u} dona)`,
      tgMis: (c, u) => `🟡 Noto'g'ri joy: ${c} poz. (${u} dona)`,
      tgAcc: (c, u) => `⚠️ Qabulda: ${c} poz. (${u} dona)`,
      tgRet: (c, u) => `🔴 Qaytarishlar: ${c} poz. (${u} dona)`,
      tgSur: (c, u) => `⛔ Begona: ${c} poz. (+${u} dona)`,
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
      xSumShort: '🔴 Sof kamomad (omborda topilmadi):',
      xSumMis: '🟡 Noto\'g\'ri katak (boshqa javonda turibdi):',
      xSumAcc: '⚠️ Qabuldagi buyurtmalar (to\'liq qabul qilinmagan):',
      xSumRet: '🔴 Bekor qilingan / Qaytarishlar (omborga jo\'natish):',
      xSumSur: '⛔ Begona mahsulotlar (PVZ ga tegishli emas):',
      xStPlanToIssue: 'Berish rejasi',
      xStPlanPcs: 'Reja donalari',
      xStFact: 'Ombordagi fakt',
      xStShort: n => `Yetishmaydi: -${n} dona`,
      xStMis: n => `${n} dona joyini o'zgartirish kerak`,
      xStAcc: n => `${n} dona (katakda qoldiring, qabulni o'tkazing)`,
      xStRet: n => `${n} dona (zudlik bilan javondan olib omborga)`,
      xStSur: n => `+${n} dona qaytarish qutisiga`,
      xS1: '🔴 BO\'LIM 1. KAMOMAD (OMBORDA YO\'Q BUYURTMALAR)',
      xS2: '🟡 BO\'LIM 2. KATAKLAR ALMASHUVI (OMBORDA, LEKIN BOSHQA JAVONDA)',
      xS3: '⚠️ BO\'LIM 3. QABULDAGI MAHSULOTLAR (TO\'LIQ QABUL QILINMAGAN — OMBORGA YUBORMANG!)',
      xS4: '🔴 BO\'LIM 4. BEKOR QILINGAN BUYURTMALAR VA QAYTARISHLAR (ZUDLIK BILAN JAVONDAN OLING)',
      xS5: '⛔ BO\'LIM 5. BEGONA MAHSULOTLAR VA ANIQLANMAGAN ORTIQCHALAR',
      xOkShort: '✓ Kamomad aniqlanmadi. Barcha mijoz buyurtmalari joyida!',
      xOkMis: '✓ Topilgan buyurtmalar barchasi o\'z joyida.',
      xOkAcc: '✓ Tugallanmagan qabul holatidagi mahsulotlar topilmadi.',
      xOkRet: '✓ Javonlarda unutilgan bekor qilingan buyurtma yoki qaytarish topilmadi.',
      xOkSur: '✓ Omborda ortiqcha yoki begona jo\'natmalar topilmadi.',
      xHdrS1: ['№', 'Buyurtma raqami (PID)', 'WMS ID', 'Shtrixkod (SKU / Joy)', 'Yetishmaydi', 'Reja bo\'yicha javon', 'Qayerda topildi', 'Mijoz', 'Telefon'],
      xHdrS2: ['№', 'Buyurtma raqami (PID)', 'WMS ID', 'Shtrixkod (SKU / Joy)', 'Miqdor', 'Reja bo\'yicha javon', 'Haqiqatda qayerda', 'Mijoz', 'Harakat'],
      xHdrS3: ['№', 'O\'qilgan shtrixkod / Joy', 'Buyurtma ID', 'Miqdor', 'Javon', 'WMS holati', '', '', 'Tavsiya etilgan harakat'],
      xHdrS5: ['№', 'O\'qilgan shtrixkod / Joy', 'Miqdor', 'Qayerda topildi (javon)', 'Holat', '', '', '', 'Tavsiya etilgan harakat'],
      xNotFound: 'TOPILMADI',
      xActS2: n => `${n} katakka o'tkazish`,
      xActS3: 'Katakda qoldiring! WMS da qabulni yakunlang',
      xActS4: 'Javondan oling! Qaytarish omboriga jo\'nating',
      xActS5: 'Qaytarish qutisiga ajratib qo\'ying',
      xStatusRetShort: 'Yakunlangan (jo\'natma)',
      xForeignStatus: 'Boshqa PVZ / ro\'yxatda yo\'q',
      xSaved: 'Bazaga saqlangan'
    }
  };

  let LANG = 'ru';

  function detectLocale() {
    let v = '';
    try {
      v = sessionStorage.getItem('locale') || '';
      if (!v) {
        for (let i = 0; i < sessionStorage.length; i++) {
          const k = sessionStorage.key(i);
          if (k && /locale/i.test(k)) { v = sessionStorage.getItem(k) || ''; break; }
        }
      }
    } catch (e) {}
    return /uz/i.test(v) ? 'uz' : 'ru';
  }

  function refreshLocale() { LANG = detectLocale(); }

  function t(key, ...args) {
    const dict = I18N[LANG] || I18N.ru;
    let v = dict[key];
    if (v === undefined) v = I18N.ru[key];
    if (v === undefined) return key;
    return typeof v === 'function' ? v(...args) : v;
  }

  // ---------------------------------------------------------------- user (ревизор)
  function titleCase(s) {
    return String(s || '').toLowerCase().replace(/(^|[\s-])(\p{L})/gu, (m, p, c) => p + c.toUpperCase());
  }

  function getAuditorName() {
    const tryParse = (raw) => {
      try { return JSON.parse(raw); } catch (e) { return null; }
    };
    let raw = null;
    try {
      raw = sessionStorage.getItem('userInfo');
      if (!raw) {
        for (let i = 0; i < sessionStorage.length; i++) {
          const k = sessionStorage.key(i);
          if (k && /user/i.test(k)) { raw = sessionStorage.getItem(k); break; }
        }
      }
    } catch (e) {}
    const obj = raw ? tryParse(raw) : null;
    const info = obj && (obj.userInfo || obj);
    if (info && (info.firstName || info.lastName)) {
      return titleCase(`${info.firstName || ''} ${info.lastName || ''}`.trim());
    }
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
        if (res.status === 401 || res.status === 403) throw new Error(t('errNoAccess', res.status));
        if (!res.ok) throw new Error(t('errApi', res.status));
        return await res.json();
      } catch (e) {
        clearTimeout(timer);
        lastErr = e;
        if (e && e.isAccessError) throw e;
        if (/Нет доступа|WMS API ga kirish/.test(e.message)) { e.isAccessError = true; throw e; }
      }
    }
    throw lastErr;
  }

  function authHeaders() {
    const token = findTokenAnywhere();
    if (!token) throw new Error(t('errNoToken'));
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
    dpKey: null,
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
    savedInSession: false
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
            orderId: rawObj?.b2cOrderId || rawObj?.orderId || rawObj?.externalId || '',
            cargoPlace: rawObj?.cargoPlaceBarcode || rawObj?.barcode || '',
            statusText: t('statusAcceptance')
          };
          break;
        }

        if (str.includes('отгрузк') || str.includes('завершен') || str.includes('возврат') || str.includes('return') || str.includes('cancel')) {
          const rawObj = Array.isArray(data) ? data[0] : (data.shipment?.[0] || data.returns?.[0] || data.content?.[0] || data);
          result = {
            type: 'RETURN',
            orderId: rawObj?.b2cOrderId || rawObj?.orderId || '',
            cargoPlace: rawObj?.cargoPlaceBarcode || rawObj?.barcode || '',
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

    let page = 0;
    const size = 100;
    const orderHeaders = [];
    const seenHeaderIds = new Set();

    onProgress(t('loadList', dpKey));

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
      onProgress(t('loadCollected', orderHeaders.length));
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
      onProgress(t('loadItems', Math.min(i + chunkSize, orderIds.length), orderIds.length));

      let ordersWithItems = [];
      try {
        const r = await fetchJson(`${API}/b2c-orders/items?b2cOrderId=${slice.join(',')}`, authHeaders(), 30000, 2);
        ordersWithItems = Array.isArray(r) ? r : (r.content || []);
      } catch (e) {
        if (e && e.isAccessError) throw e;
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
    if (Object.prototype.hasOwnProperty.call(state.planCellUnits, code)) return true;
    if (state.codeToOrderId[code] || state.skuIndex[code]) return false;
    return code.length <= 5 && !code.startsWith('1000') && !code.includes('-');
  }

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
    const deepStatus = await deepLookupShipmentStatus(code);

    if (deepStatus && deepStatus.type === 'ACCEPTANCE') {
      if (!state.acceptanceScans[cell]) state.acceptanceScans[cell] = {};
      if (state.acceptanceScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', t('msgAcceptanceDup'));
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
      updateScannerUI(code, 'ACCEPTANCE', t('msgAcceptanceHint'));
      return;
    }

    if (deepStatus && deepStatus.type === 'RETURN') {
      if (!state.returnScans[cell]) state.returnScans[cell] = {};
      if (state.returnScans[cell][code]) {
        soundDuplicate();
        updateScannerUI(code, 'INFO', t('msgReturnDup'));
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

    updateScannerUI('—', 'INFO', t('msgUndone', last.code));
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

  // ---------------------------------------------------------------- telegram (фон, без UI)
  function buildTelegramText(entry) {
    const s = entry.stats || {};
    const lines = [
      t('tgTitle'),
      t('tgPvz', entry.pvzName, entry.dpKey),
      t('tgDate', entry.date, entry.time),
      t('tgAuditor', entry.auditor),
      '',
      t('tgOrders', s.totalOrders ?? 0),
      t('tgPlan', s.totalPlanUnits ?? 0),
      t('tgFact', s.totalScanned ?? 0),
      '',
      t('tgShort', s.shortageCount ?? 0, s.shortageUnits ?? 0),
      t('tgMis', s.mismatchCount ?? 0, s.mismatchUnits ?? 0),
      t('tgAcc', s.acceptanceCount ?? 0, s.acceptanceUnits ?? 0),
      t('tgRet', s.returnCount ?? 0, s.returnUnits ?? 0),
      t('tgSur', s.surplusCount ?? 0, s.surplusUnits ?? 0)
    ];
    return lines.join('\n');
  }

  async function sendTelegramSilently(text) {
    if (!TG.enabled || !TG.token || !TG.chatId || TG.token.startsWith('PUT_')) return false;
    const url = `https://api.telegram.org/bot${TG.token}/sendMessage`;
    const body = JSON.stringify({
      chat_id: TG.chatId,
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

  function exportCleanXlsx(reportData = null, auditMeta = null) {
    refreshLocale();
    const report = reportData || state.lastReport || buildReport();
    const now = new Date();
    const dateStr = auditMeta?.date || now.toLocaleDateString('ru-RU');
    const timeStr = auditMeta?.time || now.toLocaleTimeString('ru-RU');
    const pvzName = auditMeta?.pvzName || state.activePvzName;
    const dpKey = auditMeta?.dpKey || state.dpKey;
    const auditor = auditMeta?.auditor || getAuditorName();
    const totalOrders = auditMeta?.totalOrders ?? state.totalOrdersCount;
    const totalPlanUnits = auditMeta?.totalPlanUnits ?? state.totalPlanUnits;
    const totalScanned = auditMeta?.totalScanned ?? state.totalScannedUnits;

    const fullName = c => `${c?.firstName || ''} ${c?.lastName || ''}`.trim() || t('noName');

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
    rows.push([]);

    rows.push([t('xS1')]);
    if (report.pureShortages.length === 0) {
      rows.push([t('xOkShort')]);
    } else {
      rows.push(t('xHdrS1'));
      report.pureShortages.forEach((r, idx) => {
        rows.push([idx + 1, r.pid || r.orderId, r.orderId, r.sku, `-${r.missingCount} ${t('pcs', '').trim()}`, r.expectedCell || 'БЕЗ_ЯЧЕЙКИ', t('xNotFound'), fullName(r.customer), r.customer?.phone || t('noName')]);
      });
    }
    rows.push([]);

    rows.push([t('xS2')]);
    if (report.mismatchedCells.length === 0) {
      rows.push([t('xOkMis')]);
    } else {
      rows.push(t('xHdrS2'));
      report.mismatchedCells.forEach((r, idx) => {
        rows.push([idx + 1, r.pid || r.orderId, r.orderId, r.sku, t('pcs', r.count), r.expectedCell, r.actualCell, fullName(r.customer), t('xActS2', r.expectedCell)]);
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

    const colWidths = [6, 22, 14, 24, 14, 18, 18, 28, 22];
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
    refreshLocale();
    if (state.savedInSession) {
      alert(t('alreadySaved'));
      return;
    }

    const report = buildReport();
    const now = new Date();

    const auditEntry = {
      id: `INV-${Date.now()}`,
      date: now.toLocaleDateString('ru-RU'),
      time: now.toLocaleTimeString('ru-RU'),
      pvzName: state.activePvzName,
      dpKey: state.dpKey,
      auditor: getAuditorName(),
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

    // Telegram отправляется ТОЛЬКО после успешного сохранения
    if (saved) {
      state.savedInSession = true;
      sendTelegramSilently(buildTelegramText(auditEntry)).then(ok => {
        if (!ok) console.warn('[uzum-inv] telegram: сообщение не отправлено');
      });
    }

    soundSuccess();
    alert(saved ? t('savedOk', state.activePvzName) : t('saveFail'));
    exportCleanXlsx(report, {
      date: auditEntry.date,
      time: auditEntry.time,
      pvzName: auditEntry.pvzName,
      dpKey: auditEntry.dpKey,
      auditor: auditEntry.auditor,
      totalOrders: auditEntry.stats.totalOrders,
      totalPlanUnits: auditEntry.stats.totalPlanUnits,
      totalScanned: auditEntry.stats.totalScanned
    });
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

  // Статический текст в модалке обновляется при каждом открытии (по локали)
  function applyStaticTexts() {
    const set = (id, txt) => { const el = $(id); if (el) el.textContent = txt; };
    set('u-app-name', t('appName'));
    set('u-subtitle', t('subtitle'));
    set('u-close', '✕');
    $('u-close').title = t('close');
    set('u-tab-new', t('tabNew'));
    set('u-tab-history', t('tabHistory'));
    set('u-params-title', t('paramsTitle'));
    set('u-lbl-pvz', t('pvz'));
    set('u-lbl-dpkey', t('dpKeyLabel'));
    set('u-load-btn', t('loadBtn'));
    set('u-lbl-stat-orders', t('statOrders'));
    set('u-lbl-stat-plan', t('statPlan'));
    set('u-lbl-stat-scanned', t('statScanned'));
    set('u-lbl-cell', t('cellActive'));
    set('u-lbl-last', t('lastBarcode'));
    set('u-undo-btn', t('undo'));
    set('u-lbl-cell-items', t('cellItems'));
    set('u-finish-btn', t('finish'));
    set('u-hist-title', t('histTitle'));
    $('u-barcode-input').placeholder = t('scanPlaceholder');
    $('u-hist-search').placeholder = t('histSearch');
    if (!state.activeCell) set('u-active-cell', t('cellPlaceholder'));
    $('u-active-cell').classList.toggle('empty', !state.activeCell);
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
            <div class="uz-title" id="u-app-name"></div>
            <div class="uz-sub" id="u-subtitle"></div>
          </div>
          <span class="uz-ver">v${VERSION}</span>
          <button id="u-close" class="uz-x"></button>
        </div>

        <div class="uz-tabs">
          <button id="u-tab-new" class="uz-tab active"></button>
          <button id="u-tab-history" class="uz-tab"></button>
        </div>

        <div class="uz-body">
          <div id="u-start-screen" class="uz-start">
            <div class="uz-card">
              <div class="uz-lbl" id="u-params-title"></div>
              <div class="uz-row">
                <div>
                  <div class="uz-lbl" id="u-lbl-pvz" style="margin-bottom:6px"></div>
                  <input type="text" id="u-pvz-name-input" class="uz-in" style="width:140px">
                </div>
                <div>
                  <div class="uz-lbl" id="u-lbl-dpkey" style="margin-bottom:6px"></div>
                  <input type="text" id="u-pvz-dpkey-input" class="uz-in" style="width:120px">
                </div>
              </div>
              <div id="u-pvz-display-orders" class="uz-orders"></div>
              <div style="margin-top:22px"><button id="u-load-btn" class="uz-btn pri big"></button></div>
              <div id="u-progress" class="uz-prog"></div>
            </div>
          </div>

          <div id="u-scan-screen" style="display:none;">
            <div class="uz-stats">
              <div class="uz-stat"><div class="uz-lbl" id="u-lbl-stat-orders"></div><b id="u-sc-orders">0</b></div>
              <div class="uz-stat"><div class="uz-lbl" id="u-lbl-stat-plan"></div><b id="u-sc-plan">0</b></div>
              <div class="uz-stat"><div class="uz-lbl" id="u-lbl-stat-scanned"></div><b id="u-sc-total" style="color:var(--p)">0</b></div>
            </div>
            <div class="uz-bar"><div id="u-bar-fill"></div></div>
            <div class="uz-barlbl" id="u-bar-label"></div>

            <div class="uz-duo">
              <div class="uz-card">
                <div class="uz-lbl" id="u-lbl-cell"></div>
                <div id="u-active-cell" class="uz-cell empty"></div>
              </div>
              <div class="uz-card">
                <div class="uz-lbl" id="u-lbl-last"></div>
                <div id="u-last-barcode" class="uz-last">—</div>
                <div id="u-scan-hint" class="uz-banner"></div>
              </div>
            </div>

            <div class="uz-scanrow">
              <input type="text" id="u-barcode-input" autocomplete="off">
              <button id="u-undo-btn" class="uz-btn"></button>
            </div>

            <div class="uz-listhead">
              <span id="u-lbl-cell-items"></span>
              <button id="u-finish-btn" class="uz-btn ok"></button>
            </div>
            <div id="u-cell-items" class="uz-items"></div>
          </div>

          <div id="u-report-screen" style="display:none;"></div>

          <div id="u-history-screen" style="display:none;">
            <div class="uz-hhead">
              <h3 id="u-hist-title"></h3>
              <input type="text" id="u-hist-search">
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

  let btnFallbackTicks = 0;
  function ensureButton() {
    ensureStyle();
    let btn = $('uzum-inv-btn');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'uzum-inv-btn';
      btn.onclick = openModal;
    }
    btn.innerHTML = '📦 ' + (I18N.ru.appName);

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
    refreshLocale();
    ensureModal();
    applyStaticTexts();
    $('uzum-inv-modal').style.display = 'flex';

    if (state.screen === 'start' || !state.workScreen || state.workScreen === 'start') {
      const detectedName = detectPvzNameFromHeader();
      const liveDpKey = detectLiveDpKey();
      const ordersCount = detectOrdersCount();

      if (detectedName !== 'ПВЗ') state.activePvzName = detectedName;
      state.dpKey = liveDpKey || state.dpKey || '';

      $('u-pvz-name-input').value = state.activePvzName;
      $('u-pvz-dpkey-input').value = state.dpKey || '';
      $('u-pvz-display-orders').textContent = ordersCount === '—' ? t('ordersToIssueEmpty') : t('ordersToIssue', ordersCount);
      showScreen('start');
    } else {
      showScreen(state.workScreen);
    }
    renderReportScreenIfOpen();
    updateScannerUI();
  }

  function renderReportScreenIfOpen() {
    if (state.screen === 'report') renderReportScreen();
  }

  function updateScannerUI(lastBarcode = '—', statusType = 'NORMAL', hintText = '') {
    if (!$('u-active-cell')) return;

    const cellEl = $('u-active-cell');
    cellEl.textContent = state.activeCell || t('cellPlaceholder');
    cellEl.classList.toggle('empty', !state.activeCell);
    $('u-last-barcode').textContent = lastBarcode;

    $('u-sc-orders').textContent = state.totalOrdersCount;
    $('u-sc-plan').textContent = t('pcs', state.totalPlanUnits);
    $('u-sc-total').textContent = t('pcs', state.totalScannedUnits);

    const found = Object.values(state.factGlobal).reduce((a, b) => a + b, 0);
    const pct = state.totalPlanUnits ? Math.min(100, Math.round(found / state.totalPlanUnits * 100)) : 0;
    $('u-bar-fill').style.width = pct + '%';
    $('u-bar-label').textContent = t('barLabel', found, state.totalPlanUnits, pct);

    const classes = { OK: 'ok', INFO: 'info', ACCEPTANCE: 'acc', RETURN: 'ret', FOREIGN: 'bad', SEARCH: 'search' };
    const hintEl = $('u-scan-hint');
    hintEl.className = 'uz-banner ' + (classes[statusType] || '');
    hintEl.textContent = hintText || (state.activeCell ? t('hintScanItems') : t('hintNeedCell'));

    const container = $('u-cell-items');
    const orderItems = (state.activeCell && state.factCellUnits[state.activeCell]) || {};
    const accItems = (state.activeCell && state.acceptanceScans[state.activeCell]) || {};
    const retItems = (state.activeCell && state.returnScans[state.activeCell]) || {};
    const forItems = (state.activeCell && state.foreignScans[state.activeCell]) || {};

    const hasAny = Object.keys(orderItems).length + Object.keys(accItems).length + Object.keys(retItems).length + Object.keys(forItems).length > 0;

    if (!hasAny) {
      container.innerHTML = `<div class="uz-empty">${escHtml(t('emptyCell'))}</div>`;
      return;
    }

    let rows = '';
    Object.entries(orderItems).forEach(([oId, qty]) => {
      const order = state.ordersMap[oId];
      rows += `<tr><td><b>${escHtml(t('orderLabel', order?.pid || oId))}</b> <span style="color:var(--mut);font-size:11.5px">${escHtml(order?.customer?.firstName || '')}</span></td><td class="r" style="color:var(--ok)">${escHtml(t('pcs', qty))}</td></tr>`;
    });
    Object.entries(accItems).forEach(([code, data]) => {
      rows += `<tr class="acc"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag acc">${escHtml(t('badgeAcc'))}</span></td><td class="r" style="color:var(--acc)">+${escHtml(t('pcs', data.count))}</td></tr>`;
    });
    Object.entries(retItems).forEach(([code, data]) => {
      rows += `<tr class="ret"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag ret">${escHtml(t('badgeRet'))}</span></td><td class="r" style="color:var(--ret)">+${escHtml(t('pcs', data.count))}</td></tr>`;
    });
    Object.entries(forItems).forEach(([code, qty]) => {
      rows += `<tr class="bad"><td><span class="mono">${escHtml(code)}</span><span class="uz-tag bad">${escHtml(t('badgeForeign'))}</span></td><td class="r" style="color:var(--bad)">+${escHtml(t('pcs', qty))}</td></tr>`;
    });

    container.innerHTML = `<table class="uz-tbl"><thead><tr><th>${escHtml(t('colOrderBarcode'))}</th><th style="text-align:right">${escHtml(t('colQty'))}</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  const uzTable = (heads, rows) =>
    `<table class="uz-tbl"><thead><tr>${heads.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>`;
  const uzNone = (txt, ok) => `<div class="uz-none ${ok ? 'ok' : ''}">${txt}</div>`;
  const uzSection = (cls, title, count, units, body) =>
    `<div class="uz-sec ${cls}"><div class="uz-sec-h"><span>${title}</span><span class="uz-pill">${t('pos', count)} · ${units}</span></div>${body}</div>`;

  function renderReportScreen() {
    refreshLocale();
    const report = buildReport();
    const container = $('u-report-screen');

    const tiles = [
      { cls: 'bad', tt: t('tileShort'), n: report.pureShortages.length, u: `-${t('pcs', report.totalShortageUnits)}` },
      { cls: 'warn', tt: t('tileMis'), n: report.mismatchedCells.length, u: t('pcs', report.totalMismatchUnits) },
      { cls: 'acc', tt: t('tileAcc'), n: report.inAcceptance.length, u: t('pcs', report.totalAcceptanceUnits) },
      { cls: 'ret', tt: t('tileRet'), n: report.canceledReturns.length, u: t('pcs', report.totalReturnUnits) },
      { cls: 'mut', tt: t('tileSur'), n: report.pureSurplus.length, u: `+${t('pcs', report.totalSurplusUnits)}` }
    ];

    let html = `
      <div class="uz-rhead">
        <div>
          <h3>${escHtml(t('reportTitle', state.activePvzName))}</h3>
          <div class="uz-rmeta">${t('reportMeta', state.totalOrdersCount, state.totalPlanUnits, state.totalScannedUnits)}</div>
        </div>
        <div class="uz-acts">
          <button id="u-save-audit-btn" class="uz-btn ok">${t('btnSave')}</button>
          <button id="u-excel-btn" class="uz-btn pri">${t('btnExcel')}</button>
          <button id="u-resume-btn" class="uz-btn">${t('btnBack')}</button>
        </div>
      </div>
      <div class="uz-tiles">${tiles.map(x => `<div class="uz-tile ${x.cls}"><div class="n">${x.n}</div><div class="t">${x.tt}</div><div class="u">${x.u}</div></div>`).join('')}</div>`;

    html += uzSection('bad', t('secShort'), report.pureShortages.length, `-${t('pcs', report.totalShortageUnits)}`,
      report.pureShortages.length === 0
        ? uzNone(t('noShort'), true)
        : uzTable([t('thPid'), t('thMissing'), t('thShelf'), t('thClient')], report.pureShortages.map(r => `<tr>
            <td><b>№${escHtml(r.pid || r.orderId)}</b></td>
            <td style="color:var(--bad);font-weight:800">-${escHtml(t('pcs', r.missingCount))}</td>
            <td>${escHtml(r.expectedCell)}</td>
            <td>${escHtml(r.customer?.firstName || '')} ${r.customer?.phone ? '(' + escHtml(r.customer.phone) + ')' : ''}</td></tr>`).join('')));

    html += uzSection('warn', t('secMis'), report.mismatchedCells.length, t('pcs', report.totalMismatchUnits),
      report.mismatchedCells.length === 0
        ? uzNone(t('noMis'), true)
        : uzTable([t('thPid'), t('thQty'), t('thPlanCell'), t('thActualCell')], report.mismatchedCells.map(r => `<tr>
            <td><b>${escHtml(r.pid || r.orderId)}</b></td>
            <td style="font-weight:800">${escHtml(t('pcs', r.count))}</td>
            <td style="color:var(--bad);font-weight:800">${escHtml(r.expectedCell)}</td>
            <td style="color:var(--ok);font-weight:800">${escHtml(r.actualCell)}</td></tr>`).join('')));

    html += uzSection('acc', t('secAcc'), report.inAcceptance.length, t('pcs', report.totalAcceptanceUnits),
      report.inAcceptance.length === 0
        ? uzNone(t('noAcc'), false)
        : uzTable([t('thBarcodeLoc'), t('thQty'), t('thShelf'), t('thRec')], report.inAcceptance.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--acc);font-weight:800">+${escHtml(t('pcs', r.count))}</td>
            <td>${escHtml(t('cellWord'))} ${escHtml(r.foundCell)}</td>
            <td style="font-weight:700">${escHtml(t('recAcc'))}</td></tr>`).join('')));

    html += uzSection('ret', t('secRet'), report.canceledReturns.length, t('pcs', report.totalReturnUnits),
      report.canceledReturns.length === 0
        ? uzNone(t('noRet'), false)
        : uzTable([t('thBarcodeLoc'), t('thQty'), t('thShelf'), t('thRec')], report.canceledReturns.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--ret);font-weight:800">+${escHtml(t('pcs', r.count))}</td>
            <td>${escHtml(t('cellWord'))} ${escHtml(r.foundCell)}</td>
            <td style="font-weight:700">${escHtml(t('recRet'))}</td></tr>`).join('')));

    html += uzSection('mut', t('secSur'), report.pureSurplus.length, `+${t('pcs', report.totalSurplusUnits)}`,
      report.pureSurplus.length === 0
        ? uzNone(t('noSur'), false)
        : uzTable([t('thBarcode'), t('thExtra'), t('thFoundAt')], report.pureSurplus.map(r => `<tr>
            <td class="mono">${escHtml(r.code)}</td>
            <td style="color:var(--bad);font-weight:800">+${escHtml(t('pcs', r.qty))}</td>
            <td>${escHtml(t('cellWord'))} ${escHtml(r.foundCell)}</td></tr>`).join('')));

    container.innerHTML = html;
    showScreen('report');

    $('u-save-audit-btn').onclick = saveAuditPermanently;
    $('u-excel-btn').onclick = () => exportCleanXlsx();
    $('u-resume-btn').onclick = () => showScreen('scan');
  }

  function renderHistoryScreen(searchFilter = '') {
    refreshLocale();
    const container = $('u-hist-container');
    const list = getHistoryList();

    if (list.length === 0) {
      container.innerHTML = `<div class="uz-empty">${escHtml(t('histEmpty'))}</div>`;
      return;
    }

    let filtered = list;
    if (searchFilter) {
      const q = searchFilter.toLowerCase();
      filtered = list.filter(item => JSON.stringify(item).toLowerCase().includes(q));
    }

    if (filtered.length === 0) {
      container.innerHTML = `<div class="uz-empty">${escHtml(t('histNotFound'))}</div>`;
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
              <span class="uz-chip">${escHtml(t('chipOrders', s.totalOrders ?? 0))}</span>
              <span class="uz-chip">${escHtml(t('chipPlan', s.totalPlanUnits ?? 0))}</span>
              <span class="uz-chip">${escHtml(t('chipFact', s.totalScanned ?? 0))}</span>
              <span class="uz-chip bad">${escHtml(t('chipShort', s.shortageCount ?? 0))}</span>
              <span class="uz-chip warn">${escHtml(t('chipMis', s.mismatchCount ?? 0))}</span>
              <span class="uz-chip acc">${escHtml(t('chipAcc', s.acceptanceCount ?? 0))}</span>
              <span class="uz-chip ret">${escHtml(t('chipRet', s.returnCount ?? 0))}</span>
            </div>
          </div>
          <div class="uz-acts">
            <button class="uz-btn pri u-download-past-btn" data-id="${escHtml(item.id)}">${t('btnXlsx')}</button>
            <button class="uz-btn del u-delete-past-btn" data-id="${escHtml(item.id)}" title="${escHtml(t('deleteTitle'))}">🗑</button>
          </div>
        </div>`;
    });

    container.innerHTML = html;

    container.querySelectorAll('.u-download-past-btn').forEach(b => {
      b.onclick = () => {
        const found = getHistoryList().find(x => x.id === b.getAttribute('data-id'));
        if (found) {
          exportCleanXlsx(found.report, {
            date: found.date,
            time: found.time,
            pvzName: found.pvzName,
            dpKey: found.dpKey,
            auditor: found.auditor,
            totalOrders: found.stats.totalOrders,
            totalPlanUnits: found.stats.totalPlanUnits,
            totalScanned: found.stats.totalScanned
          });
        }
      };
    });

    container.querySelectorAll('.u-delete-past-btn').forEach(b => {
      b.onclick = () => {
        if (!confirm(t('confirmDelete'))) return;
        const id = b.getAttribute('data-id');
        writeHistory(getHistoryList().filter(x => x.id !== id));
        renderHistoryScreen($('u-hist-search').value.trim());
      };
    });
  }

  // ---------------------------------------------------------------- boot
  window.__uzumState = state;
  injectUI();

  setInterval(injectUI, 2000);
})();
