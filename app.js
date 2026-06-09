const VAT_RATE = 0.12;
const SAMPLE_FILE = 'Заказы с товарами - АСОСИЙ.xls';
const RED = 'FF0000';
const BLACK = '000000';
const SUPPLIER_DEFAULTS = {
  name: '"NILPAK" MCHJ QK',
  address: 'Toshkent shahri, Yangihayot tumani, Janubiy Sanoat hududi, 28-uy',
  inn: '305124079.',
  vatReg: '326070033725.',
  account: '20208000900811474001.',
  bank: '01042,ТОШКЕНТ Ш., "КАПИТАЛБАНК" АТ БАНКИНИНГ СИРГАЛИ ФИЛ',
};
const BUYER_DEFAULTS = {
  name: '"ANGELESEY FOOD" MCHJ XK',
  address: 'Toshkent shaxri, Chilonzor tumani, Turob Tula ko\'chasi, 57-uy, "Besh Yog\'och" bozori hududida',
  inn: '202099756.',
  vatReg: '326060002860.',
  account: '20208000600578902001.',
  bank: '00450, ТОШКЕНТ Ш., "ТИФ МИЛЛИЙ БАНКИ" АЖ БОШ ОФИСИ',
};

const state = {
  fileName: '',
  rawRows: [],
  stores: [],
  warnings: [],
};

const els = {
  tabs: document.querySelectorAll('.tab'),
  panels: document.querySelectorAll('.panel'),
  fileInput: document.querySelector('#fileInput'),
  dropzone: document.querySelector('.dropzone'),
  statusText: document.querySelector('#statusText'),
  summaryList: document.querySelector('#summaryList'),
  resultsContent: document.querySelector('#resultsContent'),
  invoiceContent: document.querySelector('#invoiceContent'),
  loadSampleBtn: document.querySelector('#loadSampleBtn'),
  clearBtn: document.querySelector('#clearBtn'),
  exportCsvBtn: document.querySelector('#exportCsvBtn'),
  exportExcelBtn: document.querySelector('#exportExcelBtn'),
  exportInvoiceExcelBtn: document.querySelector('#exportInvoiceExcelBtn'),
  printBtn: document.querySelector('#printBtn'),
  buildInvoiceBtn: document.querySelector('#buildInvoiceBtn'),
  invoiceDate: document.querySelector('#invoiceDate'),
  supplierName: document.querySelector('#supplierName'),
  invoicePrefix: document.querySelector('#invoicePrefix'),
  invoiceMode: document.querySelector('#invoiceMode'),
  showVat: document.querySelector('#showVat'),
  groupByPrice: document.querySelector('#groupByPrice'),
};

els.invoiceDate.valueAsDate = new Date();

els.tabs.forEach((tab) => {
  tab.addEventListener('click', () => activateTab(tab.dataset.tab));
});

els.fileInput.addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) readFile(file);
});

['dragenter', 'dragover'].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.remove('dragover');
  });
});

els.dropzone.addEventListener('drop', (event) => {
  const [file] = event.dataTransfer.files;
  if (file) readFile(file);
});

els.loadSampleBtn.addEventListener('click', async () => {
  try {
    setStatus('Загружается пример из репозитория...');
    const response = await fetch(encodeURI(SAMPLE_FILE));
    if (!response.ok) throw new Error('Файл-пример не найден. Откройте сайт через локальный сервер.');
    const buffer = await response.arrayBuffer();
    processWorkbook(buffer, SAMPLE_FILE);
  } catch (error) {
    showError(error.message);
  }
});

els.clearBtn.addEventListener('click', () => {
  state.fileName = '';
  state.rawRows = [];
  state.stores = [];
  state.warnings = [];
  els.fileInput.value = '';
  setStatus('Файл ещё не загружен.');
  els.summaryList.innerHTML = '';
  els.resultsContent.className = 'empty-state';
  els.resultsContent.textContent = 'Сначала загрузите Excel-файл.';
  els.invoiceContent.className = 'invoice-preview empty-state';
  els.invoiceContent.textContent = 'Сначала загрузите Excel-файл.';
});

els.exportCsvBtn.addEventListener('click', exportCsv);
els.exportExcelBtn.addEventListener('click', exportInvoiceTemplateExcel);
els.exportInvoiceExcelBtn.addEventListener('click', exportInvoiceTemplateExcel);
els.printBtn.addEventListener('click', () => {
  activateTab('invoice');
  window.print();
});
els.buildInvoiceBtn.addEventListener('click', renderInvoice);
[els.supplierName, els.invoicePrefix, els.invoiceDate, els.invoiceMode, els.showVat].forEach((el) => {
  el.addEventListener('change', renderInvoice);
  el.addEventListener('input', renderInvoice);
});
els.groupByPrice.addEventListener('change', () => {
  if (state.rawRows.length) {
    const parsed = parseRows(state.rawRows);
    state.stores = parsed.stores;
    state.warnings = parsed.warnings;
    renderSummary();
    renderResults();
  }
  renderInvoice();
});

function activateTab(name) {
  els.tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.tab === name));
  els.panels.forEach((panel) => panel.classList.toggle('active', panel.id === name));
}

function readFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => processWorkbook(event.target.result, file.name);
  reader.onerror = () => showError('Ошибка чтения файла.');
  setStatus(`${file.name} читается...`);
  reader.readAsArrayBuffer(file);
}

function processWorkbook(buffer, fileName) {
  if (!window.XLSX) {
    showError('Библиотека XLSX не загружена. Проверьте подключение к интернету или подключите xlsx-js-style локально.');
    return;
  }

  try {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, raw: true });
    const allRows = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false });
      rows.forEach((row) => allRows.push(row));
    });

    const parsed = parseRows(allRows);
    state.fileName = fileName;
    state.rawRows = allRows;
    state.stores = parsed.stores;
    state.warnings = parsed.warnings;

    renderSummary();
    renderResults();
    renderInvoice();
    activateTab('results');
  } catch (error) {
    showError(`Не удалось проанализировать Excel-файл: ${error.message}`);
  }
}

function parseRows(rows) {
  const flatResult = parseFlatTableRows(rows);
  if (flatResult.stores.length) return flatResult;

  const warnings = [];
  const storeMap = new Map();
  let currentStore = null;
  let currentTotal = 0;
  let productColumns = null;

  rows.forEach((row, rowIndex) => {
    const normalized = row.map(normalizeText);

    const shopHeader = findShopHeader(normalized);
    if (shopHeader) {
      const dataRow = findNextDataRow(rows, rowIndex + 1, shopHeader.nameIndex);
      if (dataRow) {
        const storeName = cleanName(dataRow.row[shopHeader.nameIndex]);
        currentTotal = parseNumber(dataRow.row[shopHeader.totalIndex]);
        if (storeName) currentStore = getStore(storeMap, storeName, currentTotal);
      }
      productColumns = null;
      return;
    }

    const directShopIndex = normalized.findIndex((cell) => cell === 'магазин');
    if (directShopIndex !== -1 && !currentStore) {
      const name = cleanName(row[directShopIndex + 1]) || cleanName(row[directShopIndex]);
      if (name && name.toLowerCase() !== 'магазин') currentStore = getStore(storeMap, name, 0);
    }

    const detectedProductColumns = findProductColumns(normalized);
    if (detectedProductColumns) {
      productColumns = detectedProductColumns;
      return;
    }

    if (!productColumns || !currentStore) return;
    if (isSectionBreak(normalized)) return;

    const productName = cleanName(row[productColumns.product]);
    if (!productName || isIgnoredProductName(productName)) return;

    const quantity = parseNumber(row[productColumns.quantity]);
    const price = parseNumber(row[productColumns.price]);
    const net = parseNumber(row[productColumns.net]) || quantity * price;
    const vatCell = parseNumber(row[productColumns.vat]);
    const vat = vatCell && vatCell !== 12 ? vatCell : net * VAT_RATE;
    const gross = parseNumber(row[productColumns.gross]) || net + vat;

    if (!quantity && !net && !gross) return;

    addProduct(currentStore, {
      name: productName,
      quantity,
      price: price || (quantity ? net / quantity : 0),
      net,
      vat,
      gross,
      sourceRow: rowIndex + 1,
    }, els.groupByPrice.checked);
  });

  const stores = Array.from(storeMap.values()).map((store) => ({
    ...store,
    products: Array.from(store.products.values()).sort((a, b) => a.name.localeCompare(b.name, 'ru')),
  })).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  if (!stores.length) warnings.push('Строки магазинов и товаров не найдены. Проверьте названия заголовков.');
  stores.forEach((store) => {
    if (!store.products.length) warnings.push(`${store.name} — товары не найдены.`);
  });

  return { stores, warnings };
}


function parseFlatTableRows(rows) {
  const headerInfo = findFlatHeader(rows);
  if (!headerInfo) return { stores: [], warnings: [] };

  const { headerIndex, columns } = headerInfo;
  const productCol = findColumn(columns, ({ header }) => header === 'товар' || header.includes('товар'));
  const quantityCol = findColumn(columns, ({ header }) => header.includes('кол-во') || header.includes('количество'));
  const priceCol = findColumn(columns, ({ header }) => header.includes('цена за единицу') || header === 'цена');
  const netCol = findColumn(columns, ({ header, full }) => header === 'стоимость' && !full.includes('ндс'));
  const vatRateCol = findColumn(columns, ({ header, parent }) => header === 'ставка' && parent.includes('ндс'));
  const vatAmountCol = findColumn(columns, ({ header, parent }) => header === 'сумма' && parent.includes('ндс'));
  const grossCol = findColumn(columns, ({ full }) => full.includes('стоимость с учетом ндс') || full.includes('стоимость с учётом ндс'));
  const totalCol = findColumn(columns, ({ full }) => full.includes('общая сумма'));
  const orderCol = findColumn(columns, ({ header, parent }) => header === 'номер' && parent.includes('заказ'));
  const storeNameCol = findColumn(columns, ({ header, parent }) => header.includes('название') && parent.includes('магазин'));

  if (storeNameCol === -1 || productCol === -1 || quantityCol === -1) return { stores: [], warnings: [] };

  const storeMap = new Map();
  const orderStorePairs = new Set();

  rows.slice(headerIndex + 1).forEach((row, offset) => {
    const sourceRow = headerIndex + offset + 2;
    const storeName = cleanName(row[storeNameCol]);
    const productName = cleanName(row[productCol]);
    if (!storeName || !productName || isIgnoredProductName(productName)) return;

    const quantity = parseNumber(row[quantityCol]);
    const price = parseNumber(row[priceCol]);
    const net = parseNumber(row[netCol]) || quantity * price;
    const vatRate = parseNumber(row[vatRateCol]) || 12;
    const vat = parseNumber(row[vatAmountCol]) || net * (vatRate / 100 || VAT_RATE);
    const gross = parseNumber(row[grossCol]) || net + vat;
    if (!quantity && !net && !gross) return;

    const orderNumber = cleanName(row[orderCol]);
    const declaredTotal = parseNumber(row[totalCol]);
    const store = getStore(storeMap, storeName, 0, false);
    const pairKey = `${normalizeText(storeName)}|${orderNumber || sourceRow}`;
    if (!orderStorePairs.has(pairKey)) {
      orderStorePairs.add(pairKey);
      store.parts += 1;
      store.declaredTotal += declaredTotal || 0;
    }

    addProduct(store, {
      name: productName,
      quantity,
      price: price || (quantity ? net / quantity : 0),
      net,
      vat,
      gross,
      sourceRow,
    }, els.groupByPrice.checked);
  });

  const stores = Array.from(storeMap.values()).map((store) => ({
    ...store,
    products: Array.from(store.products.values()).sort((a, b) => a.name.localeCompare(b.name, 'ru')),
  })).filter((store) => store.products.length).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  return { stores, warnings: [] };
}

function findFlatHeader(rows) {
  for (let index = 0; index < Math.min(rows.length, 20); index += 1) {
    const headerRow = rows[index] || [];
    const normalized = headerRow.map(normalizeText);
    const hasProduct = normalized.some((cell) => cell === 'товар' || cell.includes('товар'));
    const hasQuantity = normalized.some((cell) => cell.includes('кол-во') || cell.includes('количество'));
    const hasPrice = normalized.some((cell) => cell.includes('цена за единицу'));
    if (!hasProduct || !hasQuantity || !hasPrice) continue;

    const parentRow = rows[Math.max(0, index - 1)] || [];
    const parents = fillMergedParents(parentRow, headerRow.length);
    const columns = headerRow.map((cell, columnIndex) => {
      const header = normalizeText(cell);
      const parent = normalizeText(parents[columnIndex]);
      return {
        index: columnIndex,
        header,
        parent,
        full: `${parent} ${header}`.trim(),
      };
    });
    return { headerIndex: index, columns };
  }
  return null;
}

function fillMergedParents(parentRow, length) {
  const parents = [];
  let current = '';
  for (let index = 0; index < length; index += 1) {
    const value = cleanName(parentRow[index]);
    if (value) current = value;
    parents[index] = current;
  }
  return parents;
}

function findColumn(columns, predicate) {
  const column = columns.find(predicate);
  return column ? column.index : -1;
}

function findShopHeader(normalized) {
  const nameIndex = normalized.findIndex((cell) => cell === 'название' || cell.includes('название'));
  const totalIndex = normalized.findIndex((cell) => cell.includes('общая сумма'));
  const hasShop = normalized.some((cell) => cell.includes('магазин'));
  if (nameIndex !== -1 && (totalIndex !== -1 || hasShop)) {
    return { nameIndex, totalIndex };
  }
  return null;
}

function findNextDataRow(rows, start, nameIndex) {
  for (let i = start; i < Math.min(rows.length, start + 6); i += 1) {
    const row = rows[i] || [];
    const name = cleanName(row[nameIndex]);
    const normalized = row.map(normalizeText);
    if (name && !normalized.includes('название') && !normalized.includes('товар')) return { row, index: i };
  }
  return null;
}

function findProductColumns(normalized) {
  const product = normalized.findIndex((cell) => cell === 'товар' || cell.includes('товар'));
  const quantity = normalized.findIndex((cell) => cell.includes('кол-во') || cell.includes('количество'));
  const price = normalized.findIndex((cell) => cell.includes('цена за единицу') || cell === 'цена');
  const gross = normalized.findIndex((cell) => cell.includes('стоимость с учетом ндс') || cell.includes('стоимость с учётом ндс'));
  const vat = normalized.findIndex((cell, index) => index !== gross && (cell === 'ндс' || cell.includes('ндс')));
  const net = normalized.findIndex((cell, index) => index !== gross && (cell === 'стоимость' || cell.includes('стоимость')));

  if (product !== -1 && quantity !== -1) {
    return { product, quantity, price, net, vat, gross };
  }
  return null;
}

function isSectionBreak(normalized) {
  return normalized.some((cell) => ['итого', 'всего', 'название', 'магазин'].includes(cell));
}

function isIgnoredProductName(name) {
  const value = normalizeText(name);
  return ['товар', 'итого', 'всего'].includes(value) || value.includes('общая сумма');
}

function getStore(storeMap, name, declaredTotal, incrementParts = true) {
  const key = normalizeText(name);
  if (!storeMap.has(key)) {
    storeMap.set(key, { name, declaredTotal: 0, products: new Map(), parts: 0 });
  }
  const store = storeMap.get(key);
  if (incrementParts) store.parts += 1;
  store.declaredTotal += declaredTotal || 0;
  return store;
}

function addProduct(store, product, groupByPrice) {
  const key = groupByPrice ? `${normalizeText(product.name)}|${normalizeNumber(product.price)}` : normalizeText(product.name);
  if (!store.products.has(key)) {
    store.products.set(key, { ...product, rows: [product.sourceRow] });
    return;
  }
  const existing = store.products.get(key);
  existing.quantity += product.quantity;
  existing.net += product.net;
  existing.vat += product.vat;
  existing.gross += product.gross;
  existing.price = existing.quantity ? existing.net / existing.quantity : product.price;
  existing.rows.push(product.sourceRow);
}

function normalizeText(value) {
  return String(value ?? '')
    .replace(/ё/g, 'е')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function cleanName(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const text = String(value ?? '').replace(/\s/g, '').replace('%', '').replace(',', '.').replace(/[^\d.-]/g, '');
  const number = Number.parseFloat(text);
  return Number.isFinite(number) ? number : 0;
}

function totals(products) {
  return products.reduce((acc, item) => {
    acc.quantity += item.quantity;
    acc.net += item.net;
    acc.vat += item.vat;
    acc.gross += item.gross;
    return acc;
  }, { quantity: 0, net: 0, vat: 0, gross: 0 });
}

function renderSummary() {
  const productCount = state.stores.reduce((sum, store) => sum + store.products.length, 0);
  const grand = totals(state.stores.flatMap((store) => store.products));
  setStatus(`<b>${state.fileName}</b> успешно проанализирован.`);
  els.summaryList.innerHTML = [
    `Магазины: <b>${state.stores.length}</b>`,
    `Объединённые строки товаров: <b>${productCount}</b>`,
    `Общая сумма с НДС: <b>${formatMoney(grand.gross)}</b>`,
    ...state.warnings.map((warning) => `<span class="error">${escapeHtml(warning)}</span>`),
  ].map((item) => `<li>${item}</li>`).join('');
}

function renderResults() {
  if (!state.stores.length) {
    els.resultsContent.className = 'empty-state';
    els.resultsContent.textContent = 'Нет результатов анализа.';
    return;
  }

  els.resultsContent.className = 'store-grid';
  els.resultsContent.innerHTML = state.stores.map((store) => storeTable(store)).join('');
}

function storeTable(store) {
  const total = totals(store.products);
  return `
    <article class="store-card">
      <div class="store-head">
        <div><strong>${escapeHtml(store.name)}</strong><br><span class="badge">${store.parts} блок(ов) объединено</span></div>
        <div class="kpis">
          <span class="kpi">Товар: ${store.products.length}</span>
          <span class="kpi">НДС: ${formatMoney(total.vat)}</span>
          <span class="kpi">Итого: ${formatMoney(total.gross)}</span>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>${productHeader(true)}</thead>
          <tbody>${store.products.map(productRow).join('')}</tbody>
          <tfoot>${totalRow(total, true)}</tfoot>
        </table>
      </div>
    </article>`;
}

function renderInvoice() {
  if (!state.stores.length) {
    els.invoiceContent.className = 'invoice-preview empty-state';
    els.invoiceContent.textContent = 'Сначала загрузите Excel-файл.';
    return;
  }

  const stores = rebuildStoresForCurrentGrouping();
  const mode = els.invoiceMode.value;
  els.invoiceContent.className = 'invoice-preview';
  els.invoiceContent.innerHTML = mode === 'single'
    ? invoiceDocument({ name: 'Общая счёт-фактура', products: stores.flatMap((store) => store.products) }, 1)
    : stores.map((store, index) => invoiceDocument(store, index + 1)).join('');
}

function rebuildStoresForCurrentGrouping() {
  const storeMap = new Map();
  state.stores.forEach((store) => {
    const target = getStore(storeMap, store.name, store.declaredTotal);
    store.products.forEach((product) => addProduct(target, product, els.groupByPrice.checked));
  });
  return Array.from(storeMap.values()).map((store) => ({ ...store, products: Array.from(store.products.values()) }));
}

function invoiceDocument(store, number) {
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const showVat = els.showVat.checked;
  return `
    <article class="invoice-document">
      <div class="invoice-title">
        <h2>Счёт-фактура № ${escapeHtml(els.invoicePrefix.value || 'NK')}-${String(number).padStart(3, '0')}</h2>
        <p>${formatDate(date)}</p>
      </div>
      <div class="invoice-meta">
        <div><b>Поставщик:</b> ${escapeHtml(els.supplierName.value || '-')}</div>
        <div><b>Покупатель / Магазин:</b> ${escapeHtml(store.name)}</div>
        <div><b>Исходный файл:</b> ${escapeHtml(state.fileName)}</div>
        <div><b>12% НДС:</b> ${formatMoney(total.vat)}</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>${productHeader(showVat)}</thead>
          <tbody>${store.products.map((item, index) => productRow(item, index + 1, showVat)).join('')}</tbody>
          <tfoot>${totalRow(total, showVat)}</tfoot>
        </table>
      </div>
      <div class="signatures">
        <div>Отпустил</div>
        <div>Получил</div>
      </div>
    </article>`;
}

function productHeader(showVat) {
  return `<tr>
    <th>№</th><th>Наименование товаров</th><th class="num">Кол-во</th><th class="num">Цена за единицу</th><th class="num">Стоимость</th>
    ${showVat ? '<th class="num">НДС 12%</th><th class="num">Стоимость с учётом НДС</th>' : ''}
  </tr>`;
}

function productRow(item, index = '', showVat = true) {
  return `<tr>
    <td>${index}</td><td>${escapeHtml(item.name)}</td><td class="num">${formatQuantity(item.quantity)}</td><td class="num">${formatMoney(item.price)}</td><td class="num">${formatMoney(item.net)}</td>
    ${showVat ? `<td class="num">${formatMoney(item.vat)}</td><td class="num">${formatMoney(item.gross)}</td>` : ''}
  </tr>`;
}

function totalRow(total, showVat) {
  return `<tr><td colspan="2">Итого</td><td class="num">${formatQuantity(total.quantity)}</td><td></td><td class="num">${formatMoney(total.net)}</td>${showVat ? `<td class="num">${formatMoney(total.vat)}</td><td class="num">${formatMoney(total.gross)}</td>` : ''}</tr>`;
}

function exportCsv() {
  if (!state.stores.length) return;
  const lines = buildFlatExportRows();
  const csv = lines.map((line) => line.map(csvCell).join(';')).join('\n');
  downloadBlob(csv, `nakladnoy-${Date.now()}.csv`, 'text/csv;charset=utf-8');
}

function exportExcel() {
  if (!state.stores.length || !window.XLSX) return;
  const workbook = XLSX.utils.book_new();
  const stores = rebuildStoresForCurrentGrouping();
  const mode = els.invoiceMode.value;

  if (mode === 'single') {
    appendStoreSheet(workbook, { name: 'Общая счёт-фактура', products: stores.flatMap((store) => store.products) }, 'Счёт-фактура');
  } else {
    stores.forEach((store, index) => appendStoreSheet(workbook, store, `${index + 1}-${store.name}`));
  }

  const summarySheet = XLSX.utils.aoa_to_sheet(buildFlatExportRows());
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Анализ');
  XLSX.writeFile(workbook, `nakladnoy-${Date.now()}.xlsx`);
}

function appendStoreSheet(workbook, store, sheetName) {
  const total = totals(store.products);
  const rows = [
    [`Счёт-фактура № ${els.invoicePrefix.value || 'NK'}`, '', '', '', '', '', ''],
    ['Дата', formatDate(els.invoiceDate.value || new Date().toISOString().slice(0, 10)), '', 'Магазин', store.name, '', ''],
    ['Поставщик', els.supplierName.value || '-', '', 'Исходный файл', state.fileName, '', ''],
    [],
    ['№', 'Наименование товаров', 'Кол-во', 'Цена за единицу', 'Стоимость', 'НДС 12%', 'Стоимость с учётом НДС'],
    ...store.products.map((item, index) => [index + 1, item.name, normalizeNumber(item.quantity), normalizeNumber(item.price), normalizeNumber(item.net), normalizeNumber(item.vat), normalizeNumber(item.gross)]),
    ['Итого', '', normalizeNumber(total.quantity), '', normalizeNumber(total.net), normalizeNumber(total.vat), normalizeNumber(total.gross)],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet['!cols'] = [{ wch: 8 }, { wch: 44 }, { wch: 12 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(sheetName));
}


function exportInvoiceTemplateExcel() {
  if (!state.stores.length || !window.XLSX) return;

  const workbook = XLSX.utils.book_new();
  const stores = rebuildStoresForCurrentGrouping();
  const mode = els.invoiceMode.value;
  const invoiceStores = mode === 'single'
    ? [{ name: 'Общая счёт-фактура', products: stores.flatMap((store) => store.products), parts: stores.length }]
    : stores;

  invoiceStores.forEach((store, storeIndex) => {
    const chunks = chunkProducts(store.products, 18);
    chunks.forEach((products, chunkIndex) => {
      const sheetStore = { ...store, products };
      const suffix = chunks.length > 1 ? `-${chunkIndex + 1}` : '';
      appendTemplateInvoiceSheet(workbook, sheetStore, storeIndex + 1, `${storeIndex + 1}-${store.name}${suffix}`);
    });
  });

  XLSX.writeFile(workbook, `schf-korzinka-${Date.now()}.xlsx`);
}

function appendTemplateInvoiceSheet(workbook, store, invoiceNumber, sheetName) {
  const worksheet = XLSX.utils.aoa_to_sheet(makeBlankRows(113, 22));
  worksheet['!cols'] = templateColumns();
  worksheet['!rows'] = templateRows();
  worksheet['!merges'] = templateMerges();
  worksheet['!margins'] = { left: 0.19685, right: 0.19685, top: 0.3937, bottom: 0.3937, header: 0, footer: 0 };
  worksheet['!pageSetup'] = { paperSize: 9, orientation: 'portrait', scale: 49, fitToWidth: 1, fitToHeight: 1 };

  fillInvoiceCopy(worksheet, 0, store, invoiceNumber);
  fillInvoiceCopy(worksheet, 57, store, invoiceNumber);
  applyTemplateInvoiceStyles(worksheet);

  XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(sheetName));
}

function fillInvoiceCopy(worksheet, offset, store, invoiceNumber) {
  const titleRow = 1 + offset;
  const contractRow = 2 + offset;
  const infoStart = 3 + offset;
  const headerRow = 10 + offset;
  const vatSubHeaderRow = 11 + offset;
  const firstProductRow = 12 + offset;
  const totalRowIndex = 30 + offset;
  const signRow = 34 + offset;
  const stampRow = 36 + offset;
  const total = totals(store.products);
  const date = els.invoiceDate.value || new Date().toISOString().slice(0, 10);
  const invoiceNo = `${els.invoicePrefix.value || 'NK'}-${String(invoiceNumber).padStart(3, '0')}`;

  const monthText = formatRussianInvoiceMonth(date);
  setCell(worksheet, `A${titleRow}`, store.name);
  setCell(worksheet, `H${titleRow}`, `СЧЕТ-ФАКТУРА № ${invoiceNo}`);
  setCell(worksheet, `P${titleRow}`, 'от');
  setCell(worksheet, `R${titleRow}`, `${monthText}г.`);
  setCell(worksheet, `A${contractRow}`, '      к Договору № 15/365 от «28» ноября 2022г.');

  setCell(worksheet, `A${infoStart}`, 'Поставщик:');
  setCell(worksheet, `C${infoStart}`, els.supplierName.value || SUPPLIER_DEFAULTS.name);
  setCell(worksheet, `M${infoStart}`, 'Покупатель:');
  setCell(worksheet, `O${infoStart}`, BUYER_DEFAULTS.name);
  setCell(worksheet, `A${infoStart + 1}`, 'Адрес:');
  setCell(worksheet, `C${infoStart + 1}`, SUPPLIER_DEFAULTS.address);
  setCell(worksheet, `M${infoStart + 1}`, 'Адрес:');
  setCell(worksheet, `O${infoStart + 1}`, BUYER_DEFAULTS.address);
  setCell(worksheet, `A${infoStart + 2}`, 'ИНН:');
  setCell(worksheet, `C${infoStart + 2}`, SUPPLIER_DEFAULTS.inn);
  setCell(worksheet, `M${infoStart + 2}`, 'ИНН:');
  setCell(worksheet, `O${infoStart + 2}`, BUYER_DEFAULTS.inn);
  setCell(worksheet, `A${infoStart + 3}`, 'РКП НДС:');
  setCell(worksheet, `C${infoStart + 3}`, SUPPLIER_DEFAULTS.vatReg);
  setCell(worksheet, `M${infoStart + 3}`, 'РКП НДС:');
  setCell(worksheet, `O${infoStart + 3}`, BUYER_DEFAULTS.vatReg);
  setCell(worksheet, `A${infoStart + 4}`, 'Банковский счет:');
  setCell(worksheet, `C${infoStart + 4}`, SUPPLIER_DEFAULTS.account);
  setCell(worksheet, `M${infoStart + 4}`, 'Банковский счет:');
  setCell(worksheet, `O${infoStart + 4}`, BUYER_DEFAULTS.account);
  setCell(worksheet, `A${infoStart + 5}`, 'МФО банка:');
  setCell(worksheet, `C${infoStart + 5}`, SUPPLIER_DEFAULTS.bank);
  setCell(worksheet, `M${infoStart + 5}`, 'МФО банка:');
  setCell(worksheet, `O${infoStart + 5}`, BUYER_DEFAULTS.bank);

  setCell(worksheet, `A${headerRow}`, '№');
  setCell(worksheet, `B${headerRow}`, 'Наименование товаров ');
  setCell(worksheet, `I${headerRow}`, 'Ед.');
  setCell(worksheet, `J${headerRow}`, 'Кол-во');
  setCell(worksheet, `M${headerRow}`, 'Цена');
  setCell(worksheet, `N${headerRow}`, 'Стоимость поставки');
  setCell(worksheet, `Q${headerRow}`, 'НДС');
  setCell(worksheet, `S${headerRow}`, 'Стоим. поставки с учетом НДС');
  setCell(worksheet, `Q${vatSubHeaderRow}`, 'Ставка');
  setCell(worksheet, `R${vatSubHeaderRow}`, 'Сумма');

  store.products.slice(0, 18).forEach((item, index) => {
    const row = firstProductRow + index;
    setCell(worksheet, `A${row}`, index + 1);
    setCell(worksheet, `B${row}`, item.name);
    setCell(worksheet, `I${row}`, 'пачка');
    setCell(worksheet, `J${row}`, normalizeNumber(item.quantity));
    setCell(worksheet, `M${row}`, normalizeNumber(item.price));
    setCell(worksheet, `N${row}`, normalizeNumber(item.net));
    setCell(worksheet, `Q${row}`, 12);
    setCell(worksheet, `R${row}`, normalizeNumber(item.vat));
    setCell(worksheet, `S${row}`, normalizeNumber(item.gross));
  });

  setCell(worksheet, `A${totalRowIndex}`, 'Итого:');
  setCell(worksheet, `N${totalRowIndex}`, normalizeNumber(total.net));
  setCell(worksheet, `R${totalRowIndex}`, normalizeNumber(total.vat));
  setCell(worksheet, `S${totalRowIndex}`, normalizeNumber(total.gross));
  setCell(worksheet, `A${signRow}`, 'Товар отпустил:');
  setCell(worksheet, `P${signRow}`, 'Получил:');
  setCell(worksheet, `B${stampRow}`, 'М.П.');
}

function makeBlankRows(rowCount, columnCount) {
  return Array.from({ length: rowCount }, () => Array.from({ length: columnCount }, () => ''));
}

function setCell(worksheet, address, value) {
  const existing = worksheet[address] || {};
  worksheet[address] = {
    ...existing,
    t: typeof value === 'number' ? 'n' : 's',
    v: value,
  };
  if (typeof value === 'number') worksheet[address].z = '#,##0.00';
}

function chunkProducts(products, size) {
  const chunks = [];
  for (let index = 0; index < products.length; index += size) {
    chunks.push(products.slice(index, index + size));
  }
  return chunks.length ? chunks : [[]];
}

function templateColumns() {
  return [
    { wch: 5.83 }, { wch: 17 }, { wch: 8.83 }, { wch: 4.33 }, { wch: 5 }, { wch: 5.33 }, { wch: 7.66 }, { wch: 33.16 },
    { wch: 11.83 }, { wch: 6.33 }, { wch: 6.33 }, { wch: 6.33 }, { wch: 23.33 }, { wch: 10.16 }, { wch: 10.16 }, { wch: 10.16 },
    { wch: 14 }, { wch: 11 }, { wch: 10.16 }, { wch: 10.16 }, { wch: 10.16 }, { wch: 2 },
  ];
}

function templateRows() {
  return Array.from({ length: 113 }, (_, index) => {
    const rowNumber = index + 1;
    if ([1, 2, 58, 59].includes(rowNumber)) return { hpt: 23.25 };
    if (rowNumber === 57) return { hpt: 24 };
    if ([4, 8, 61, 65].includes(rowNumber)) return { hpt: 30 };
    return { hpt: 18 };
  });
}

function templateMerges() {
  const merges = [];
  [0, 57].forEach((offset) => {
    merges.push(
      mergeRange(offset + 1, 1, offset + 1, 5),
      mergeRange(offset + 1, 8, offset + 1, 14),
      mergeRange(offset + 1, 16, offset + 1, 17),
      mergeRange(offset + 1, 18, offset + 1, 21),
      mergeRange(offset + 2, 1, offset + 2, 21),
      mergeRange(offset + 3, 3, offset + 3, 12),
      mergeRange(offset + 3, 15, offset + 3, 21),
      mergeRange(offset + 4, 3, offset + 4, 12),
      mergeRange(offset + 4, 15, offset + 4, 21),
      mergeRange(offset + 5, 3, offset + 5, 12),
      mergeRange(offset + 5, 15, offset + 5, 21),
      mergeRange(offset + 6, 3, offset + 6, 12),
      mergeRange(offset + 6, 15, offset + 6, 21),
      mergeRange(offset + 7, 3, offset + 7, 12),
      mergeRange(offset + 7, 15, offset + 7, 21),
      mergeRange(offset + 8, 3, offset + 8, 12),
      mergeRange(offset + 8, 15, offset + 8, 21),
      mergeRange(offset + 10, 1, offset + 11, 1),
      mergeRange(offset + 10, 2, offset + 11, 8),
      mergeRange(offset + 10, 9, offset + 11, 9),
      mergeRange(offset + 10, 10, offset + 11, 12),
      mergeRange(offset + 10, 13, offset + 11, 13),
      mergeRange(offset + 10, 14, offset + 11, 16),
      mergeRange(offset + 10, 17, offset + 10, 18),
      mergeRange(offset + 10, 19, offset + 11, 21),
      mergeRange(offset + 30, 1, offset + 30, 8),
      mergeRange(offset + 30, 14, offset + 30, 16),
      mergeRange(offset + 30, 19, offset + 30, 21),
      mergeRange(offset + 34, 1, offset + 34, 8),
      mergeRange(offset + 34, 16, offset + 34, 21),
      mergeRange(offset + 36, 2, offset + 36, 4),
    );

    for (let row = offset + 12; row <= offset + 29; row += 1) {
      merges.push(
        mergeRange(row, 2, row, 8),
        mergeRange(row, 10, row, 12),
        mergeRange(row, 14, row, 16),
        mergeRange(row, 19, row, 21),
      );
    }
  });
  return merges;
}


function applyTemplateInvoiceStyles(worksheet) {
  worksheet['!ref'] = 'A1:V113';
  [0, 57].forEach((offset, copyIndex) => {
    const color = copyIndex === 0 ? BLACK : RED;
    const infoStyle = templateCellStyle({ color: RED, bold: true, size: 9, horizontal: 'left', vertical: 'center', bottom: true });
    const titleStyle = templateCellStyle({ color, bold: true, size: 16, horizontal: 'center', vertical: 'center' });
    const contractStyle = templateCellStyle({ color, bold: true, size: 14, horizontal: 'center', vertical: 'center' });
    const headerStyle = templateCellStyle({ color, bold: true, size: 13, horizontal: 'center', vertical: 'center', wrapText: true, border: true });
    const bodyStyle = templateCellStyle({ color, bold: true, size: 12, horizontal: 'center', vertical: 'center', border: true });
    const nameStyle = templateCellStyle({ color, bold: true, size: 12, horizontal: 'left', vertical: 'center', border: true });
    const totalStyle = templateCellStyle({ color, bold: true, size: 12, horizontal: 'right', vertical: 'center', border: true });
    const signatureStyle = templateCellStyle({ color: RED, bold: true, size: 10, horizontal: 'left', vertical: 'center', top: true });

    styleRange(worksheet, offset + 1, 1, offset + 1, 21, titleStyle);
    styleRange(worksheet, offset + 2, 1, offset + 2, 21, contractStyle);
    styleRange(worksheet, offset + 3, 1, offset + 8, 21, infoStyle);
    styleRange(worksheet, offset + 10, 1, offset + 11, 21, headerStyle);
    styleRange(worksheet, offset + 12, 1, offset + 29, 21, bodyStyle);
    styleRange(worksheet, offset + 30, 1, offset + 30, 21, totalStyle);
    styleRange(worksheet, offset + 34, 1, offset + 34, 21, signatureStyle);
    styleRange(worksheet, offset + 36, 1, offset + 36, 4, templateCellStyle({ color: RED, bold: true, size: 10, horizontal: 'left', vertical: 'center' }));

    for (let row = offset + 12; row <= offset + 29; row += 1) {
      setStyle(worksheet, row, 2, nameStyle);
    }

    // Red header accents match the uploaded СЧФ template and screenshot.
    ['A', 'B', 'I', 'J', 'M', 'N', 'Q', 'R', 'S'].forEach((col) => {
      const address = `${col}${offset + 10}`;
      worksheet[address].s = { ...worksheet[address].s, font: { ...worksheet[address].s.font, color: { rgb: RED } } };
    });
    ['Q', 'R'].forEach((col) => {
      const address = `${col}${offset + 11}`;
      worksheet[address].s = { ...worksheet[address].s, font: { ...worksheet[address].s.font, color: { rgb: RED } } };
    });
  });
}

function templateCellStyle({ color = BLACK, bold = false, size = 11, horizontal = 'center', vertical = 'center', wrapText = false, border = false, bottom = false, top = false } = {}) {
  const thin = { style: 'thin', color: { rgb: BLACK } };
  const light = { style: 'thin', color: { rgb: 'D9D9D9' } };
  return {
    font: { name: 'Arial', sz: size, bold, color: { rgb: color } },
    alignment: { horizontal, vertical, wrapText },
    border: border
      ? { top: thin, right: thin, bottom: thin, left: thin }
      : { bottom: bottom ? thin : light, top: top ? thin : light, right: light, left: light },
  };
}

function styleRange(worksheet, startRow, startCol, endRow, endCol, style) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      setStyle(worksheet, row, col, style);
    }
  }
}

function setStyle(worksheet, row, col, style) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  worksheet[address] = worksheet[address] || { t: 's', v: '' };
  worksheet[address].s = style;
}

function mergeRange(startRow, startCol, endRow, endCol) {
  return { s: { r: startRow - 1, c: startCol - 1 }, e: { r: endRow - 1, c: endCol - 1 } };
}

function formatRussianDate(value) {
  const [year, month, day] = value.split('-');
  const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const monthIndex = Number(month) - 1;
  if (!day || monthIndex < 0 || monthIndex > 11 || !year) return value;
  return `${Number(day)} ${monthNames[monthIndex]} ${year}`;
}

function formatRussianInvoiceMonth(value) {
  const [year, month] = value.split('-');
  const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const monthIndex = Number(month) - 1;
  if (monthIndex < 0 || monthIndex > 11 || !year) return value;
  return `${monthNames[monthIndex]} ${year}`;
}

function buildFlatExportRows() {
  const lines = [['Магазин', 'Товар', 'Кол-во', 'Цена', 'Стоимость', 'НДС 12%', 'Стоимость с НДС']];
  state.stores.forEach((store) => {
    store.products.forEach((item) => lines.push([store.name, item.name, item.quantity, item.price, item.net, item.vat, item.gross]));
  });
  return lines;
}

function safeSheetName(name) {
  return String(name || 'Sheet').replace(/[:\\/?*\[\]]/g, ' ').slice(0, 31) || 'Sheet';
}

function csvCell(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}

function downloadBlob(content, name, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function setStatus(html) {
  els.statusText.innerHTML = html;
}

function showError(message) {
  setStatus(`<span class="error">${escapeHtml(message)}</span>`);
}

function roundMoney(value) {
  return normalizeNumber(value, 2);
}

function normalizeNumber(value, digits = 6) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return 0;
  return Math.round((number + Number.EPSILON) * (10 ** digits)) / (10 ** digits);
}

function formatMoney(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 6 }).format(normalizeNumber(value));
}

function formatQuantity(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 6 }).format(normalizeNumber(value));
}

function formatDate(value) {
  const [year, month, day] = value.split('-');
  return day && month && year ? `${day}.${month}.${year}` : value;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;',
  }[char]));
}
