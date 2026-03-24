const SHEET_ID = '1o5XdAT1wg8Jv5Boj6rdy_klGJHlB_am4E7eZXs9S15s';
const SHEETS = ['한국_견적', '일본_견적_오모테산도', '일본_견적_와테라스', 'MD_판매', '환율', '종합_대시보드'];

let exchangeRate = 9.5;
let sheetData = {};

function fmt(n) {
  if (n === null || n === undefined || n === '' || isNaN(n)) return '-';
  return Math.round(Number(n)).toLocaleString('ko-KR');
}

function fmtWon(n) {
  if (n === null || n === undefined || n === '' || isNaN(n)) return '-';
  return '₩' + Math.round(Number(n)).toLocaleString('ko-KR');
}

function fmtYen(n) {
  if (n === null || n === undefined || n === '' || isNaN(n)) return '-';
  return '¥' + Math.round(Number(n)).toLocaleString('ko-KR');
}

async function fetchExchangeRate() {
  try {
    const res = await fetch('https://open.er-api.com/v6/latest/JPY');
    const data = await res.json();
    exchangeRate = data.rates.KRW;
    const date = data.time_last_update_utc.split(' ').slice(0, 4).join(' ');
    document.getElementById('exchange-rate').textContent = `1 JPY = ${exchangeRate.toFixed(2)} KRW`;
    document.getElementById('exchange-date').textContent = `기준: ${date}`;
  } catch (e) {
    document.getElementById('exchange-rate').textContent = `1 JPY = ${exchangeRate} KRW (기본값)`;
    console.error('Exchange rate fetch failed:', e);
  }
}

async function fetchSheetData(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
  try {
    const res = await fetch(url);
    const text = await res.text();
    const json = JSON.parse(text.substring(47, text.length - 2));
    return json.table;
  } catch (e) {
    console.error(`Failed to fetch ${sheetName}:`, e);
    return null;
  }
}

function getCellValue(row, colIndex) {
  if (!row || !row.c || !row.c[colIndex]) return null;
  return row.c[colIndex].v;
}

function renderKoreaTable(table) {
  if (!table) return;
  const tbody = document.getElementById('kr-tbody');
  tbody.innerHTML = '';

  let totalCost = 0;
  let actorFee = 0;

  table.rows.forEach((row, i) => {
    if (i < 2) return;
    const tr = document.createElement('tr');
    const col0 = getCellValue(row, 0) || '';
    const col1 = getCellValue(row, 1) || '';
    const col2 = getCellValue(row, 2);
    const col3 = getCellValue(row, 3);
    const col4 = getCellValue(row, 4);
    const col5 = getCellValue(row, 5);
    const col6 = getCellValue(row, 6) || '';

    const isSubtotal = col1 === '소계' || col1.includes('소계');
    const isGrandTotal = col1.includes('합계') || col1.includes('순이익');

    if (isGrandTotal) tr.classList.add('grand-total');
    else if (isSubtotal) tr.classList.add('subtotal');

    tr.innerHTML = `
      <td>${col0}</td><td>${col1}</td>
      <td class="num">${fmt(col2)}</td><td class="num">${fmt(col3)}</td>
      <td class="num">${fmt(col4)}</td><td class="num">${fmt(col5)}</td>
      <td>${col6}</td>`;
    tbody.appendChild(tr);

    if (col1.includes('제작비합계') || col1.includes('제작비 합계')) totalCost = col5 || 0;
    if (col0.includes('배우') && col1 === '소계') actorFee = col5 || 0;
  });

  document.getElementById('kr-cost').textContent = fmtWon(totalCost);
  return actorFee;
}

function renderJapanTable(table, tbodyId, prefix) {
  if (!table) return;
  const tbody = document.getElementById(tbodyId);
  tbody.innerHTML = '';

  let totalIncome = 0, totalExpense = 0;

  table.rows.forEach((row, i) => {
    if (i < 3) return;
    const tr = document.createElement('tr');
    const col0 = getCellValue(row, 0) || '';
    const col1 = getCellValue(row, 1);
    const col2 = getCellValue(row, 2);
    const col3 = getCellValue(row, 3);
    const col4 = getCellValue(row, 4);

    const isHeader = col0.startsWith('===') || col0.startsWith('[');
    const isSubtotal = col0.includes('소계') || col0.includes('합계');
    const isSummary = col0.includes('이익') || col0.includes('몫');

    if (isHeader) {
      tr.innerHTML = `<td colspan="5" style="font-weight:600;background:var(--bg);font-size:13px;color:var(--text-secondary);text-transform:uppercase;letter-spacing:0.02em">${col0}</td>`;
    } else {
      if (isSummary) tr.classList.add('grand-total');
      else if (isSubtotal) tr.classList.add('subtotal');

      const yenVal = col1 || col3;
      const wonVal = col4;

      tr.innerHTML = `
        <td>${col0}</td>
        <td class="num">${fmt(col1)}</td>
        <td class="num">${fmt(col2)}</td>
        <td class="num">${fmtYen(col3)}</td>
        <td class="num">${fmtWon(col4)}</td>`;

      if (col0 === '총 수입') totalIncome = col4 || (yenVal * exchangeRate) || 0;
      if (col0 === '총 지출') totalExpense = col4 || (yenVal * exchangeRate) || 0;
    }

    tbody.appendChild(tr);
  });

  const profit = totalIncome - totalExpense;
  const share = profit * 0.5;

  document.getElementById(`${prefix}-income`).textContent = fmtWon(totalIncome);
  document.getElementById(`${prefix}-expense`).textContent = fmtWon(totalExpense);
  document.getElementById(`${prefix}-profit`).textContent = fmtWon(profit);
  document.getElementById(`${prefix}-share`).textContent = fmtWon(share);

  return { share, profit };
}

function renderMDTable(table) {
  if (!table) return;
  const tbody = document.getElementById('md-tbody');
  tbody.innerHTML = '';

  let krMdTotal = 0, jp1MdTotal = 0, jp2MdTotal = 0;
  let currentSection = '';

  table.rows.forEach((row, i) => {
    const col0 = getCellValue(row, 0) || '';
    const tr = document.createElement('tr');

    if (col0.startsWith('===') || col0.startsWith('[')) {
      currentSection = col0;
      tr.classList.add('md-section-header');
      tr.innerHTML = `<td colspan="8" style="font-size:14px;padding:12px 10px">${col0}</td>`;
      tbody.appendChild(tr);
      return;
    }

    if (col0 === '소계') {
      tr.classList.add('subtotal');
      const col3 = getCellValue(row, 3);
      const col4 = getCellValue(row, 4);
      const col6 = getCellValue(row, 6);
      const col7 = getCellValue(row, 7);

      if (currentSection.includes('한국')) krMdTotal = col6 || 0;

      tr.innerHTML = `<td></td><td><strong>소계</strong></td><td></td><td></td>
        <td class="num">${fmt(col3)}</td><td class="num">${fmt(col4)}</td>
        <td class="num">${fmt(col6)}</td><td class="num">${fmt(col7)}</td>`;
      tbody.appendChild(tr);
      return;
    }

    // Skip header rows and empty rows
    if (!col0 || col0 === '상품명') return;

    const col1 = getCellValue(row, 1);
    const col2 = getCellValue(row, 2);
    const col3 = getCellValue(row, 3);
    const col4 = getCellValue(row, 4);
    const col5 = getCellValue(row, 5);
    const col6 = getCellValue(row, 6);
    const col7 = getCellValue(row, 7);

    tr.innerHTML = `<td></td><td>${col0}</td>
      <td class="num">${fmt(col1)}</td><td class="num">${fmt(col2)}</td>
      <td class="num">${fmt(col3)}</td><td class="num">${fmt(col4)}</td>
      <td class="num">${fmt(col5)}</td><td class="num">${fmt(col6 || col7)}</td>`;
    tbody.appendChild(tr);
  });

  document.getElementById('kr-md').textContent = fmtWon(krMdTotal);
}

function updateTotals(jp1Share, jp2Share, jp1Profit, jp2Profit, actorFee) {
  const krCostEl = document.getElementById('kr-cost').textContent;
  const krMdEl = document.getElementById('kr-md').textContent;

  const krCost = parseFloat(krCostEl.replace(/[₩,\-]/g, '')) || 0;
  const krMd = parseFloat(krMdEl.replace(/[₩,\-]/g, '')) || 0;
  const krProfit = krMd - krCost;

  document.getElementById('kr-profit').textContent = fmtWon(krProfit);
  if (krProfit < 0) document.getElementById('kr-profit').classList.add('negative');

  // 석필름 = 한국 순이익 + 일본 50%
  const sukOmo = krProfit + (jp1Share || 0);
  const sukWat = krProfit + (jp2Share || 0);
  // IMX = 일본 50%
  const imxOmo = jp1Share || 0;
  const imxWat = jp2Share || 0;

  document.getElementById('total-omo-suk').textContent = fmtWon(sukOmo);
  document.getElementById('total-wat-suk').textContent = fmtWon(sukWat);
  document.getElementById('total-omo-imx').textContent = fmtWon(imxOmo);
  document.getElementById('total-wat-imx').textContent = fmtWon(imxWat);
  document.getElementById('total-omo-actor').textContent = fmtWon(actorFee);
  document.getElementById('total-wat-actor').textContent = fmtWon(actorFee);
}

// Collapsible sections
document.querySelectorAll('.detail-section h2').forEach(h2 => {
  h2.addEventListener('click', () => {
    const section = h2.parentElement;
    const container = section.querySelector('[id$="-container"]');
    if (container) {
      container.style.display = container.style.display === 'none' ? 'block' : 'none';
      section.classList.toggle('open');
    }
  });
});

async function refreshData() {
  document.getElementById('last-update').textContent = '데이터 로딩 중...';

  await fetchExchangeRate();

  const [krData, jp1Data, jp2Data, mdData] = await Promise.all([
    fetchSheetData('한국_견적'),
    fetchSheetData('일본_견적_오모테산도'),
    fetchSheetData('일본_견적_와테라스'),
    fetchSheetData('MD_판매'),
  ]);

  const actorFee = renderKoreaTable(krData);
  const jp1 = renderJapanTable(jp1Data, 'jp1-tbody', 'jp1');
  const jp2 = renderJapanTable(jp2Data, 'jp2-tbody', 'jp2');
  renderMDTable(mdData);
  updateTotals(jp1.share, jp2.share, jp1.profit, jp2.profit, actorFee);

  const now = new Date().toLocaleString('ko-KR');
  document.getElementById('last-update').textContent = `마지막 업데이트: ${now}`;
}

// Initial load
refreshData();
