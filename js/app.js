const SHEET_ID = '1o5XdAT1wg8Jv5Boj6rdy_klGJHlB_am4E7eZXs9S15s';
let exchangeRate = 9.5;

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
    document.getElementById('exchange-rate').textContent = `1 JPY = ${exchangeRate.toFixed(2)} KRW`;
    document.getElementById('exchange-date').textContent = `기준: ${data.time_last_update_utc.split(' ').slice(0,4).join(' ')}`;
  } catch(e) {
    document.getElementById('exchange-rate').textContent = `1 JPY = ${exchangeRate} KRW (기본값)`;
  }
}

async function fetchSheetData(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
  try {
    const res = await fetch(url);
    const text = await res.text();
    return JSON.parse(text.substring(47, text.length - 2)).table;
  } catch(e) { return null; }
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

  table.rows.forEach((row, i) => {
    if (i < 2) return;
    const tr = document.createElement('tr');
    const c = Array.from({length: 7}, (_, j) => getCellValue(row, j));
    const col0 = c[0] || '', col1 = c[1] || '';

    if (col1.includes('소계')) tr.classList.add('subtotal');
    if (col1.includes('합계') || col1.includes('순이익')) tr.classList.add('grand-total');

    tr.innerHTML = `<td>${col0}</td><td>${col1}</td><td class="num">${fmt(c[2])}</td><td class="num">${fmt(c[3])}</td><td class="num">${fmt(c[4])}</td><td class="num">${fmt(c[5])}</td><td>${c[6]||''}</td>`;
    tbody.appendChild(tr);

    if (col1.includes('제작비') && col1.includes('VAT')) totalCost = c[5] || 0;
  });

  document.getElementById('kr-cost').textContent = fmtWon(totalCost);
}

function renderJapanTable(table) {
  if (!table) return;
  const tbody = document.getElementById('jp-tbody');
  tbody.innerHTML = '';
  let totalIncome = 0, totalExpense = 0;

  table.rows.forEach((row, i) => {
    if (i < 1) return;
    const tr = document.createElement('tr');
    const c = Array.from({length: 6}, (_, j) => getCellValue(row, j));
    const col0 = c[0] || '';

    if (col0.startsWith('[') || col0.startsWith('===')) {
      tr.innerHTML = `<td colspan="5" style="font-weight:600;background:var(--bg);font-size:13px;color:var(--text-secondary);letter-spacing:0.02em">${col0}</td>`;
    } else {
      if (col0.includes('소계') || col0.includes('합계')) tr.classList.add('subtotal');
      if (col0.includes('이익') || col0.includes('몫')) tr.classList.add('grand-total');
      tr.innerHTML = `<td>${col0}</td><td class="num">${fmt(c[1])}</td><td class="num">${fmt(c[2])}</td><td class="num">${fmtYen(c[3]||c[4])}</td><td class="num">${fmtWon(c[4]||c[5])}</td>`;
      if (col0 === '총 수입') totalIncome = c[4] || 0;
      if (col0 === '총 지출') totalExpense = c[4] || 0;
    }
    tbody.appendChild(tr);
  });

  document.getElementById('jp-income').textContent = fmtWon(totalIncome);
  document.getElementById('jp-expense').textContent = fmtWon(totalExpense);
  document.getElementById('jp-profit').textContent = fmtWon(totalIncome - totalExpense);
}

function renderMDTable(table) {
  if (!table) return;
  const tbody = document.getElementById('md-tbody');
  tbody.innerHTML = '';
  let krMdTotal = 0, currentSection = '';

  table.rows.forEach((row) => {
    const col0 = getCellValue(row, 0) || '';
    const tr = document.createElement('tr');

    if (col0.startsWith('[') || col0.startsWith('===')) {
      currentSection = col0;
      tr.classList.add('md-section-header');
      tr.innerHTML = `<td colspan="8" style="font-size:13px;padding:12px 10px;font-weight:600">${col0}</td>`;
      tbody.appendChild(tr); return;
    }
    if (col0 === '소계') {
      tr.classList.add('subtotal');
      const c = Array.from({length: 8}, (_, j) => getCellValue(row, j));
      if (currentSection.includes('한국')) krMdTotal = c[6] || 0;
      tr.innerHTML = `<td></td><td><strong>소계</strong></td><td></td><td></td><td class="num">${fmt(c[3])}</td><td class="num">${fmt(c[4])}</td><td class="num">${fmt(c[6])}</td><td class="num">${fmt(c[7])}</td>`;
      tbody.appendChild(tr); return;
    }
    if (!col0 || col0 === '상품명') return;
    const c = Array.from({length: 8}, (_, j) => getCellValue(row, j));
    tr.innerHTML = `<td></td><td>${col0}</td><td class="num">${fmt(c[1])}</td><td class="num">${fmt(c[2])}</td><td class="num">${fmt(c[3])}</td><td class="num">${fmt(c[4])}</td><td class="num">${fmt(c[5])}</td><td class="num">${fmt(c[6]||c[7])}</td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('kr-md').textContent = fmtWon(krMdTotal);
}

function renderDistribution(table) {
  if (!table) return;
  // Row map: 비율5-9, 한국게런티12-14, 일본게런티17-19, 한국계산22-24, 와테라스27-29
  let krSukRate=70, krHevRate=30, jpSukRate=50, jpImxRate=50;
  let krActorA=0, krActorB=0, jpActorA=0, jpActorB=0;
  let krSuk=0, krHev=0, jpSuk=0, jpImx=0;

  table.rows.forEach((row, i) => {
    const v = getCellValue(row, 2);
    const r = i + 1;
    if (r===5) krSukRate = v||70;
    if (r===6) krHevRate = v||30;
    if (r===8) jpSukRate = v||50;
    if (r===9) jpImxRate = v||50;
    if (r===12) krActorA = v||0;
    if (r===13) krActorB = v||0;
    if (r===17) jpActorA = v||0;
    if (r===18) jpActorB = v||0;
    if (r===23) krSuk = v||0;
    if (r===24) krHev = v||0;
    if (r===28) jpSuk = v||0;
    if (r===29) jpImx = v||0;
  });

  document.getElementById('dist-kr-suk').textContent = fmtWon(krSuk);
  document.getElementById('dist-kr-hev').textContent = fmtWon(krHev);
  document.getElementById('dist-kr-actor-a').textContent = fmtWon(krActorA);
  document.getElementById('dist-kr-actor-b').textContent = fmtWon(krActorB);
  document.getElementById('dist-kr-ratio').textContent = `${krSukRate} : ${krHevRate}`;

  document.getElementById('dist-jp-suk').textContent = fmtWon(jpSuk);
  document.getElementById('dist-jp-imx').textContent = fmtWon(jpImx);
  document.getElementById('dist-jp-actor-a').textContent = fmtWon(jpActorA);
  document.getElementById('dist-jp-actor-b').textContent = fmtWon(jpActorB);
  document.getElementById('dist-jp-ratio').textContent = `${jpSukRate} : ${jpImxRate}`;
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

  const [krData, jpData, mdData, distData] = await Promise.all([
    fetchSheetData('한국_견적'),
    fetchSheetData('일본_견적_와테라스'),
    fetchSheetData('MD_판매'),
    fetchSheetData('수익배분'),
  ]);

  renderKoreaTable(krData);
  renderJapanTable(jpData);
  renderMDTable(mdData);

  // 한국 순이익
  const krCost = parseFloat((document.getElementById('kr-cost').textContent).replace(/[₩,\-]/g,''))||0;
  const krMd = parseFloat((document.getElementById('kr-md').textContent).replace(/[₩,\-]/g,''))||0;
  document.getElementById('kr-profit').textContent = fmtWon(krMd - krCost);

  renderDistribution(distData);

  document.getElementById('last-update').textContent = `마지막 업데이트: ${new Date().toLocaleString('ko-KR')}`;
}

refreshData();
