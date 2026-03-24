const SHEET_ID = '1o5XdAT1wg8Jv5Boj6rdy_klGJHlB_am4E7eZXs9S15s';

function fmt(n) {
  if (n===null||n===undefined||n===''||isNaN(n)) return '-';
  return Math.round(Number(n)).toLocaleString('ko-KR');
}
function fmtWon(n) {
  if (n===null||n===undefined||n===''||isNaN(n)) return '-';
  const v = Math.round(Number(n));
  return (v<0?'-':'') + '₩' + Math.abs(v).toLocaleString('ko-KR');
}

async function fetchExchangeRate() {
  try {
    const res = await fetch('https://open.er-api.com/v6/latest/JPY');
    const d = await res.json();
    document.getElementById('exchange-rate').textContent = `1 JPY = ${d.rates.KRW.toFixed(2)} KRW`;
    document.getElementById('exchange-date').textContent = `기준: ${d.time_last_update_utc.split(' ').slice(0,4).join(' ')}`;
  } catch(e) {
    document.getElementById('exchange-rate').textContent = '환율 로드 실패';
  }
}

async function fetchSheet(name) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(name)}`;
  try {
    const res = await fetch(url);
    const txt = await res.text();
    const table = JSON.parse(txt.substring(47, txt.length-2)).table;
    return table.rows.map(r => r.c ? r.c.map(c => c ? c.v : null) : []);
  } catch(e) { return []; }
}

// Generic table renderer
function renderTable(rows, container, opts={}) {
  const div = document.createElement('div');
  const table = document.createElement('table');
  table.className = 'data-table';
  rows.forEach(cells => {
    if (!cells || cells.every(c => !c)) return;
    const tr = document.createElement('tr');
    const first = (cells[0]||'').toString();
    const second = (cells[1]||'').toString();

    // Section header
    if (first.startsWith('[')) {
      tr.innerHTML = `<td colspan="${cells.length}" class="section-header">${first}</td>`;
      table.appendChild(tr); return;
    }
    // Table header (dark)
    if (['구분','항목','좌석 종류','상품명'].includes(first) || ['소구분','항목','가격(₩)','단가(₩)','판매원가(¥)','판매원가(₩)','엔화(¥)','금액/비율'].includes(second)) {
      cells.forEach(c => { const th = document.createElement('th'); th.textContent = c||''; tr.appendChild(th); });
      table.appendChild(tr); return;
    }

    const isSub = first.includes('소계')||second.includes('소계')||second.includes('합계');
    const isProfit = first==='이익'||second.includes('순이익')||first.includes('이익');
    if (isSub) tr.classList.add('subtotal');
    if (isProfit) tr.classList.add('profit-row');

    cells.forEach((c,i) => {
      const td = document.createElement('td');
      if (typeof c === 'number') {
        td.className = 'num';
        td.textContent = fmt(c);
      } else {
        td.textContent = c||'';
      }
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  container.innerHTML = '';
  div.appendChild(table);
  container.appendChild(div);
}

// Dashboard - render as cards from 종합_대시보드 sheet
function renderDashboard(rows) {
  const area = document.getElementById('dashboard-area');
  area.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'summary-cards';

  let currentCard = null, cardDiv = null;
  rows.forEach(cells => {
    if (!cells) return;
    const c0 = (cells[0]||'').toString();
    const c1 = cells[1];
    const c2 = cells[2];

    if (c0.startsWith('[')) {
      // New card
      if (cardDiv) grid.appendChild(cardDiv);
      cardDiv = document.createElement('div');
      cardDiv.className = 'card' + (c0.includes('한국')?' kr':' jp2');
      const h3 = document.createElement('h3');
      h3.textContent = c0.replace(/[\[\]]/g,'').trim();
      cardDiv.appendChild(h3);
      return;
    }
    if (!cardDiv || !c0 || c0.startsWith('on&off') || ['항목',''].includes(c0)) return;

    const row = document.createElement('div');
    row.className = 'card-row';
    if (c0.includes('순이익')||c0==='이익') row.classList.add('total');

    const label = document.createElement('span');
    label.textContent = c0;
    const val = document.createElement('span');
    val.className = 'amount';
    if (c2 !== null && c2 !== undefined) {
      val.textContent = typeof c1==='number' ? fmtWon(c2) : (c1||'');
      // Show both yen and won for japan
      if (typeof c1==='number' && typeof c2==='number') {
        val.textContent = fmtWon(c2);
      }
    } else {
      val.textContent = typeof c1==='number' ? fmtWon(c1) : (c1||'');
    }
    row.appendChild(label);
    row.appendChild(val);
    cardDiv.appendChild(row);
  });
  if (cardDiv) grid.appendChild(cardDiv);
  area.appendChild(grid);
}

// Distribution - render as cards from 수익배분 sheet
function renderDistribution(rows) {
  const area = document.getElementById('distribution-area');
  area.innerHTML = '';

  // Parse data by section + label
  let section = '';
  const d = {};
  rows.forEach(cells => {
    if (!cells) return;
    const c0 = (cells[0]||'').toString();
    const c1 = (cells[1]||'').toString();
    const c2 = cells[2];
    if (c0.startsWith('[')) { section = c0; return; }

    if (section.includes('배분 비율')) {
      if (c0.includes('한국')&&c1==='석필름') d.krSukRate=c2;
      if (c1==='헤븐리') d.krHevRate=c2;
      if (c0.includes('일본')&&c1==='석필름') d.jpSukRate=c2;
      if (c1==='IMX'&&!section.includes('자동')) d.jpImxRate=c2;
    }
    if (section.includes('게런티')&&section.includes('한국')) {
      if (c1==='한지우') d.krActorA=c2;
      if (c1==='조윤') d.krActorB=c2;
    }
    if (section.includes('게런티')&&section.includes('일본')) {
      if (c1==='한지우') d.jpActorA=c2;
      if (c1==='조윤') d.jpActorB=c2;
    }
    if (section.includes('자동계산')&&section.includes('한국')) {
      if (c1==='석필름 배분') d.krSuk=c2;
      if (c1==='헤븐리 배분') d.krHev=c2;
      if (c1.includes('중계')&&c1.includes('석필름')) d.krVodSuk=c2;
      if (c1.includes('중계')&&c1.includes('헤븐리')) d.krVodHev=c2;
    }
    if (section.includes('자동계산')&&section.includes('일본')) {
      if (c1==='석필름 배분') d.jpSuk=c2;
      if (c1==='IMX 배분') d.jpImx=c2;
    }
    if (section.includes('중계')&&section.includes('한국')) {
      if (c1==='석필름') d.krVodSuk=c2;
      if (c1==='헤븐리') d.krVodHev=c2;
    }
    if (section.includes('중계')&&section.includes('일본')) {
      if (c1.includes('소계')) d.jpVod=c2;
    }
  });

  const grid = document.createElement('div');
  grid.className = 'summary-cards';

  // 한국 카드
  const kr = makeCard('한국 (헤븐리)', 'kr', [
    ['석필름', d.krSuk],
    ['헤븐리', d.krHev],
    ['중계/VOD 석필름', d.krVodSuk],
    ['중계/VOD 헤븐리', d.krVodHev],
    ['한지우', d.krActorA],
    ['조윤', d.krActorB],
  ], `배분 ${d.krSukRate||50} : ${d.krHevRate||50}`);

  // 일본 카드
  const jp = makeCard('일본 (IMX) - 와테라스', 'jp2', [
    ['석필름', d.jpSuk],
    ['IMX', d.jpImx],
    ['중계/VOD (석필름 단독)', d.jpVod],
    ['한지우', d.jpActorA],
    ['조윤', d.jpActorB],
  ], `배분 ${d.jpSukRate||50} : ${d.jpImxRate||50}`);

  grid.appendChild(kr);
  grid.appendChild(jp);
  area.appendChild(grid);

  const note = document.createElement('p');
  note.className = 'note';
  note.textContent = '※ 배분 비율·배우 게런티는 스프레드시트 수익배분 탭에서 수정';
  area.appendChild(note);
}

function makeCard(title, cls, items, subtitle) {
  const card = document.createElement('div');
  card.className = `card ${cls}`;
  const h3 = document.createElement('h3');
  h3.textContent = title;
  card.appendChild(h3);

  items.forEach(([label, value]) => {
    const row = document.createElement('div');
    row.className = 'card-row';
    row.innerHTML = `<span>${label}</span><span class="amount">${fmtWon(value||0)}</span>`;
    card.appendChild(row);
  });

  if (subtitle) {
    const sub = document.createElement('div');
    sub.className = 'card-row sub';
    sub.innerHTML = `<span>${subtitle}</span>`;
    card.appendChild(sub);
  }
  return card;
}

// Collapsible
document.querySelectorAll('.detail-section h2').forEach(h2 => {
  h2.addEventListener('click', () => {
    const el = h2.nextElementSibling;
    if (el) { el.style.display = el.style.display==='none'?'block':'none'; }
  });
});

async function refreshData() {
  document.getElementById('last-update').textContent = '로딩 중...';
  await fetchExchangeRate();

  const [kr, jp, md, dist, dash] = await Promise.all([
    fetchSheet('한국_견적'),
    fetchSheet('일본_견적_와테라스'),
    fetchSheet('MD_판매'),
    fetchSheet('수익배분'),
    fetchSheet('종합_대시보드'),
  ]);

  renderDashboard(dash);
  renderDistribution(dist);
  renderTable(kr, document.getElementById('kr-detail'));
  renderTable(jp, document.getElementById('jp-detail'));
  renderTable(md, document.getElementById('md-detail'));

  document.getElementById('last-update').textContent = `마지막 업데이트: ${new Date().toLocaleString('ko-KR')}`;
}

refreshData();
