const SID = '1o5XdAT1wg8Jv5Boj6rdy_klGJHlB_am4E7eZXs9S15s';

const f = n => {
  if (n==null||n===''||isNaN(n)) return '-';
  const v = Math.round(Number(n));
  return (v<0?'-₩':'₩') + Math.abs(v).toLocaleString('ko-KR');
};
const n = v => {
  if (v==null||v===''||isNaN(v)) return '-';
  return Math.round(Number(v)).toLocaleString('ko-KR');
};

async function load(sheet) {
  try {
    const r = await fetch(`https://docs.google.com/spreadsheets/d/${SID}/gviz/tq?tqx=out:json&headers=0&sheet=${encodeURIComponent(sheet)}`);
    const t = await r.text();
    const m = t.match(/google\.visualization\.Query\.setResponse\(([\s\S]+)\);/);
    if (!m) return [];
    const j = JSON.parse(m[1]);
    return j.table.rows.map(r => r.c ? r.c.map(c => c ? (c.v ?? null) : null) : []);
  } catch(e) { console.error(sheet, e); return []; }
}

async function init() {
  document.getElementById('status').textContent = '로딩 중...';

  // 환율
  try {
    const r = await fetch('https://open.er-api.com/v6/latest/JPY');
    const d = await r.json();
    document.getElementById('rate').textContent = `1¥ = ${d.rates.KRW.toFixed(1)}₩`;
  } catch(e) { document.getElementById('rate').textContent = '환율 로드 실패'; }

  const [dash, dist, kr, jp, md] = await Promise.all([
    load('종합_대시보드'), load('수익배분'),
    load('한국_견적'), load('일본_견적_와테라스'), load('MD_판매')
  ]);

  // 종합 대시보드 카드
  let html = '<div class="grid">';
  let card = '', sec = '';
  dash.forEach(row => {
    const a = (row[0]||'').toString(), b = row[1], c = row[2];
    if (a.startsWith('[')) {
      if (card) html += card + '</div>';
      const cls = a.includes('한국') ? 'orange' : a.includes('일본') ? 'green' : 'blue';
      card = `<div class="card ${cls}"><div class="card-title">${a.replace(/[\[\]]/g,'')}</div>`;
      return;
    }
    if (!card || a === '항목' || a.startsWith('on&off') || !a) return;
    const isTotal = a.includes('순이익') || a === '이익';
    const val = c != null && typeof b === 'number' ? f(c) : (typeof b === 'number' ? f(b) : (b||''));
    card += `<div class="row${isTotal?' total':''}"><span>${a}</span><span>${val}</span></div>`;
  });
  if (card) html += card + '</div>';
  html += '</div>';
  document.getElementById('summary').innerHTML = html;

  // 수익 배분
  let section = '';
  const D = {};
  dist.forEach(row => {
    const a = (row[0]||'').toString(), b = (row[1]||'').toString(), c = row[2];
    if (a.startsWith('[')) { section = a; return; }
    if (section.includes('배분 비율')) {
      if (a.includes('한국') && b === '석필름') D.ks = c;
      if (b === '헤븐리' && !section.includes('자동')) D.kh = c;
      if (a.includes('일본') && b === '석필름') D.js = c;
      if (b === 'IMX' && !section.includes('자동')) D.ji = c;
    }
    if (section.includes('게런티') && section.includes('한국')) {
      if (b === '한지우') D.ka1 = c; if (b === '조윤') D.ka2 = c;
    }
    if (section.includes('게런티') && section.includes('일본')) {
      if (b === '한지우') D.ja1 = c; if (b === '조윤') D.ja2 = c;
    }
    if (section.includes('자동계산') && section.includes('한국')) {
      if (b === '석필름 배분') D.kSuk = c;
      if (b === '헤븐리 배분') D.kHev = c;
      if (b.includes('중계') && b.includes('석필름')) D.kvs = c;
      if (b.includes('중계') && b.includes('헤븐리')) D.kvh = c;
    }
    if (section.includes('자동계산') && section.includes('일본')) {
      if (b === '석필름 배분') D.jSuk = c;
      if (b === 'IMX 배분') D.jImx = c;
    }
    if (section.includes('중계') && section.includes('한국')) {
      if (b === '석필름') D.kvs = c; if (b === '헤븐리') D.kvh = c;
    }
    if (section.includes('중계') && section.includes('일본')) {
      if (b.includes('소계')) D.jv = c;
    }
  });

  const distHtml = `<div class="grid">
    <div class="card orange"><div class="card-title">한국 (헤븐리)</div>
      <div class="row"><span>석필름</span><span>${f(D.kSuk)}</span></div>
      <div class="row"><span>헤븐리</span><span>${f(D.kHev)}</span></div>
      <div class="row"><span>중계/VOD 석필름</span><span>${f(D.kvs)}</span></div>
      <div class="row"><span>중계/VOD 헤븐리</span><span>${f(D.kvh)}</span></div>
      <div class="row"><span>한지우</span><span>${f(D.ka1)}</span></div>
      <div class="row"><span>조윤</span><span>${f(D.ka2)}</span></div>
      <div class="sub">배분 ${D.ks||50} : ${D.kh||50}</div>
    </div>
    <div class="card green"><div class="card-title">일본 (IMX) - 와테라스</div>
      <div class="row"><span>석필름</span><span>${f(D.jSuk)}</span></div>
      <div class="row"><span>IMX</span><span>${f(D.jImx)}</span></div>
      <div class="row"><span>중계/VOD (석필름 단독)</span><span>${f(D.jv)}</span></div>
      <div class="row"><span>한지우</span><span>${f(D.ja1)}</span></div>
      <div class="row"><span>조윤</span><span>${f(D.ja2)}</span></div>
      <div class="sub">배분 ${D.js||50} : ${D.ji||50}</div>
    </div>
  </div>
  <p class="note">※ 배분 비율·배우 게런티는 스프레드시트 수익배분 탭에서 수정</p>`;
  document.getElementById('dist').innerHTML = distHtml;

  // 상세 테이블
  document.getElementById('kr-detail').innerHTML = makeTable(kr);
  document.getElementById('jp-detail').innerHTML = makeTable(jp);
  document.getElementById('md-detail').innerHTML = makeTable(md);

  document.getElementById('status').textContent = new Date().toLocaleString('ko-KR') + ' 업데이트';
}

function makeTable(rows) {
  let h = '<table>';
  rows.forEach(cells => {
    if (!cells || cells.every(c => c == null)) return;
    const a = (cells[0]||'').toString(), b = (cells[1]||'').toString();
    if (a.startsWith('[')) {
      h += `<tr><td colspan="${cells.length}" class="sec">${a}</td></tr>`;
      return;
    }
    const isHeader = ['구분','항목','좌석 종류','상품명'].includes(a) ||
      ['소구분','항목','가격(₩)','단가(₩)','판매원가(¥)','판매원가(₩)','엔화(¥)','금액/비율'].includes(b);
    const isSub = a.includes('소계')||b.includes('소계')||b.includes('합계');
    const isProfit = a==='이익'||b.includes('순이익');

    if (isHeader) {
      h += '<tr class="th">' + cells.map(c => `<td>${c||''}</td>`).join('') + '</tr>';
    } else {
      h += `<tr class="${isSub?'sub':''} ${isProfit?'profit':''}">` +
        cells.map((c,i) => `<td${typeof c==='number'?' class="r"':''}>${typeof c==='number'?n(c):(c||'')}</td>`).join('') + '</tr>';
    }
  });
  return h + '</table>';
}

// Toggle
document.querySelectorAll('.toggle').forEach(el => {
  el.querySelector('h2').addEventListener('click', () => {
    const c = el.querySelector('.content');
    c.style.display = c.style.display === 'none' ? 'block' : 'none';
  });
});

init();
