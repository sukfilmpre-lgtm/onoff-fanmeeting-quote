var SID = '1o5XdAT1wg8Jv5Boj6rdy_klGJHlB_am4E7eZXs9S15s';

function fw(n) {
  if (n == null || n === '' || isNaN(n)) return '-';
  var v = Math.round(Number(n));
  return (v < 0 ? '-' : '') + '\u20A9' + Math.abs(v).toLocaleString('ko-KR');
}
function fn(v) {
  if (v == null || v === '' || isNaN(v)) return '-';
  return Math.round(Number(v)).toLocaleString('ko-KR');
}

function load(sheet) {
  var url = 'https://docs.google.com/spreadsheets/d/' + SID + '/gviz/tq?tqx=out:json&headers=0&sheet=' + encodeURIComponent(sheet);
  return fetch(url).then(function(r) { return r.text(); }).then(function(t) {
    var m = t.match(/setResponse\(([\s\S]+)\);/);
    if (!m) return [];
    var j = JSON.parse(m[1]);
    return j.table.rows.map(function(r) {
      if (!r.c) return [];
      return r.c.map(function(c) { return c ? (c.v !== undefined ? c.v : null) : null; });
    });
  }).catch(function(e) { console.error(sheet, e); return []; });
}

function init() {
  document.getElementById('status').textContent = '로딩 중...';

  // 환율
  fetch('https://open.er-api.com/v6/latest/JPY').then(function(r) { return r.json(); }).then(function(d) {
    document.getElementById('rate').textContent = '1\u00A5 = ' + d.rates.KRW.toFixed(1) + '\u20A9';
  }).catch(function() {
    document.getElementById('rate').textContent = '환율 로드 실패';
  });

  Promise.all([
    load('종합_대시보드'), load('수익배분'),
    load('한국_견적'), load('일본_견적_와테라스'), load('MD_판매')
  ]).then(function(results) {
    var dash = results[0], dist = results[1], kr = results[2], jp = results[3], md = results[4];

    // 종합 대시보드 - 한국/일본 동일 형식 카드 (수익배분 섹션 제외)
    var html = '<div class="grid">';
    var card = '', skip = false;
    dash.forEach(function(cells) {
      if (!cells || !cells[0]) return;
      var a = String(cells[0] || ''), b = cells[1], c = cells[2];
      if (a.startsWith('[')) {
        if (a.includes('수익 배분') || a.includes('배분')) { skip = true; return; }
        skip = false;
        if (card) html += card + '</div>';
        var cls = a.includes('한국') ? 'orange' : 'green';
        card = '<div class="card ' + cls + '"><div class="card-title">' + a.replace(/[\[\]]/g, '').trim() + '</div>';
        return;
      }
      if (skip || !card) return;
      if (a === '항목' || a.startsWith('on&off') || a === '적용 환율' || a === '환율 기준일') return;

      var isTotal = a.includes('순이익') || a === '이익';
      var val;
      if (c != null && typeof c === 'number') val = fw(c);
      else if (typeof b === 'number') val = fw(b);
      else val = b || '';

      card += '<div class="row' + (isTotal ? ' total' : '') + '"><span>' + a + '</span><span>' + val + '</span></div>';
    });
    if (card) html += card + '</div>';
    html += '</div>';
    document.getElementById('summary').innerHTML = html;

    // 수익 배분
    var section = '', D = {};
    dist.forEach(function(cells) {
      if (!cells) return;
      var a = String(cells[0] || ''), b = String(cells[1] || ''), c = cells[2];
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
        if (b.indexOf('중계') >= 0 && b.indexOf('석필름') >= 0) D.kvs = c;
        if (b.indexOf('중계') >= 0 && b.indexOf('헤븐리') >= 0) D.kvh = c;
      }
      if (section.includes('자동계산') && section.includes('일본')) {
        if (b === '석필름 배분') D.jSuk = c;
        if (b === 'IMX 배분') D.jImx = c;
      }
      if (section.includes('중계') && section.includes('한국')) {
        if (b === '석필름') D.kvs = c; if (b === '헤븐리') D.kvh = c;
      }
      if (section.includes('중계') && section.includes('일본')) {
        if (b.indexOf('소계') >= 0) D.jv = c;
      }
    });

    var videoCost = document.getElementById('chk-video').checked ? 5000000 : 0;
    var sukKr = (D.kSuk||0) + (D.kvs||0);
    var sukJp = (D.jSuk||0) + (D.jv||0);
    var sukTotal = sukKr + sukJp - videoCost;
    var hevTotal = (D.kHev||0) + (D.kvh||0);
    var imxTotal = (D.jImx||0);
    var a1Kr = (D.ka1||0), a1Jp = (D.ja1||0), a1Total = a1Kr + a1Jp;
    var a2Kr = (D.ka2||0), a2Jp = (D.ja2||0), a2Total = a2Kr + a2Jp;

    var dh = '<div class="grid">';
    dh += makeDistCard('한국 수익 배분', 'orange', [
      ['석필름 (제작사)', D.kSuk],
      ['헤븐리 (대행사)', D.kHev],
      ['중계/VOD 석필름', D.kvs],
      ['중계/VOD 헤븐리', D.kvh],
      ['한지우 (배우)', D.ka1],
      ['조윤 (배우)', D.ka2]
    ], '배분 ' + (D.ks || 50) + ' : ' + (D.kh || 50));
    dh += makeDistCard('일본 수익 배분', 'green', [
      ['석필름 (제작사)', D.jSuk],
      ['IMX (대행사)', D.jImx],
      ['중계/VOD (석필름 단독)', D.jv],
      ['한지우 (배우)', D.ja1],
      ['조윤 (배우)', D.ja2]
    ], '배분 ' + (D.js || 50) + ' : ' + (D.ji || 50));
    dh += '</div>';

    dh += '<div class="total-box">';
    dh += '<div class="total-grid">';

    // 석필름
    dh += '<div class="total-col">';
    dh += '<div class="total-label">석필름 (제작사)</div>';
    dh += '<div class="total-row"><span>한국</span><span>' + fw(sukKr) + '</span></div>';
    dh += '<div class="total-row"><span>일본</span><span>' + fw(sukJp) + '</span></div>';
    if (videoCost > 0) {
      dh += '<div class="total-row"><span>영상제작비</span><span>-' + fw(videoCost) + '</span></div>';
    }
    dh += '<div class="total-row sum"><span>합계</span><span>' + fw(sukTotal) + '</span></div>';
    dh += '</div>';

    // 대행사
    dh += '<div class="total-col">';
    dh += '<div class="total-label">대행사</div>';
    dh += '<div class="total-row"><span>헤븐리 (한국)</span><span>' + fw(hevTotal) + '</span></div>';
    dh += '<div class="total-row"><span>IMX (일본)</span><span>' + fw(imxTotal) + '</span></div>';
    dh += '</div>';

    // 한지우
    dh += '<div class="total-col">';
    dh += '<div class="total-label">한지우</div>';
    dh += '<div class="total-row"><span>한국</span><span>' + fw(a1Kr) + '</span></div>';
    dh += '<div class="total-row"><span>일본</span><span>' + fw(a1Jp) + '</span></div>';
    dh += '<div class="total-row sum"><span>합계</span><span>' + fw(a1Total) + '</span></div>';
    dh += '</div>';

    // 조윤
    dh += '<div class="total-col">';
    dh += '<div class="total-label">조윤</div>';
    dh += '<div class="total-row"><span>한국</span><span>' + fw(a2Kr) + '</span></div>';
    dh += '<div class="total-row"><span>일본</span><span>' + fw(a2Jp) + '</span></div>';
    dh += '<div class="total-row sum"><span>합계</span><span>' + fw(a2Total) + '</span></div>';
    dh += '</div>';

    dh += '</div></div>';

    dh += '<p class="note">※ 배분 비율·배우 게런티는 스프레드시트 수익배분 탭에서 수정</p>';
    document.getElementById('dist').innerHTML = dh;

    // 상세 테이블
    document.getElementById('kr-detail').innerHTML = tbl(kr);
    document.getElementById('jp-detail').innerHTML = tbl(jp);
    document.getElementById('md-detail').innerHTML = tbl(md);

    document.getElementById('status').textContent = new Date().toLocaleString('ko-KR') + ' 업데이트';
  }).catch(function(e) {
    console.error('Init error:', e);
    document.getElementById('status').textContent = '에러: ' + e.message;
  });
}

function makeDistCard(title, cls, items, sub) {
  var h = '<div class="card ' + cls + '"><div class="card-title">' + title + '</div>';
  items.forEach(function(item) {
    h += '<div class="row"><span>' + item[0] + '</span><span>' + fw(item[1] || 0) + '</span></div>';
  });
  h += '<div class="sub">' + sub + '</div></div>';
  return h;
}

function tbl(rows) {
  var h = '<table>';
  rows.forEach(function(cells) {
    if (!cells || cells.every(function(c) { return c == null; })) return;
    var a = String(cells[0] || ''), b = String(cells[1] || '');
    if (a.startsWith('[')) {
      h += '<tr><td colspan="' + cells.length + '" class="sec">' + a + '</td></tr>';
      return;
    }
    var isH = ['구분','항목','좌석 종류','상품명'].indexOf(a) >= 0 ||
      ['소구분','항목','가격(₩)','단가(₩)','판매원가(¥)','판매원가(₩)','엔화(¥)','금액/비율'].indexOf(b) >= 0;
    var isS = a.indexOf('소계') >= 0 || b.indexOf('소계') >= 0 || b.indexOf('합계') >= 0;
    var isP = a === '이익' || b.indexOf('순이익') >= 0;

    if (isH) {
      h += '<tr class="th">' + cells.map(function(c) { return '<td>' + (c || '') + '</td>'; }).join('') + '</tr>';
    } else {
      var cls = isS ? ' class="sub"' : (isP ? ' class="profit"' : '');
      h += '<tr' + cls + '>' + cells.map(function(c) {
        if (typeof c === 'number') return '<td class="r">' + fn(c) + '</td>';
        return '<td>' + (c || '') + '</td>';
      }).join('') + '</tr>';
    }
  });
  return h + '</table>';
}

document.querySelectorAll('.toggle').forEach(function(el) {
  el.querySelector('h2').addEventListener('click', function() {
    var c = el.querySelector('.content');
    if (c) c.style.display = c.style.display === 'none' ? 'block' : 'none';
  });
});

init();
