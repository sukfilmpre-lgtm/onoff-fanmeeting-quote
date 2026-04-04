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

function init(forceLive) {
  document.getElementById('status').textContent = '로딩 중...';

  // 환율
  fetch('https://open.er-api.com/v6/latest/JPY').then(function(r) { return r.json(); }).then(function(d) {
    document.getElementById('rate').textContent = '1\u00A5 = ' + d.rates.KRW.toFixed(1) + '\u20A9';
  }).catch(function() {
    document.getElementById('rate').textContent = '환율 로드 실패';
  });

  // 모든 스냅샷 로드 + latest 체크
  loadAllSnapshots();

  if (forceLive) {
    _activeTab = '';
    renderTabs();
    loadLive();
    return;
  }

  ghRead(GH_PATH + 'latest.json').then(function(json) {
    if (json && json.data) {
      if (json.opts) {
        document.getElementById('chk-video').checked = json.opts.video || false;
        document.getElementById('chk-rs').checked = json.opts.rs || false;
      }
      _lastData = json.data;
      _activeTab = json.name || '';
      renderAll(json.data[0], json.data[1], json.data[2], json.data[3], json.data[4]);
      document.getElementById('snap-label').textContent = '(' + json.name + ')';
      document.getElementById('status').textContent = '"' + json.name + '" (저장본) · ' + new Date(json.date).toLocaleString('ko-KR');
      setTimeout(renderTabs, 500);
    } else {
      document.getElementById('summary').innerHTML = '<p class="loading">저장된 견적이 없습니다. 파일관리에서 실시간 불러오기 후 저장하세요.</p>';
      document.getElementById('dist').innerHTML = '';
    }
  });
}

function rerender() {
  if (_lastData) renderAll(_lastData[0], _lastData[1], _lastData[2], _lastData[3], _lastData[4]);
}

function loadLiveAndClose() {
  document.getElementById('admin-modal').style.display = 'none';
  _activeTab = '';
  renderTabs();
  document.getElementById('snap-label').textContent = '(실시간)';
  loadLive();
}

function loadLive() {
  document.getElementById('snap-label').textContent = '';
  Promise.all([
    load('종합_대시보드'), load('수익배분'),
    load('한국_견적'), load('일본_견적_와테라스'), load('MD_판매')
  ]).then(function(results) {
    _lastData = results;
    renderAll(results[0], results[1], results[2], results[3], results[4]);
  }).catch(function(e) {
    console.error('Init error:', e);
    document.getElementById('status').textContent = '에러: ' + e.message;
  });
}

function renderAll(dash, dist, kr, jp, md) {
    // 티켓 판매수 파싱
    var krTickets = 0, jpTickets = 0;
    kr.forEach(function(cells) {
      if (!cells) return;
      var a = String(cells[0] || '');
      if (a === '티켓 소계' && typeof cells[2] === 'number') krTickets = cells[2];
    });
    jp.forEach(function(cells) {
      if (!cells) return;
      var a = String(cells[0] || '');
      if (a === '티켓 소계' && typeof cells[2] === 'number') jpTickets = cells[2];
    });
    var totalTickets = krTickets + jpTickets;

    // 종합 대시보드 - 한국/일본 동일 형식 카드 (수익배분 섹션 제외)
    var html = '';
    if (totalTickets > 0) {
      html += '<div class="ticket-summary">총 티켓 <b>' + fn(totalTickets) + '매</b> (한국 ' + fn(krTickets) + ' + 일본 ' + fn(jpTickets) + ')</div>';
    }
    html += '<div class="grid">';
    var card = '', skip = false;
    dash.forEach(function(cells) {
      if (!cells || !cells[0]) return;
      var a = String(cells[0] || ''), b = cells[1], c = cells[2];
      if (a.startsWith('[')) {
        if (a.includes('수익 배분') || a.includes('배분')) { skip = true; return; }
        skip = false;
        if (card) html += card + '</div>';
        var cls = a.includes('한국') ? 'orange' : 'green';
        var tickets = a.includes('한국') ? krTickets : jpTickets;
        card = '<div class="card ' + cls + '"><div class="card-title">' + a.replace(/[\[\]]/g, '').trim() + '<span class="ticket-badge">' + fn(tickets) + '매</span></div>';
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
    var useRS = document.getElementById('chk-rs').checked;
    var sukKr = (D.kSuk||0) + (D.kvs||0);
    var sukJp = (D.jSuk||0) + (D.jv||0);
    // RS: 석필름 순이익의 5%/인 (2명) = 10%를 석필름에서 차감
    var sukBeforeRS = sukKr + sukJp - videoCost;
    var rsKr = 0, rsJp = 0, rsTotal = 0;
    if (useRS && sukBeforeRS > 0) {
      rsKr = sukKr > 0 ? sukKr * 0.05 : 0;
      rsJp = sukJp > 0 ? sukJp * 0.05 : 0;
      rsTotal = (rsKr + rsJp) * 2;
      // RS 적용 후에도 마이너스면 RS 취소
      if (sukBeforeRS - rsTotal < 0) { rsKr = 0; rsJp = 0; rsTotal = 0; }
    }
    var sukTotal = sukBeforeRS - rsTotal;
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
    if (useRS) {
      dh += '<div class="total-row"><span>배우 RS (5%x2명)</span><span>-' + fw(rsTotal) + '</span></div>';
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
    dh += '<div class="total-row"><span>한국 (고정)</span><span>' + fw(a1Kr) + '</span></div>';
    dh += '<div class="total-row"><span>일본 (고정)</span><span>' + fw(a1Jp) + '</span></div>';
    if (useRS) {
      dh += '<div class="total-row"><span>RS 한국 (5%)</span><span>' + fw(rsKr) + '</span></div>';
      dh += '<div class="total-row"><span>RS 일본 (5%)</span><span>' + fw(rsJp) + '</span></div>';
    }
    dh += '<div class="total-row sum"><span>합계</span><span>' + fw(a1Total + (useRS ? rsKr + rsJp : 0)) + '</span></div>';
    dh += '</div>';

    // 조윤
    dh += '<div class="total-col">';
    dh += '<div class="total-label">조윤</div>';
    dh += '<div class="total-row"><span>한국 (고정)</span><span>' + fw(a2Kr) + '</span></div>';
    dh += '<div class="total-row"><span>일본 (고정)</span><span>' + fw(a2Jp) + '</span></div>';
    if (useRS) {
      dh += '<div class="total-row"><span>RS 한국 (5%)</span><span>' + fw(rsKr) + '</span></div>';
      dh += '<div class="total-row"><span>RS 일본 (5%)</span><span>' + fw(rsJp) + '</span></div>';
    }
    dh += '<div class="total-row sum"><span>합계</span><span>' + fw(a2Total + (useRS ? rsKr + rsJp : 0)) + '</span></div>';
    dh += '</div>';

    dh += '</div></div>';

    dh += '<p class="note">※ 배분 비율·배우 게런티는 스프레드시트 수익배분 탭에서 수정</p>';
    document.getElementById('dist').innerHTML = dh;

    // 상세 테이블
    document.getElementById('kr-detail').innerHTML = tbl(kr, 'kr');
    document.getElementById('jp-detail').innerHTML = tbl(jp, 'jp');
    document.getElementById('md-detail').innerHTML = tbl(md, 'md');

    document.getElementById('status').textContent = new Date().toLocaleString('ko-KR') + ' 업데이트';
}

function makeDistCard(title, cls, items, sub) {
  var h = '<div class="card ' + cls + '"><div class="card-title">' + title + '</div>';
  items.forEach(function(item) {
    h += '<div class="row"><span>' + item[0] + '</span><span>' + fw(item[1] || 0) + '</span></div>';
  });
  h += '<div class="sub">' + sub + '</div></div>';
  return h;
}

function tbl(rows, type) {
  var h = '<table>';
  var jpSection = '';
  rows.forEach(function(cells) {
    if (!cells || cells.every(function(c) { return c == null; })) return;
    var a = String(cells[0] || ''), b = String(cells[1] || '');
    if (a.startsWith('[')) {
      jpSection = a;
      h += '<tr><td colspan="' + cells.length + '" class="sec">' + a + '</td></tr>';
      // 일본 수익 요약 섹션 뒤에 헤더 행 삽입
      if (type === 'jp' && a.indexOf('수익 요약') >= 0) {
        h += '<tr class="th"><td>항목</td><td>엔화(¥)</td><td></td><td></td><td>원화(₩)</td><td></td></tr>';
      }
      return;
    }
    var isH = ['구분','항목','좌석 종류','상품명'].indexOf(a) >= 0 ||
      ['소구분','항목','가격(₩)','단가(₩)','판매원가(¥)','판매원가(₩)','엔화(¥)','금액/비율'].indexOf(b) >= 0;
    var isS = a.indexOf('소계') >= 0 || b.indexOf('소계') >= 0 || b.indexOf('합계') >= 0;
    var isP = a === '이익' || b.indexOf('순이익') >= 0;
    // 일본 시트에서 석필름/IMX 배분 행 제거
    if (type === 'jp' && (a.indexOf('석필름') >= 0 || a.indexOf('IMX') >= 0)) return;

    if (isH) {
      // 일본 시트: 빈 헤더 칼럼에 단위 채우기
      if (type === 'kr') {
        if (a === '좌석 종류') cells = ['좌석 종류','가격(₩)','매수','금액(₩)','','','비고'];
        if (a === '구분' && b === '단가(₩)') cells = ['구분','단가(₩)','판매수','금액(₩)','','','비고'];
        if (a === '구분' && b === '소구분') cells = ['구분','소구분','단가','단위','수량','금액','비고'];
      }
      if (type === 'jp') {
        if (a === '좌석 종류') cells = ['좌석 종류','가격(¥)','매수','엔화(¥)','원화(₩)',''];
        if (a === '항목' && jpSection.indexOf('지출A') >= 0) cells = ['항목','엔화(¥)','','','원화(₩)',''];
        if (a === '항목' && jpSection.indexOf('지출B') >= 0) cells = ['항목','단위','횟수','단가(¥)','엔화(¥)','원화(₩)'];
        if (a === '항목' && jpSection.indexOf('수익') >= 0) cells = ['','엔화(¥)','','','원화(₩)',''];
      }
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

var _lastData = null;
var _allSnapshots = {};
var _activeTab = 'live';
var GH_TOKEN = localStorage.getItem('gh_token') || '';
var GH_REPO = 'sukfilmpre-lgtm/onoff-fanmeeting-quote';
var GH_PATH = 'snapshots/';

function ghApi(path, method, body) {
  var opts = {method: method || 'GET', headers: {'Authorization': 'Bearer ' + GH_TOKEN, 'Accept': 'application/vnd.github+json'}};
  if (body) { opts.body = JSON.stringify(body); opts.headers['Content-Type'] = 'application/json'; }
  return fetch('https://api.github.com/repos/' + GH_REPO + '/contents/' + path, opts).then(function(r) { return r.json(); });
}

function ghRead(path) {
  return fetch('https://api.github.com/repos/' + GH_REPO + '/contents/' + path)
    .then(function(r) { return r.json(); })
    .then(function(f) {
      if (!f || !f.content) return null;
      return JSON.parse(decodeURIComponent(escape(atob(f.content.replace(/\n/g, '')))));
    }).catch(function() { return null; });
}

function ensureToken() {
  if (!GH_TOKEN) {
    var t = prompt('GitHub 토큰을 입력하세요 (최초 1회):');
    if (!t) return false;
    GH_TOKEN = t.trim();
    localStorage.setItem('gh_token', GH_TOKEN);
  }
  return true;
}

var PASS = '1234';
function checkAdmin() {
  if (sessionStorage.getItem('authed')) return true;
  var p = prompt('관리자 비밀번호를 입력하세요:');
  if (p !== PASS) { alert('비밀번호가 틀렸습니다.'); return false; }
  sessionStorage.setItem('authed', '1');
  return ensureToken();
}

// === 탭 로드 ===
function loadAllSnapshots() {
  fetch('https://api.github.com/repos/' + GH_REPO + '/contents/' + GH_PATH)
    .then(function(r) { return r.json(); })
    .then(function(files) {
      if (!Array.isArray(files)) return;
      var jsons = files.filter(function(f) { return f.name.endsWith('.json') && f.name !== 'latest.json'; });
      var promises = jsons.map(function(f) { return ghRead(GH_PATH + f.name); });
      return Promise.all(promises).then(function(results) {
        _allSnapshots = {};
        results.forEach(function(snap) { if (snap && snap.name) _allSnapshots[snap.name] = snap; });
        renderTabs();
      });
    }).catch(function() { renderTabs(); });
}

function renderTabs() {
  var el = document.getElementById('scenario-tabs');
  var names = Object.keys(_allSnapshots);
  if (names.length === 0) { el.innerHTML = ''; return; }
  var h = '';
  names.forEach(function(name) {
    h += '<button class="tab' + (_activeTab === name ? ' active' : '') + '" onclick="switchTab(\'' + name.replace(/'/g, "\\'") + '\')">' + name + '</button>';
  });
  el.innerHTML = h;
}

function switchTab(name) {
  _activeTab = name;
  renderTabs();
  if (name === 'live') {
    document.getElementById('snap-label').textContent = '';
    loadLive();
  } else {
    var snap = _allSnapshots[name];
    if (!snap) return;
    if (snap.opts) {
      document.getElementById('chk-video').checked = snap.opts.video || false;
      document.getElementById('chk-rs').checked = snap.opts.rs || false;
    }
    _lastData = snap.data;
    renderAll(snap.data[0], snap.data[1], snap.data[2], snap.data[3], snap.data[4]);
    document.getElementById('snap-label').textContent = '(' + name + ')';
    document.getElementById('status').textContent = '"' + name + '" · ' + new Date(snap.date).toLocaleString('ko-KR');
  }
}

// === 저장 (비번 없음) ===
function saveSnapshot() {
  if (!_lastData) { alert('데이터를 먼저 로드하세요.'); return; }
  if (!ensureToken()) return;
  var name = prompt('저장 이름을 입력하세요 (예: 전석매진, 70%판매):');
  if (!name || !name.trim()) return;
  name = name.trim();
  var fname = name.replace(/[^a-zA-Z0-9가-힣_-]/g, '_') + '.json';
  var payload = {
    name: name, date: new Date().toISOString(), data: _lastData,
    opts: { video: document.getElementById('chk-video').checked, rs: document.getElementById('chk-rs').checked }
  };
  var content = btoa(unescape(encodeURIComponent(JSON.stringify(payload, null, 2))));
  ghApi(GH_PATH + fname).then(function(existing) {
    var body = { message: '견적 스냅샷: ' + name, content: content };
    if (existing && existing.sha) body.sha = existing.sha;
    return ghApi(GH_PATH + fname, 'PUT', body);
  }).then(function(r) {
    if (!r.content) { alert('저장 실패: ' + (r.message || '')); return; }
    // latest.json 업데이트
    ghApi(GH_PATH + 'latest.json').then(function(ex) {
      var body = { message: '최신: ' + name, content: content };
      if (ex && ex.sha) body.sha = ex.sha;
      return ghApi(GH_PATH + 'latest.json', 'PUT', body);
    }).then(function() {
      alert('"' + name + '" 저장 완료');
      _allSnapshots[name] = payload;
      _activeTab = name;
      renderTabs();
      document.getElementById('snap-label').textContent = '(' + name + ')';
    });
  }).catch(function(e) { alert('저장 에러: ' + e.message); });
}

// === 관리 (비번 필요) ===
function openAdmin() {
  if (!checkAdmin()) return;
  var list = document.getElementById('admin-list');
  list.innerHTML = '<div class="snap-empty">불러오는 중...</div>';
  document.getElementById('admin-modal').style.display = 'flex';
  ghApi(GH_PATH).then(function(files) {
    if (!Array.isArray(files)) { list.innerHTML = '<div class="snap-empty">저장된 견적이 없습니다</div>'; return; }
    var jsons = files.filter(function(f) { return f.name.endsWith('.json') && f.name !== 'latest.json'; });
    if (jsons.length === 0) { list.innerHTML = '<div class="snap-empty">저장된 견적이 없습니다</div>'; return; }
    list.innerHTML = jsons.map(function(f) {
      return '<div class="snap-item">' +
        '<div class="snap-name">' + f.name.replace('.json', '') + '</div>' +
        '<button class="snap-del" onclick="delSnapshot(\'' + f.name + '\')">삭제</button>' +
        '</div>';
    }).join('');
  }).catch(function() { list.innerHTML = '<div class="snap-empty">로드 실패</div>'; });
}

function delSnapshot(fname) {
  if (!confirm('"' + fname.replace('.json', '') + '" 삭제하시겠습니까?')) return;
  ghApi(GH_PATH + fname).then(function(f) {
    if (!f || !f.sha) { alert('파일을 찾을 수 없습니다.'); return; }
    return ghApi(GH_PATH + fname, 'DELETE', { message: '삭제: ' + fname, sha: f.sha });
  }).then(function(r) {
    if (r && r.commit) {
      var name = fname.replace('.json', '');
      delete _allSnapshots[name];
      if (_activeTab === name) { _activeTab = 'live'; loadLive(); }
      renderTabs();
      alert('삭제 완료');
      setTimeout(function() { openAdmin(); }, 1500);
    }
  }).catch(function(e) { alert('삭제 실패: ' + e.message); });
}

document.querySelectorAll('.toggle').forEach(function(el) {
  el.querySelector('h2').addEventListener('click', function() {
    var c = el.querySelector('.content');
    if (c) c.style.display = c.style.display === 'none' ? 'block' : 'none';
  });
});

init();
