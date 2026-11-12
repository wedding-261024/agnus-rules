/**
 * 아뉴스 청년회 회칙 개정안 실시간 투표 — 서버 코드 (Google Apps Script)
 *
 * 처음 한 번만: 편집기 위쪽 [▶ 실행] 옆 함수 선택 칸이 'setup' 인지 확인하고 [▶ 실행] 을 누르세요.
 *   → 회원 / 조항 / 투표기록 / 결과 / 설정 탭이 자동으로 만들어집니다.
 *   (시트 화면을 새로고침하면 위쪽 메뉴에 '🗳 회칙투표' 메뉴도 생깁니다. 거기서도 실행할 수 있습니다.)
 *
 * 화면 주소
 *   투표자:  배포한 /exec 주소
 *   관리자:  /exec 주소 뒤에 ?admin=1
 */

const SHEETS = { MEMBERS: '회원', ARTICLES: '조항', VOTES: '투표기록', RESULTS: '결과', SETTINGS: '설정' };
const CHOICES = { yes: '찬성', no: '반대', abstain: '기권' };
const CACHE_SEC = 21600; // 6시간 (캐시 최대치)
const LIST_CACHE_SEC = 300; // 회원·조항 목록은 5분마다 시트에서 다시 읽음

/* ───────────── 최초 설정 ───────────── */

function setup() {
  // 시트에서 '확장 프로그램 → Apps Script' 로 열었으면 그 시트를, 아니면 새 시트를 만들어 사용
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const savedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ss && savedId) ss = SpreadsheetApp.openById(savedId);
  if (!ss) {
    ss = SpreadsheetApp.create('회칙투표');
    Logger.log('연결된 시트가 없어 새 시트를 만들었습니다 → ' + ss.getUrl());
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  const members = ensureSheet_(ss, SHEETS.MEMBERS, ['이름', '전화번호 뒷 4자리']);
  members.getRange('B:B').setNumberFormat('@'); // 0054 같은 번호의 앞자리 0이 사라지지 않게

  const arts = ensureSheet_(ss, SHEETS.ARTICLES, ['번호', '제목', '한줄 요약', '개정 전', '개정 후']);
  if (arts.getLastRow() < 2) {
    arts.getRange(2, 1, DEFAULT_ARTICLES.length, 5).setValues(DEFAULT_ARTICLES);
  }
  arts.getRange('A:E').setVerticalAlignment('top').setWrap(true);
  arts.setColumnWidths(1, 1, 60); arts.setColumnWidths(2, 1, 200); arts.setColumnWidths(3, 3, 360);

  ensureSheet_(ss, SHEETS.VOTES, ['조항', '제목', '이름', '선택', '투표 시각']);
  ensureSheet_(ss, SHEETS.RESULTS, ['조항', '제목', '찬성', '반대', '기권', '미투표', '참여', '마감 시각']);

  const set = ensureSheet_(ss, SHEETS.SETTINGS, ['항목', '값']);
  if (!set.getRange('B2').getValue()) set.getRange('A2:B2').setValues([['관리자 비밀번호', '1234']]);
  set.getRange('B2').setNumberFormat('@');

  clearListCache_();
  const msg = '설정 완료! 시트에 회원/조항/투표기록/결과/설정 탭이 만들어졌습니다.\n'
    + '회원 탭에 이름과 전화번호 뒷 4자리를 붙여넣은 뒤 웹 앱으로 배포하세요.\n시트 주소: ' + ss.getUrl();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* 편집기에서 실행하면 알림창 없이 로그에만 표시 */ }
}

/** 시트를 열 때 위쪽 메뉴에 '🗳 회칙투표' 메뉴를 추가 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗳 회칙투표')
    .addItem('① 초기 설정 (탭 만들기)', 'setup')
    .addItem('조항 탭을 최신 개정안으로 다시 채우기', 'reloadArticlesFromMenu')
    .addSeparator()
    .addItem('모든 투표 기록 지우기 (리허설 후)', 'resetAllFromMenu')
    .addToUi();
}

/** 조항 탭을 이 코드에 들어 있는 최신 개정안(DEFAULT_ARTICLES)으로 덮어씀 */
function reloadArticles() {
  const sh = sheet_(SHEETS.ARTICLES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  sh.getRange(2, 1, DEFAULT_ARTICLES.length, 5).setValues(DEFAULT_ARTICLES);
  // 진행 상태를 첫 조항 '시작 전'으로 되돌림 (기존 표는 지우지 않음 → 필요하면 '모든 투표 기록 지우기')
  props_().deleteProperty('state');
  cache_().remove('state');
  clearListCache_();
  Logger.log('조항 탭을 최신 개정안(' + DEFAULT_ARTICLES.map(a => a[0] + '조').join(', ') + ')으로 다시 채웠습니다.');
}

function reloadArticlesFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const list = DEFAULT_ARTICLES.map(a => a[0] + '조').join(', ');
  if (ui.alert('조항 탭을 최신 개정안으로 덮어쓸까요?\n(' + list + ')\n\n직접 고친 요약 문구도 기본값으로 돌아갑니다.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  reloadArticles();
  ui.alert('조항 탭을 다시 채웠습니다. 테스트 표가 남아 있다면 [모든 투표 기록 지우기]도 눌러 주세요.');
}

function resetAllFromMenu() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('모든 투표 기록을 지울까요? 되돌릴 수 없습니다.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  resetAll();
  ui.alert('모든 투표 기록을 지웠습니다.');
}

/* ───────────── 화면 열기 (웹 앱) ───────────── */

function doGet(e) {
  const isAdmin = e && e.parameter && e.parameter.admin;
  const page = isAdmin ? 'Admin' : 'Vote';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(isAdmin ? '회칙 투표 · 관리자' : '아뉴스 청년회 회칙 투표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#eef2f7');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ───────────── 공통 도구 ───────────── */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) { return ss_().getSheetByName(name); }
function cache_() { return CacheService.getScriptCache(); }
function props_() { return PropertiesService.getScriptProperties(); }
function norm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }
function pin_(s) { const d = String(s == null ? '' : s).replace(/\D/g, ''); return d.slice(-4).padStart(4, '0'); }
function vkey_(no, name) { return 'v|' + no + '|' + name; }

function clearListCache_() { cache_().removeAll(['members', 'articles', 'adminpw']); }

function members_() {
  const hit = cache_().get('members');
  if (hit) return JSON.parse(hit);
  const rows = sheet_(SHEETS.MEMBERS).getDataRange().getDisplayValues().slice(1);
  const seen = {};
  const list = [];
  rows.forEach(r => {
    const name = norm_(r[0]);
    if (!name || seen[name]) return;
    seen[name] = true;
    list.push({ name: name, pin: pin_(r[1]) });
  });
  cache_().put('members', JSON.stringify(list), LIST_CACHE_SEC);
  return list;
}

function articles_() {
  const hit = cache_().get('articles');
  if (hit) return JSON.parse(hit);
  const rows = sheet_(SHEETS.ARTICLES).getDataRange().getDisplayValues().slice(1);
  const list = rows.filter(r => String(r[0]).trim()).map(r => ({
    no: String(r[0]).trim(), title: r[1], summary: r[2], before: r[3], after: r[4]
  }));
  cache_().put('articles', JSON.stringify(list), LIST_CACHE_SEC);
  return list;
}

function article_(no) { return articles_().find(a => a.no === String(no)) || null; }

function adminPw_() {
  const hit = cache_().get('adminpw');
  if (hit) return hit;
  const pw = String(sheet_(SHEETS.SETTINGS).getRange('B2').getDisplayValue()).trim() || '1234';
  cache_().put('adminpw', pw, LIST_CACHE_SEC);
  return pw;
}

/* 진행 상태: { no: 현재 조항 번호, status: 'ready'(대기) | 'open'(투표 중) | 'closed'(마감) } */
function getState_() {
  const hit = cache_().get('state');
  if (hit) return JSON.parse(hit);
  let st = props_().getProperty('state');
  if (st) st = JSON.parse(st);
  else { const first = articles_()[0]; st = { no: first ? first.no : '', status: 'ready' }; }
  cache_().put('state', JSON.stringify(st), CACHE_SEC);
  return st;
}
function setState_(st) {
  st.at = Date.now();
  props_().setProperty('state', JSON.stringify(st));
  cache_().put('state', JSON.stringify(st), CACHE_SEC);
}

function getVote_(no, name) {
  const key = vkey_(no, name);
  let v = cache_().get(key);
  if (v === null) {
    v = props_().getProperty(key) || '';
    cache_().put(key, v, CACHE_SEC);
  }
  return v ? JSON.parse(v) : null;
}

/* ───────────── 투표자 기능 ───────────── */

function auth_(name, pin) {
  const n = norm_(name), p = pin_(pin);
  const m = members_().find(x => x.name === n && x.pin === p);
  if (!m) throw new Error('이름 또는 전화번호 뒷 4자리가 맞지 않습니다.');
  return m;
}

function voterView_(m) {
  const st = getState_();
  const arts = articles_();
  const idx = arts.findIndex(a => a.no === String(st.no));
  const mine = st.no ? getVote_(st.no, m.name) : null;
  return {
    name: m.name,
    status: st.status,
    no: st.no,
    article: idx >= 0 ? arts[idx] : null,
    order: idx + 1,
    total: arts.length,
    myVote: mine ? mine.c : null
  };
}

function apiLogin(name, pin) { return voterView_(auth_(name, pin)); }
function apiState(name, pin) { return voterView_(auth_(name, pin)); }

function apiVote(name, pin, no, choice) {
  const m = auth_(name, pin);
  if (!CHOICES[choice]) throw new Error('잘못된 선택입니다.');
  const st = getState_();
  if (st.status !== 'open' || String(st.no) !== String(no)) {
    throw new Error('이 조항은 투표가 마감되어 더 이상 바꿀 수 없습니다.');
  }
  const key = vkey_(no, m.name);
  const val = JSON.stringify({ c: choice, t: Date.now() });
  props_().setProperty(key, val);
  cache_().put(key, val, CACHE_SEC);
  return voterView_(m);
}

/* ───────────── 관리자 기능 ───────────── */

function adminAuth_(pw) {
  if (String(pw == null ? '' : pw).trim() !== adminPw_()) throw new Error('관리자 비밀번호가 틀렸습니다.');
}

/** 모든 조항의 투표를 { 조항번호: { 이름: {c, t} } } 형태로 모음 */
function allVotes_() {
  const all = props_().getProperties();
  const out = {};
  Object.keys(all).forEach(k => {
    if (k.indexOf('v|') !== 0) return;
    const parts = k.split('|');
    const no = parts[1], name = parts.slice(2).join('|');
    (out[no] = out[no] || {})[name] = JSON.parse(all[k]);
  });
  const done = {};
  Object.keys(all).forEach(k => { if (k.indexOf('done|') === 0) done[k.slice(5)] = Number(all[k]); });
  return { votes: out, done: done };
}

function tally_(votesOfArticle, members) {
  const c = { yes: 0, no: 0, abstain: 0, none: 0 };
  members.forEach(m => {
    const v = votesOfArticle && votesOfArticle[m.name];
    if (v && c[v.c] !== undefined) c[v.c]++; else c.none++;
  });
  return c;
}

function adminView_() {
  const st = getState_();
  const arts = articles_();
  const mem = members_();
  const all = allVotes_();
  const cur = all.votes[String(st.no)] || {};
  return {
    status: st.status,
    no: st.no,
    voteUrl: ScriptApp.getService().getUrl(),
    sheetUrl: ss_().getUrl(),
    memberCount: mem.length,
    articles: arts.map(a => ({
      no: a.no, title: a.title, summary: a.summary,
      counts: tally_(all.votes[a.no], mem),
      done: all.done[a.no] || null
    })),
    voters: mem.map(m => ({ name: m.name, choice: cur[m.name] ? cur[m.name].c : null }))
  };
}

function apiAdmin(pw) { adminAuth_(pw); return adminView_(); }

function apiAdminAction(pw, action, no) {
  adminAuth_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = getState_();
    const arts = articles_();
    switch (action) {
      case 'open': // 투표 시작 (이미 마감된 조항을 다시 열면 기존 표는 유지되고 수정 가능해짐)
        setState_({ no: st.no, status: 'open' });
        break;
      case 'close': // 마감 → 투표자는 더 이상 수정 불가
        close_(st);
        break;
      case 'next': { // 현재 조항 마감 + 다음 조항 바로 투표 시작
        close_(st);
        const i = arts.findIndex(a => a.no === String(st.no));
        const nxt = arts[i + 1];
        if (nxt) setState_({ no: nxt.no, status: 'open' });
        break;
      }
      case 'select': // 목록에서 특정 조항 고르기 (대기 상태로)
        if (st.status === 'open') close_(st);
        if (!article_(no)) throw new Error('없는 조항입니다.');
        setState_({ no: String(no), status: props_().getProperty('done|' + no) ? 'closed' : 'ready' });
        break;
      case 'reset': { // 현재 조항의 표를 모두 지우고 대기 상태로
        const keys = members_().map(m => vkey_(st.no, m.name));
        const p = props_();
        Object.keys(p.getProperties()).forEach(k => { if (k.indexOf('v|' + st.no + '|') === 0) p.deleteProperty(k); });
        p.deleteProperty('done|' + st.no);
        cache_().removeAll(keys);
        setState_({ no: st.no, status: 'ready' });
        syncSheets_();
        break;
      }
      case 'reload': // 시트에서 회원·조항·비밀번호를 다시 읽기
        clearListCache_();
        break;
      case 'sync':
        syncSheets_();
        break;
      default:
        throw new Error('알 수 없는 동작입니다.');
    }
  } finally {
    lock.releaseLock();
  }
  return adminView_();
}

function close_(st) {
  if (st.status !== 'open') return;
  setState_({ no: st.no, status: 'closed' });
  Utilities.sleep(1500); // 마감 직전에 눌린 표가 저장될 시간을 줌
  props_().setProperty('done|' + st.no, String(Date.now()));
  syncSheets_();
}

/** 투표기록·결과 탭을 현재 데이터로 다시 씀 */
function syncSheets_() {
  const arts = articles_();
  const mem = members_();
  const all = allVotes_();
  const tz = Session.getScriptTimeZone();
  const fmt = t => t ? Utilities.formatDate(new Date(Number(t)), tz, 'yyyy-MM-dd HH:mm:ss') : '';

  const voteRows = [['조항', '제목', '이름', '선택', '투표 시각']];
  const resultRows = [['조항', '제목', '찬성', '반대', '기권', '미투표', '참여', '마감 시각']];
  arts.forEach(a => {
    const v = all.votes[a.no];
    if (!v && !all.done[a.no]) return; // 아직 진행하지 않은 조항
    mem.forEach(m => {
      const x = v && v[m.name];
      voteRows.push([a.no, a.title, m.name, x ? CHOICES[x.c] : '미투표', x ? fmt(x.t) : '']);
    });
    const c = tally_(v, mem);
    resultRows.push([a.no, a.title, c.yes, c.no, c.abstain, c.none,
      (mem.length - c.none) + ' / ' + mem.length, fmt(all.done[a.no])]);
  });
  writeSheet_(SHEETS.VOTES, voteRows);
  writeSheet_(SHEETS.RESULTS, resultRows);
}

function writeSheet_(name, rows) {
  const sh = sheet_(name);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
}

/** 전체 결과 (CSV 다운로드용) */
function apiExport(pw) {
  adminAuth_(pw);
  const arts = articles_(), mem = members_(), all = allVotes_();
  const rows = [['조항', '제목', '이름', '선택']];
  arts.forEach(a => {
    const v = all.votes[a.no];
    if (!v && !all.done[a.no]) return;
    mem.forEach(m => rows.push([a.no, a.title, m.name, v && v[m.name] ? CHOICES[v[m.name].c] : '미투표']));
  });
  return rows;
}

/** 리허설 후 전체 초기화: 편집기에서 직접 실행 (모든 표가 지워집니다) */
function resetAll() {
  const p = props_();
  Object.keys(p.getProperties()).forEach(k => {
    if (k.indexOf('v|') === 0 || k.indexOf('done|') === 0 || k === 'state') p.deleteProperty(k);
  });
  cache_().remove('state');
  const keys = [];
  articles_().forEach(a => members_().forEach(m => keys.push(vkey_(a.no, m.name))));
  for (let i = 0; i < keys.length; i += 500) cache_().removeAll(keys.slice(i, i + 500));
  syncSheets_();
  Logger.log('모든 투표 기록을 지웠습니다.');
}

/* ───────────── 기본 조항 데이터 (setup 때 '조항' 탭에 한 번 채워짐) ───────────── */
// [번호, 제목, 한줄 요약, 개정 전, 개정 후]  — 회칙 개정안 v1.4 기준
// '개정 후'의 [[ ]] 부분 = 개정안 PDF에서 빨간색으로 표시된 핵심 변경 내용 (투표 화면에서 형광펜으로 표시)
const DEFAULT_ARTICLES = [
 [
  6,
  "제6조 (자격)",
  "정규 청년이 되는 기준이 '입단 후 2개월간 미사 8번 참례, 이후 2개월 경과'에서 '입회 후 3개월 이내에 미사 4회 이상 참례'로 바뀝니다.",
  "본회는 정자동 주교좌 성당에 교적을 둔 미혼 20세 ~ 만 39세 청년으로서 회비 납부 후 지속적인 활동이 가능한 자를 회원으로 인정한다.\n1. 자격의 제한 : 만일 자격의 제한이 되는 인원 중 청년회 활동을 지속하고 싶은 이는 지도 신부와 면담을 통해 활동 여부를 결정한다.\n2. 신규 청년 입회 절차 : 신규 청년 입단 후 2개월 간 청년 미사 8번 참례 기준 신규 청년으로서 취급한다. 이후 2개월 경과 후 정규 청년으로 취급한다.\n3. 예비신자 입회 청년회의 예비신자가 청년회에 입회할 수 있으며, 입회하고자 하는 경우 예비신자 교육을 받는 상태여야 한다.",
  "본회는 정자동 주교좌 성당에 교적을 둔 만 20세 이상 만 39세 이하의 미혼 청년으로서, 회비를 납부하고 지속적으로 활동할 수 있는 자를 회원으로 인정한다.\n1. 자격의 제한 : 위 자격에 해당하지 않는 자 중 본회 활동을 계속하고자 하는 자는 지도 신부와 면담하여 활동 여부를 결정한다.\n2. 신규 청년 : 입회한 자는 [[신규 청년으로 하며, 입회 후 3개월 이내에 청년 미사에 4회 이상 참례한 때부터 정규 청년으로 한다.]]\n3. 예비신자 입회 : 예비신자 교리를 받고 있는 자는 교적이 없더라도 본회에 입회할 수 있다."
 ],
 [
  22,
  "제22조 (정기회의 및 모임 일정)",
  "회장이 긴급하거나 별도 논의가 필요하다고 판단하면 임시회의를 소집할 수 있다는 조항(5항)이 추가됩니다.",
  "1. 매달 첫째 주 청년 미사 후 임원 회의를 가진다.\n  1.1 첫째 주에 정기 회의가 불가능하다고 회장이 판단한 경우 시간 및 주를 변경할 수 있다.\n2. 본회의 총회는 회계연도가 끝나기 전에 실시하며, 회계 연도가 끝나는 시점에 인수인계를 완료한다.\n3. 정기회의 후 회의록은 임원진이 작성하며, 회장이 승인한 경우 청년회 청년이 작성할 수 있다.\n4. 정기회의의 회의록은 지도 신부, 지도 수녀, 청소년 위원장에게 보고한다.",
  "본회의 회의는 다음과 같이 운영한다.\n1. 정기회의(임원 회의)는 매월 첫째 주 청년 미사 후에 개최한다.\n  1.1 회장이 첫째 주에 개최하기 어렵다고 판단한 경우 일시를 변경할 수 있다.\n2. 총회는 회계연도가 끝나기 전에 개최하며, 회계연도 종료 시까지 인수인계를 완료한다.\n3. 정기회의의 회의록은 임원이 작성하며, 회장이 승인한 경우 임원이 아닌 회원이 작성할 수 있다.\n4. 정기회의의 회의록은 지도 신부, 지도 수녀, 청소년 위원장에게 보고한다.\n[[5. 회장은 긴급하거나 별도의 논의가 필요하다고 판단하는 경우 임시회의를 소집할 수 있다.]]"
 ],
 [
  23,
  "제23조 (수입)",
  "수입을 '본당 지원금 + 자부담금(전년도 이월금·회비·찬조금·기타 수입)'으로 구체적으로 정합니다.",
  "본회의 수입은 본당 지원금으로 운용되며, 청년회의 운용 편의를 위하여 회비를 납부 받을 수 있다.",
  "본회의 수입은 다음과 같다.\n[[1. 본회의 수입은 본당 지원금과 자부담금으로 구성한다.]]\n[[2. 자부담금은 전년도 이월금, 회비, 찬조금 및 기타 수입으로 한다.]]"
 ],
 [
  26,
  "제26조 (지출)",
  "재정 사용 원칙과 자부담금 지출 기준을 새로 정합니다. 임원 회의비는 1인 2만 원·연 200만 원 이내, 청년회 회합은 1인당 식비 1만 원 내외이며, 지도 신부의 배석 또는 허가와 정규 청년 50% 이상 참석이 필요합니다.",
  "본회의의 지출은 다음과 같다.\n1. 정기 회의와 친목 모임 등 각종 행사에 필요한 경비로 한다.\n2. 본회의 자산을 사용함에 있어 영수증 처리를 원칙으로 한다.",
  "본회의 지출은 다음 각 호의 기준에 의거하여 집행한다.\n[[1. 재정의 구분 및 사용 원칙]]\n  [[1.1. 본회의 재정은 본당 지원금과 청년회 자부담금으로 구성되며, 본당 지원금은 연초에 수립한 연간계획서에 따라 각종 행사에 필요한 경비로 사용한다.]]\n  [[1.2. 다만, 임원진 활동비, 청년회 회합 등은 일부 자부담금으로 처리한다.]]\n  [[1.3. 본당 지원금의 연간계획서에 해당 항목이 포함된 경우에는 본당 지원금으로 집행할 수 있다.]]\n[[2. 지출 항목 및 기준(자부담금 기준)]]\n  [[2.1. 임원진 회의비: 정기 및 임시 회의를 포함한 임원진 회의 시 1인당 20,000원 이내로 지출하며, 연간 총사용액은 2,000,000원을 초과할 수 없다.]]\n  [[2.2. 청년회 회합: 청년회 전체 친목 모임 시 1인당 식비 기준 10,000원 내외로 지출할 수 있으며, 이는 반드시 회장의 허락 하에 진행한다.]]\n  [[2.3. 친목 모임 지출 요건: 위 “2.2”항의 경비를 지출하기 위해서는 아래의 요건을 모두 충족해야 한다.]]\n    [[1) 지도 신부의 배석 또는 허가]]\n    [[2) 정규 청년의 50% 이상 참석]]\n[[3. 본회의 자산을 사용함에 있어 영수증 처리를 원칙으로 한다.]]"
 ],
 [
  28,
  "제28조 (재정원칙)",
  "재정의 사적 사용 금지, 이자 수익의 청년회 귀속, 사적 사용 시 전액 반환과 최대 10배 감사헌금 조치가 추가됩니다.",
  "1. 본 회의의 회원이 그 자격을 상실하였을 때 본 재정의 반환청구를 할 수 없다.\n2. 청년회가 주관하는 모든 행사의 결산은 행사 진행 2주 내로 완료한다.",
  "본회의 재정원칙은 다음과 같다.\n1. 회원이 자격을 상실한 경우 이미 납부한 회비 등의 반환을 청구할 수 없다.\n[[2. 본회의 모든 재정은 본회의 운영과 활동을 위한 목적으로만 사용하며, 개인적 또는 사적인 용도로 사용할 수 없다.]]\n[[3. 본회의 재정을 은행에 예치하여 발생한 이자 및 부수 수익은 모두 본회의 재정에 귀속되며, 본회의 자체 활동비로 사용한다.]]\n[[4. 본회의 재정을 개인적 또는 사적인 용도로 사용한 사실이 확인된 경우 해당 금액 전액을 반환하여야 하며, 공동체의 질서 회복을 위하여 지도 신부와 협의 후 사용 금액의 10배에 해당하는 금액을 성당에 감사헌금으로 봉헌하도록 할 수 있다.]]\n5. 본회가 주관하는 모든 행사의 결산은 행사 종료일부터 2주 이내에 완료한다."
 ],
 [
  29,
  "제29조 (재정공개)",
  "재정 공개를 '매달'에서 '반기별(6개월마다)'로 바꾸고, 회원이 요구하면 총무가 수시로 공개할 수 있게 합니다. 결산서는 매 회계연도 종료 시 보고합니다.",
  "1. 회비 및 재정관리 내용은 본 회의 회원에게 매달 공개를 원칙으로 한다.\n2. 재작년도 종료일에 사업 보고서 및 수입 지출 결산서를 본회의 회원에게 보고한다.",
  "본회의 재정은 다음과 같이 공개한다.\n[[1. 회비 및 재정 관리 내용은 반기별 공개를 원칙으로 한다.]]\n  [[1.1 회원의 요구가 있는 경우 총무는 재정을 공개할 수 있다.]]\n[[2. 매 회계연도 종료 시 수입·지출 결산서를 회원에게 보고한다.]]"
 ],
 [
  31,
  "제31조 (청년의 포상)",
  "포상을 '반년 시상'에서 '반기별 포상'으로 정리하고, 포상 대상자와 세부 기준은 임원진 협의로 정합니다.",
  "본회에서 모범으로 활동한 청년에 대하여 반년 시상을 진행한다.",
  "본회의 포상은 다음과 같이 실시한다.\n[[1. 본 회는 모범적으로 활동한 청년에 대하여 반기별로 포상을 실시한다.]]\n[[2. 포상 대상자 선정 및 세부 기준은 임원진 협의를 통해 결정한다.]]"
 ],
 [
  32,
  "제32조 (회칙의 개정)",
  "개정 발의 요건을 '정규 청년 5명 이상 및 임원 3명'에서 '또는'으로 바꾸고, 의결 기준을 '3분의 2 출석·3분의 2 찬성'에서 '2분의 1 출석·2분의 1 찬성'으로 낮춥니다.",
  "1. 본 회칙의 개정은 정규 청년 5인 이상 및 임원 3명 이상의 요청에 의해 임시총회 및 정기총회에서 정규 청년 3분의 2 출석 중 3분의 2 이상의 찬성으로 개정할 수 있다.\n2. 본회의 개정은 본회의 목적에 어긋나지 않아야 한다.",
  "회칙의 개정은 다음과 같이 한다.\n[[1. 회칙의 개정은 정규 청년 5명 이상 또는 임원 3명 이상의 요청으로 발의하며, 정기총회 또는 임시총회에서 심의한다.]]\n[[2. 회칙의 개정은 정규 청년 2분의 1 이상의 출석과 출석자 2분의 1 이상의 찬성으로 의결한다.]]\n[[3. 회칙의 개정은 본회의 목적에 어긋나지 않아야 한다.]]"
 ],
 [
  33,
  "제33조 (총회)",
  "기타 안건 의결 기준을 2분의 1로 낮추고 회장(부재 시 부회장) 참석을 요건으로 하며, 회칙 개정 결과를 메신저로 전 인원에게 공표합니다.",
  "1. 회칙에 규정되지 않는 사항은 임시총회를 소집하여 정규 청년 3분의 2 출석 중 3분의 2 이상의 찬성으로 표결하여 처리한다.\n2. 총회 및 임시총회는 대면 진행을 원칙으로 하며, 부득이하게 비대면으로 진행하여야 할 경우 지도 신부가 반드시 참석하여야 총회로서 유효하다.",
  "총회는 다음과 같이 운영한다.\n[[1. 회칙에 규정되지 않은 사항은 회장(회장 부재 시 부회장)의 참석하에 임시총회를 소집하여, 정규 청년 2분의 1 이상의 출석과 출석자 2분의 1 이상의 찬성으로 의결한다.]]\n[[2. 회칙을 개정한 경우 그 결과를 메신저를 통해 전 회원에게 공표한다.]]\n[[3. 정기총회 및 임시총회는 대면 진행을 원칙으로 하며, 부득이하게 비대면으로 진행하는 경우에는 지도 신부가 반드시 참석하여야 유효하다.]]"
 ]
];
