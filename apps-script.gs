/**
 * 입체도형, 그것이 알고 싶다!! — 학습 기록 수집 · 수업 설정 공유 (Google Apps Script)
 * -----------------------------------------------------------------------------
 * 만드는 탭
 *   1) 활동 데이터 : 전송될 때마다 한 줄씩 그대로 쌓임 (보낸 소활동 결과 + 그 시점까지의 학습 데이터 전체)
 *                    → 같은 학생의 데이터가 여러 줄 있을 수 있음
 *   2) 누적 데이터 : 학번·이름마다 한 줄. 이번에 보낸 데이터와 지금까지의 기록을 비교해 가장 높은 값으로 갱신
 *                    → 학생마다 한 줄만 있음 (이미 만들어 둔 빈 탭을 그대로 사용, 첫 줄 제목은 자동으로 채움)
 *   3) 수업설정    : 관리자가 [수업 관리]에서 저장한 잠금·통제 설정 (학생 화면이 10초마다 읽어 감)
 *   4) 활동 점수    : 학번·이름마다 한 줄. 주제 › 활동 › 소활동 › 세부활동별 점수 (2번째 줄 = 만점)
 *   5) 활동 점수 기준 : 점수 칸마다 활동 종류·완료 기준·배점 (웹앱이 보내 준 기준표로 자동 작성)
 *   6) 활동 구성    : 주제 › 활동 › 소활동 › 세부활동별 완료 기준과 배점 표 + 주제별 배점 합계
 *                    (관리자 [📊 모니터링] → [📋 활동 구성표 만들기] 또는 [🧮 활동 점수 다시 계산]을 누르면 작성)
 *   ※ 모니터링     : 학생 화면이 '지금 하는 소활동·마친 소활동'을 알려 오면 캐시에 잠시(6시간) 보관 → 관리자 [📊 모니터링]이 읽어 감
 *                    (시트에 쓰지 않으므로 빠르고, 기록 저장과 서로 기다리지 않음)
 *
 * 설치
 *   1) 구글 시트를 새로 만들고 [확장 프로그램] → [Apps Script]를 연다.
 *   2) 기본 코드를 모두 지우고 이 파일 내용을 붙여 넣은 뒤 저장한다.
 *   3) [배포] → [새 배포] → 유형: 웹 앱
 *        · 실행 사용자: 나
 *        · 액세스 권한이 있는 사용자: 모든 사용자
 *      → [배포]를 누르고 권한을 허용한 뒤, 나오는 '웹 앱 URL'(…/exec)을 복사한다.
 *   4) 웹앱 HTML 파일 맨 위 근처의 SHEET_WEBAPP_URL = '' 안에 그 주소를 붙여 넣는다.
 *      (또는 관리자 → 수업 관리 → 구글 시트 연결 칸에 넣는다: 그 기기에만 적용)
 *   ※ 코드를 고친 뒤에는 [배포] → [배포 관리] → 수정 → 버전 '새 버전'으로 다시 배포해야 반영된다.
 */

const SHEET_LOG = '활동 데이터';    // 전송된 데이터를 그대로 한 줄씩 쌓음 (같은 학생이 여러 줄)
const SHEET_BEST = '누적 데이터';   // 학번·이름마다 한 줄, 가장 높은 값으로 계속 갱신
const SHEET_CFG = '수업설정';
const SHEET_SCORE = '활동 점수';
const SHEET_SCORE_STD = '활동 점수 기준';
const SHEET_COMP = '활동 구성';
const SCORE_FIXED = ['학번', '이름', '갱신 시각', '총점', '만점', '달성률(%)'];
const DEFAULT_PW_HASH = '445cf6a7';   // 처음 비밀번호(1234)를 웹앱과 같은 방식으로 바꾼 값 — 아래 hashPw_('1234')와 같다

// 용어: 주제(0. 복습 ~ 8. 구의 부피와 겉넓이) › 활동(학습 목표, 생각 열기 …) › 소활동 › 세부활동
const LOG_HEAD = ['전송 시각', '학번', '이름', '주제', '활동', '소활동', '정답 수', '문항 수', '점수(%)', '최고 점수(%)',
  '학습 시간(초)', '방문 횟수', '완료한 주제 수', '확인 퀴즈 점수', '전송 시점까지의 학습 데이터 전체(JSON)', '입력한 값'];
const BEST_FIXED = ['학번', '이름', '최근 전송 시각', '전송 횟수', '완료한 주제 수', '완료한 소활동 수', '평균 점수(%)', '누적 학습 시간(초)', '이어하기 데이터(JSON)'];

/* ---------- 받기 ---------- */
function doPost(e) {
  // 모니터링 신호는 자주 오므로 잠금 없이 캐시에만 둔다
  let early = null; try { early = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  if (early && early.action === 'status') { try { return json_(saveStatus_(early)); } catch (err) { return json_({ ok: false, error: String(err) }); } }
  if (early && early.action === 'monitorClear') { try { return json_(clearMonitor_(early)); } catch (err) { return json_({ ok: false, error: String(err) }); } }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.action === 'record') return json_(saveRecord_(data));
    if (data.action === 'config') return json_(saveConfig_(data));
    if (data.action === 'scores') return json_(saveScores_(data));
    if (data.action === 'composition') return json_(saveComposition_(data));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* 학생 화면이 수업 설정을 읽어 갈 때: …/exec?action=config */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'config') return json_({ ok: true, config: readConfig_() });
  if (action === 'restore') return json_(restore_(e.parameter.sid, e.parameter.name));
  if (action === 'monitor') return json_(readMonitor_(e.parameter.auth));
  if (action === 'allProgress') return json_(allProgress_(e.parameter.auth));
  return json_({ ok: true, message: '입체도형 학습 기록 수집 중' });
}

/* ---------- 1) 활동 데이터 ---------- */
function saveRecord_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sid = String(d.sid || '').trim();
  if (!sid) return { ok: false, error: 'no sid' };
  const it = d.item || {};
  const quiz = quizText_(d.quiz);
  const log = sheet_(ss, SHEET_LOG, LOG_HEAD); ensureHead_(log, LOG_HEAD);
  log.appendRow([new Date(), "'" + sid, d.name || '', it.lesson || '', it.step || '', it.sub || '',
    num_(it.correct), num_(it.total), num_(it.score), num_(it.best), num_(it.time), num_(it.visits),
    num_(d.completedLessons), quiz, JSON.stringify({ results: d.snapshot || {}, quiz: d.quiz || {} }),
    (it.inputs || []).filter(String).join(' | ')]);
  updateBest_(ss, sid, d);
  return { ok: true };
}

/* ---------- 2) 누적 데이터: 학번·이름마다 한 줄, 소활동마다 열 묶음 ----------
 * 소활동 하나에 열 4개: "주제 › 활동 › 소활동 · 최고 점수(%)" / "· 최근 정답/문항" / "· 최근 입력값" / "· 학습 시간(초)"
 * 새 소활동이 오면 열을 추가하고, 모든 열을 주제 → 활동 → 소활동 번호 순서로 정리한 뒤 학생 줄을 그 열에 맞춰 다시 쓴다.
 */
const METRICS = ['최고 점수(%)', '최근 정답/문항', '최근 입력값', '학습 시간(초)'];
function updateBest_(ss, sid, d) {
  const sh = sheet_(ss, SHEET_BEST, BEST_FIXED);
  const table = readTable_(sh);
  const nm = String(d.name || '').trim();
  let obj = table.rows.find(function (o) { return o['학번'] === sid && String(o['이름']).trim() === nm; });
  if (!obj) { obj = { '학번': sid, '이름': nm }; table.rows.push(obj); }
  mergeStudent_(obj, d.snapshot || {}, null);
  obj['최근 전송 시각'] = new Date();
  obj['전송 횟수'] = (Number(obj['전송 횟수']) || 0) + 1;
  obj['완료한 주제 수'] = Math.max(Number(obj['완료한 주제 수']) || 0, Number(d.completedLessons) || 0);
  if (d.progress) {
    let old = null; try { old = JSON.parse(obj['이어하기 데이터(JSON)'] || 'null'); } catch (err) {}
    obj['이어하기 데이터(JSON)'] = JSON.stringify(mergeProgress_(old, d.progress));
  }
  writeTable_(sh, table);
}
/** 학생 한 명의 소활동 데이터 합치기 (점수·시간은 큰 값, 정답/문항·입력값은 최근 값) */
function mergeStudent_(obj, snap, names) {
  Object.keys(snap).forEach(function (k) {
    const r = snap[k] || {};
    const base = (names && names[k]) || colBase_(r, k);
    const sc = r.best != null ? Number(r.best) : (r.score != null ? Number(r.score) : 100);   // 입력 문항이 없는 소활동은 완료 = 100
    const c1 = base + ' · ' + METRICS[0], c2 = base + ' · ' + METRICS[1], c3 = base + ' · ' + METRICS[2], c4 = base + ' · ' + METRICS[3];
    obj[c1] = (obj[c1] === '' || obj[c1] == null) ? sc : Math.max(Number(obj[c1]) || 0, sc);
    obj[c2] = r.total ? (r.correct || 0) + '/' + r.total : (obj[c2] || '완료');
    const inp = (r.inputs || []).filter(function (x) { return x !== '' && x != null; }).join(' | ');
    obj[c3] = inp || obj[c3] || '';
    obj[c4] = Math.max(Number(obj[c4]) || 0, Number(r.time) || 0);
  });
  // 요약 열
  const scoreCols = Object.keys(obj).filter(function (c) { return c.slice(-METRICS[0].length) === METRICS[0] && obj[c] !== '' && obj[c] != null; });
  const timeCols = Object.keys(obj).filter(function (c) { return c.slice(-METRICS[3].length) === METRICS[3]; });
  obj['완료한 소활동 수'] = scoreCols.length;
  obj['평균 점수(%)'] = scoreCols.length ? Math.round(scoreCols.reduce(function (s, c) { return s + Number(obj[c]); }, 0) / scoreCols.length) : '';
  obj['누적 학습 시간(초)'] = timeCols.reduce(function (s, c) { return s + (Number(obj[c]) || 0); }, 0);
}
function colBase_(r, k) { return [r.lesson, r.step, r.sub].filter(function (x) { return x; }).join(' › ') || k; }
/** 열 정렬: 주제 번호 → 활동 번호 → 소활동 번호 → 항목 순서 (번호가 없는 열은 뒤로) */
function colKey_(c) {
  const mi = METRICS.findIndex(function (m) { return c.slice(-m.length) === m; });
  const base = mi >= 0 ? c.slice(0, c.length - METRICS[mi].length - 3) : c;
  const L = base.match(/^(\d+)\./), S = base.match(/ › (\d+)\. /), K = base.match(/ › (\d+)\) /);
  const pad = function (m) { return ('000' + (m ? m[1] : '999')).slice(-3); };
  return (L ? '0' : '1') + pad(L) + pad(S) + (K ? pad(K) : '000') + '|' + (L ? '' : base) + '|' + (mi < 0 ? 9 : mi);
}
function readTable_(sh) {
  const lr = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), BEST_FIXED.length);
  const vals = lr ? sh.getRange(1, 1, lr, lc).getValues() : [BEST_FIXED];
  const head = vals[0].map(String);
  const rows = vals.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    const o = {}; head.forEach(function (h, i) { if (h) o[h] = r[i]; }); o['학번'] = String(o['학번']).replace(/^'/, ''); return o;
  });
  return { head: head, rows: rows };
}
function writeTable_(sh, table) {
  const cols = {};
  table.head.forEach(function (h) { if (h && BEST_FIXED.indexOf(h) < 0) cols[h] = 1; });
  table.rows.forEach(function (o) { Object.keys(o).forEach(function (h) { if (BEST_FIXED.indexOf(h) < 0) cols[h] = 1; }); });
  const head = BEST_FIXED.concat(Object.keys(cols).sort(function (a, b) { const x = colKey_(a), y = colKey_(b); return x < y ? -1 : x > y ? 1 : 0; }));
  const out = [head].concat(table.rows.map(function (o) {
    return head.map(function (h) { const v = o[h]; return h === '학번' ? "'" + v : (v == null ? '' : v); });
  }));
  sh.clearContents();
  sh.getRange(1, 1, out.length, head.length).setValues(out);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8EAF6');
  sh.setFrozenRows(1);
}

/** 누적 데이터 다시 만들기: '활동 데이터'에 쌓인 모든 기록으로 학생별 한 줄을 새로 정리
 *  (예전 이름의 열이 섞였거나 열이 어긋났을 때 Apps Script 편집기에서 이 함수를 선택해 [실행]) */
function rebuildBest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SHEET_LOG); if (!log || log.getLastRow() < 2) return;
  const vals = log.getRange(1, 1, log.getLastRow(), log.getLastColumn()).getValues();
  const head = vals[0].map(String);
  const iT = head.indexOf('전송 시각'), iS = head.indexOf('학번'), iN = head.indexOf('이름'), iC = head.indexOf('완료한 주제 수'), iJ = head.indexOf('전송 시점까지의 학습 데이터 전체(JSON)');
  const rows = vals.slice(1);
  // 같은 소활동은 가장 최근에 보낸 이름으로 통일
  const names = {};
  rows.forEach(function (r) { let j = null; try { j = JSON.parse(r[iJ] || 'null'); } catch (e) {} const res = (j && j.results) || {};
    Object.keys(res).forEach(function (k) { names[k] = colBase_(res[k], k); }); });
  const by = {};
  rows.forEach(function (r) {
    const sid = String(r[iS]).replace(/^'/, '').trim(), nm = String(r[iN]).trim(); if (!sid) return;
    const key = sid + '|' + nm, o = by[key] || (by[key] = { '학번': sid, '이름': nm, '전송 횟수': 0 });
    let j = null; try { j = JSON.parse(r[iJ] || 'null'); } catch (e) {}
    mergeStudent_(o, (j && j.results) || {}, names);
    o['전송 횟수']++; o['최근 전송 시각'] = r[iT];
    o['완료한 주제 수'] = Math.max(Number(o['완료한 주제 수']) || 0, Number(r[iC]) || 0);
  });
  // 이어하기 데이터는 지금 시트의 값을 그대로 둔다
  const sh = sheet_(ss, SHEET_BEST, BEST_FIXED), cur = readTable_(sh);
  cur.rows.forEach(function (o) { const k = o['학번'] + '|' + String(o['이름']).trim(); if (by[k]) by[k]['이어하기 데이터(JSON)'] = o['이어하기 데이터(JSON)']; });
  writeTable_(sh, { head: BEST_FIXED, rows: Object.keys(by).map(function (k) { return by[k]; }) });
}

/* ---------- 이어하기: 같은 학번·이름의 진행 상태 돌려주기 ---------- */
function restore_(sid, name) {
  sid = String(sid || '').trim(); name = String(name || '').trim();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BEST);
  if (!sh || sh.getLastRow() < 2 || !sid) return { ok: true, progress: null };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
  const iS = head.indexOf('학번'), iN = head.indexOf('이름'), iP = head.indexOf('이어하기 데이터(JSON)');
  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][iS]).replace(/^'/, '') === sid && String(rows[r][iN]).trim() === name) {
      try { return { ok: true, progress: JSON.parse(rows[r][iP] || 'null') }; } catch (err) { return { ok: true, progress: null }; }
    }
  }
  return { ok: true, progress: null };
}
function mergeProgress_(a, b) {
  if (!a) return b; if (!b) return a;
  const out = { completed: {}, scores: {}, steps: {}, seen: {}, subtabs: {}, results: {}, detail: {}, act: {} };
  [a.completed, b.completed].forEach(function (o) { Object.keys(o || {}).forEach(function (k) { if (o[k]) out.completed[k] = true; }); });
  [a.scores, b.scores].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const v = o[k] || {}, c = out.scores[k];
    out.scores[k] = !c ? v : { first: c.first != null ? c.first : v.first, best: Math.max(c.best || 0, v.best || 0), total: Math.max(c.total || 0, v.total || 0), rounds: Math.max(c.rounds || 0, v.rounds || 0) };
  }); });
  [a.steps, b.steps].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const v = o[k] || {}, c = out.steps[k];
    out.steps[k] = !c ? { cur: v.cur || 0, reached: v.reached || 0 } : { cur: v.cur || 0, reached: Math.max(c.reached || 0, v.reached || 0) };
  }); });
  ['seen', 'subtabs', 'detail', 'act'].forEach(function (f) { [a[f], b[f]].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const s = {}; (out[f][k] || []).concat(o[k] || []).forEach(function (x) { s[x] = 1; }); out[f][k] = Object.keys(s);
  }); }); });
  [a.results, b.results].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const v = o[k] || {}, c = out.results[k];
    out.results[k] = !c ? v : Object.assign({}, v, { best: Math.max(c.best || 0, v.best || 0), time: Math.max(c.time || 0, v.time || 0), visits: Math.max(c.visits || 0, v.visits || 0) });
  }); });
  return out;
}

/* ---------- 3) 수업설정 ---------- */
function saveConfig_(d) {
  const cur = readConfig_();
  const known = (cur && cur.pwHash) || DEFAULT_PW_HASH;
  if (String(d.auth || '') !== known) return { ok: false, error: '관리자 비밀번호가 맞지 않아요.' };
  const sh = sheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_CFG, ['설정(JSON)', '저장 시각']);
  sh.getRange(2, 1, 1, 2).setValues([[JSON.stringify(d.config || {}), new Date()]]);
  try { CacheService.getScriptCache().put('cfg', JSON.stringify(d.config || {}), 30); } catch (err) {}
  return { ok: true };
}
/* 학생 화면이 10초마다 읽어 가므로, 시트를 매번 읽지 않게 30초 동안 캐시에 둔다 (저장하면 바로 바뀜) */
function readConfig_() {
  try { const c = CacheService.getScriptCache().get('cfg'); if (c) return JSON.parse(c); } catch (err) {}
  const v = readConfigSheet_();
  try { if (v) CacheService.getScriptCache().put('cfg', JSON.stringify(v), 30); } catch (err) {}
  return v;
}
function readConfigSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CFG);
  if (!sh || sh.getLastRow() < 2) return null;
  try { return JSON.parse(sh.getRange(2, 1).getValue()); } catch (err) { return null; }
}

/* ---------- 4) 활동 점수 ---------- */
/** 웹앱이 계산해 보낸 점수 쓰기: { cols, max, maxTotal, crit, items:[{ sid, name, total, got }] } — 칸 점수는 내려가지 않음(가장 높은 값 유지) */
function saveScores_(d) {
  const cols = d.cols || [], max = d.max || [];
  if (!cols.length || !(d.items || []).length) return { ok: false, error: 'no scores' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const head = SCORE_FIXED.concat(cols);
  const sh = sheet_(ss, SHEET_SCORE, head);
  if ((d.items || []).length <= 3 && fastScores_(sh, head, d)) { writeScoreStd_(ss, cols, max, d.crit || []); return { ok: true, fast: true }; }
  // 지금 표 읽기 (열 이름으로 맞춰 옮겨 담기 → 활동 구성이 바뀌어도 예전 값이 섞이지 않음)
  const lastR = sh.getLastRow(), lastC = Math.max(sh.getLastColumn(), 1);
  const old = lastR >= 1 ? sh.getRange(1, 1, lastR, lastC).getValues() : [[]];
  const oh = (old[0] || []).map(String), pos = {}; oh.forEach(function (h, i) { pos[h] = i; });
  const idx = {}, rows = [];
  for (let r = 2; r < old.length; r++) {      // 1번째 줄 = 제목, 2번째 줄 = 만점
    const o = old[r]; const sid = String(o[pos['학번']] || '').replace(/^'/, ''), nm = String(o[pos['이름']] || '').trim();
    if (!sid && !nm) continue;
    const row = head.map(function (h) { return pos[h] != null ? o[pos[h]] : ''; });
    idx[sid + '|' + nm] = rows.length; rows.push(row);
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const n0 = SCORE_FIXED.length;
  (d.items || []).forEach(function (it) {
    const sid = String(it.sid || '').trim(), nm = String(it.name || '').trim(); if (!sid) return;
    let i = idx[sid + '|' + nm]; if (i == null) { i = rows.length; idx[sid + '|' + nm] = i; rows.push(head.map(function () { return ''; })); }
    const row = rows[i]; row[0] = "'" + sid; row[1] = nm; row[2] = now;
    (it.got || []).forEach(function (v, j) { const c = n0 + j; if (((d.crit || [])[j] || {}).level === '세부활동') row[c] = Math.max(Number(row[c]) || 0, Number(v) || 0); });
  });
  // 총점·주제·활동·소활동 합계는 세부활동 칸에서 다시 더함 (crit[j].range = 세부활동 칸 범위)
  const crit = d.crit || [], leafCols = [];
  cols.forEach(function (c, j) { if ((crit[j] || {}).level === '세부활동') leafCols.push(n0 + j); });
  const maxTotal = Math.round(leafCols.reduce(function (a, c) { return a + (Number(max[c - n0]) || 0); }, 0) * 100) / 100;
  rows.forEach(function (row) {
    const tot = Math.round(leafCols.reduce(function (a, c) { return a + (Number(row[c]) || 0); }, 0) * 100) / 100;
    row[3] = tot; row[4] = maxTotal; row[5] = maxTotal ? Math.round(tot / maxTotal * 1000) / 10 : '';
    cols.forEach(function (c, j) { const rg = (crit[j] || {}).range; if (!rg) return;
      let a = 0; for (let k = rg[0]; k < rg[1]; k++) a += Number(row[leafCols[k]]) || 0; row[n0 + j] = Math.round(a * 100) / 100; });
  });
  rows.sort(function (a, b) { return String(a[0]).replace(/^'/, '').localeCompare(String(b[0]).replace(/^'/, ''), 'ko', { numeric: true }) || String(a[1]).localeCompare(String(b[1]), 'ko'); });
  const maxRow = ['만점', '', '', maxTotal, maxTotal, 100].concat(max);
  sh.clearContents();
  const all = [head, maxRow].concat(rows);
  sh.getRange(1, 1, all.length, head.length).setValues(all);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#FFF3E0').setWrap(true);
  sh.getRange(2, 1, 1, head.length).setFontWeight('bold').setBackground('#F5F5F5');
  sh.setFrozenRows(2); sh.setFrozenColumns(2);
  if (lastC > head.length) sh.getRange(1, head.length + 1, Math.max(lastR, all.length), lastC - head.length).clearContent();
  writeScoreStd_(ss, cols, max, d.crit || []);
  return { ok: true, students: rows.length };
}
/** 학생 한두 명 갱신: 제목 줄이 같으면 그 학생 줄만 고쳐 씀 (빠름) */
function fastScores_(sh, head, d) {
  const lastR = sh.getLastRow(); if (lastR < 2 || sh.getLastColumn() !== head.length) return false;
  const cur = sh.getRange(1, 1, 1, head.length).getValues()[0].map(String);
  for (let i = 0; i < head.length; i++) if (cur[i] !== head[i]) return false;
  const n0 = SCORE_FIXED.length, cols = d.cols, max = d.max || [], crit = d.crit || [];
  const leafCols = []; cols.forEach(function (c, j) { if ((crit[j] || {}).level === '세부활동') leafCols.push(n0 + j); });
  const maxTotal = Math.round(leafCols.reduce(function (a, c) { return a + (Number(max[c - n0]) || 0); }, 0) * 100) / 100;
  const ids = lastR >= 3 ? sh.getRange(3, 1, lastR - 2, 2).getValues() : [];
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  let added = false;
  d.items.forEach(function (it) {
    const sid = String(it.sid || '').trim(), nm = String(it.name || '').trim(); if (!sid) return;
    let r = -1; for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).replace(/^'/, '') === sid && String(ids[i][1]).trim() === nm) { r = i + 3; break; }
    let row;
    if (r < 0) { r = sh.getLastRow() + 1; row = head.map(function () { return ''; }); added = true; ids.push(["'" + sid, nm]); }
    else row = sh.getRange(r, 1, 1, head.length).getValues()[0];
    row[0] = "'" + sid; row[1] = nm; row[2] = now;
    (it.got || []).forEach(function (v, j) { const c = n0 + j; if ((crit[j] || {}).level === '세부활동') row[c] = Math.max(Number(row[c]) || 0, Number(v) || 0); });
    const tot = Math.round(leafCols.reduce(function (a, c) { return a + (Number(row[c]) || 0); }, 0) * 100) / 100;
    row[3] = tot; row[4] = maxTotal; row[5] = maxTotal ? Math.round(tot / maxTotal * 1000) / 10 : '';
    cols.forEach(function (c, j) { const rg = (crit[j] || {}).range; if (!rg) return;
      let a = 0; for (let k = rg[0]; k < rg[1]; k++) a += Number(row[leafCols[k]]) || 0; row[n0 + j] = Math.round(a * 100) / 100; });
    sh.getRange(r, 1, 1, head.length).setValues([row]);
  });
  sh.getRange(2, 1, 1, head.length).setValues([['만점', '', '', maxTotal, maxTotal, 100].concat(max)]);
  if (added && sh.getLastRow() > 3) sh.getRange(3, 1, sh.getLastRow() - 2, head.length).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  return true;
}
/** 활동 점수 기준 탭 (기준이 바뀌었을 때만 다시 씀) */
function writeScoreStd_(ss, cols, max, crit) {
  const sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([cols, max])));
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('scoreStdSig') === sig && ss.getSheetByName(SHEET_SCORE_STD)) return;
  const head = ['열 이름', '단계', '주제', '활동', '소활동', '세부활동', '활동 종류', '완료 기준', '배점(만점)'];
  const sh = sheet_(ss, SHEET_SCORE_STD, head);
  const rows = cols.map(function (c, i) { const k = crit[i] || {};
    return [c, k.level || '', k.lesson || '', k.step || '', k.sub || '', k.det || '', k.kind || (k.level ? k.level + ' 합계' : ''), k.crit || (k.level && k.level !== '세부활동' ? '아래 세부활동 점수의 합' : ''), max[i]]; });
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#FFF3E0');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);
  props.setProperty('scoreStdSig', sig);
}
/** 활동 구성 탭: { auth, head, rows, kinds, summary:[[주제, 종류별…, 합계]] } */
function saveComposition_(d) {
  if (!monAuth_(d.auth)) return { ok: false, error: '관리자 비밀번호가 맞지 않아요.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_COMP); if (!sh) sh = ss.insertSheet(SHEET_COMP);
  sh.clear();
  const head = d.head || [], rows = d.rows || [], W = head.length;
  sh.getRange(1, 1, 1, W).setValues([head]).setFontWeight('bold').setBackground('#FFE0B2').setHorizontalAlignment('center');
  if (rows.length) sh.getRange(2, 1, rows.length, W).setValues(rows).setVerticalAlignment('middle');
  // 주제마다 번갈아 연한 색
  const bg = ['#FFFFFF', '#F3F6FB']; let k = 0, start = 2;
  for (let i = 1; i <= rows.length; i++) {
    if (i === rows.length || rows[i][1] !== rows[i - 1][1]) { sh.getRange(start, 1, i + 2 - start, W).setBackground(bg[k % 2]); k++; start = i + 2; }
  }
  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  if (rows.length) sh.getRange(1, 1, rows.length + 1, W).createFilter();
  // 주제별 배점 합계
  const sum = d.summary || [];
  if (sum.length) {
    const r0 = rows.length + 4, sw = sum[0].length;
    sh.getRange(r0 - 1, 1).setValue('▶ 주제별 배점 합계 (활동 종류별)').setFontWeight('bold');
    sh.getRange(r0, 1, sum.length, sw).setValues(sum);
    sh.getRange(r0, 1, 1, sw).setFontWeight('bold').setBackground('#FFE0B2');
    sh.getRange(r0 + sum.length - 1, 1, 1, sw).setFontWeight('bold').setBackground('#FFF3E0');
  }
  const widths = d.widths || [];
  widths.forEach(function (w, i) { if (w) sh.setColumnWidth(i + 1, w); });
  sh.getRange(2, 1, Math.max(rows.length, 1), W).setWrap(true);
  return { ok: true, rows: rows.length };
}
/** 관리자 '활동 점수 다시 계산'용: 학생마다 이어하기 데이터 */
function allProgress_(auth) {
  if (!monAuth_(auth)) return { ok: false, error: '관리자 비밀번호가 맞지 않아요. 다시 로그인해 주세요.' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BEST);
  if (!sh || sh.getLastRow() < 2) return { ok: true, students: [] };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
  const iS = head.indexOf('학번'), iN = head.indexOf('이름'), iP = head.indexOf('이어하기 데이터(JSON)');
  const out = [];
  rows.forEach(function (r) { const sid = String(r[iS]).replace(/^'/, '').trim(); if (!sid) return;
    let pg = null; try { pg = JSON.parse(r[iP] || 'null'); } catch (err) {}
    out.push({ sid: sid, name: String(r[iN]).trim(), progress: pg }); });
  return { ok: true, students: out };
}

/* ---------- 모니터링 ---------- */
const MON_TTL = 21600;   // 6시간 (캐시에 둘 수 있는 가장 긴 시간)
function monKey_(sid, name) { return 'mon:' + String(sid).trim() + '|' + String(name).trim(); }
function monIndex_(cache) { try { return JSON.parse(cache.get('mon:index') || '[]'); } catch (err) { return []; } }
function saveStatus_(d) {
  if (!d.sid || !d.name) return { ok: false, error: 'no id' };
  const cache = CacheService.getScriptCache(), key = monKey_(d.sid, d.name);
  const rec = { sid: String(d.sid).trim(), name: String(d.name).trim(), cur: String(d.cur || ''), done: (d.done || []).slice(0, 400), ts: Date.now() };
  cache.put(key, JSON.stringify(rec), MON_TTL);
  let idx = monIndex_(cache);
  if (idx.indexOf(key) < 0) {   // 처음 알려 온 학생만 목록에 더함 (동시에 들어와도 빠지지 않게 잠깐 잠금)
    const lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) { try { idx = monIndex_(cache); if (idx.indexOf(key) < 0) { idx.push(key); cache.put('mon:index', JSON.stringify(idx), MON_TTL); } } finally { lock.releaseLock(); } }
  } else cache.put('mon:index', JSON.stringify(idx), MON_TTL);   // 목록 보관 시간 연장
  return { ok: true };
}
function monAuth_(auth) { const cur = readConfig_(); return String(auth || '') === ((cur && cur.pwHash) || DEFAULT_PW_HASH); }
function readMonitor_(auth) {
  if (!monAuth_(auth)) return { ok: false, error: '관리자 비밀번호가 맞지 않아요. 다시 로그인해 주세요.' };
  const cache = CacheService.getScriptCache(), idx = monIndex_(cache), out = [];
  for (let i = 0; i < idx.length; i += 100) {
    const got = cache.getAll(idx.slice(i, i + 100));
    Object.keys(got).forEach(function (k) { try { out.push(JSON.parse(got[k])); } catch (err) {} });
  }
  return { ok: true, now: Date.now(), students: out };
}
function clearMonitor_(d) {
  if (!monAuth_(d.auth)) return { ok: false, error: '관리자 비밀번호가 맞지 않아요.' };
  const cache = CacheService.getScriptCache(), idx = monIndex_(cache);
  for (let i = 0; i < idx.length; i += 100) cache.removeAll(idx.slice(i, i + 100));
  cache.remove('mon:index');
  return { ok: true };
}

/* ---------- 도우미 ---------- */
function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#FFF3E0');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
/** 첫 줄 제목을 지금 코드의 이름으로 맞추기 (예전 시트를 그대로 써도 열 위치가 같도록) */
function ensureHead_(sh, head) {
  const cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
  if (head.some(function (h, i) { return String(cur[i]) !== h; })) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
}
function num_(v) { return (v === null || v === undefined || v === '') ? '' : Number(v); }
function quizText_(q) {
  if (!q) return '';
  return Object.keys(q).map(function (k) { const s = q[k] || {}; return s.total ? (k + ': ' + (s.best || 0) + '/' + s.total) : ''; }).filter(String).join(', ');
}
function lessonName_(snap, lid) {
  for (const k in snap) { if (k.indexOf(lid) === 0 && snap[k].lesson) return snap[k].lesson; }
  return lid;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** 웹앱과 같은 비밀번호 변환 (관리자 비밀번호를 시트에서 직접 초기화할 때 사용) */
function hashPw_(s) {
  let h = 0x811c9dc5;
  for (const ch of 'ipche#' + String(s)) { h ^= ch.codePointAt(0); h = Math.imul(h, 0x01000193) >>> 0; }
  return h.toString(16).padStart(8, '0');
}
/** 관리자 비밀번호를 1234로 되돌리기: Apps Script 편집기에서 이 함수를 선택해 [실행] */
function resetAdminPassword() {
  const cfg = readConfig_() || {};
  cfg.pwHash = hashPw_('1234'); cfg.updatedAt = Date.now();
  const sh = sheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_CFG, ['설정(JSON)', '저장 시각']);
  sh.getRange(2, 1, 1, 2).setValues([[JSON.stringify(cfg), new Date()]]);
  try { CacheService.getScriptCache().remove('cfg'); } catch (err) {}
}
