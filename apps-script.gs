/**
 * 입체도형, 그것이 알고 싶다!! — 학습 기록 수집 · 수업 설정 공유 (Google Apps Script)
 * -----------------------------------------------------------------------------
 * 만드는 탭
 *   1) 활동 데이터 : 전송될 때마다 한 줄씩 그대로 쌓임 (보낸 소활동 결과 + 그 시점까지의 학습 데이터 전체)
 *                    → 같은 학생의 데이터가 여러 줄 있을 수 있음
 *   2) 누적 데이터 : 학번·이름마다 한 줄. 이번에 보낸 데이터와 지금까지의 기록을 비교해 가장 높은 값으로 갱신
 *                    → 학생마다 한 줄만 있음 (이미 만들어 둔 빈 탭을 그대로 사용, 첫 줄 제목은 자동으로 채움)
 *   3) 수업설정    : 관리자가 [수업 관리]에서 저장한 잠금·통제 설정 (학생 화면이 20초마다 읽어 감)
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
const DEFAULT_PW_HASH = '445cf6a7';   // 처음 비밀번호(1234)를 웹앱과 같은 방식으로 바꾼 값 — 아래 hashPw_('1234')와 같다

// 용어: 주제(0. 복습 ~ 8. 구의 부피와 겉넓이) › 활동(학습 목표, 생각 열기 …) › 소활동 › 세부활동
const LOG_HEAD = ['전송 시각', '학번', '이름', '주제', '활동', '소활동', '정답 수', '문항 수', '점수(%)', '최고 점수(%)',
  '학습 시간(초)', '방문 횟수', '완료한 주제 수', '확인 퀴즈 점수', '전송 시점까지의 학습 데이터 전체(JSON)', '입력한 값'];
const BEST_FIXED = ['학번', '이름', '최근 전송 시각', '전송 횟수', '완료한 주제 수', '완료한 소활동 수', '평균 점수(%)', '누적 학습 시간(초)', '이어하기 데이터(JSON)'];

/* ---------- 받기 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.action === 'record') return json_(saveRecord_(data));
    if (data.action === 'config') return json_(saveConfig_(data));
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

/* ---------- 2) 누적 데이터: 학번·이름 기준 한 줄, 가장 높은 값으로 갱신 ---------- */
function updateBest_(ss, sid, d) {
  const sh = sheet_(ss, SHEET_BEST, BEST_FIXED); ensureHead_(sh, BEST_FIXED);
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  // 보낸 학습 데이터 전체(snapshot)를 소활동별 열로 펼친다: 열 이름 = "차시 › 단계 › 소활동"
  const cols = {};
  const snap = d.snapshot || {};
  Object.keys(snap).forEach(function (k) {
    const r = snap[k] || {};
    const name = [r.lesson, r.step, r.sub].filter(String).join(' › ') || k;
    const v = r.best != null ? r.best : (r.score != null ? r.score : 100);   // 입력 문항이 없는 소활동은 완료 = 100
    cols[name] = Math.max(cols[name] == null ? -1 : cols[name], Number(v) || 0);
  });
  // 새 열 추가
  const add = Object.keys(cols).filter(function (c) { return head.indexOf(c) < 0; }).sort();
  if (add.length) {
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add]).setFontWeight('bold').setBackground('#E8EAF6');
    head = head.concat(add);
  }

  // 학번·이름이 같은 행 찾기 (없으면 새로)
  const last = sh.getLastRow(), nm = String(d.name || '').trim();
  let row = -1;
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 2).getValues().map(function (r) { return String(r[0]).replace(/^'/, '') + '|' + String(r[1]).trim(); });
    const i = ids.indexOf(sid + '|' + nm);
    if (i >= 0) row = i + 2;
  }
  let vals = row > 0 ? sh.getRange(row, 1, 1, head.length).getValues()[0] : new Array(head.length).fill('');
  while (vals.length < head.length) vals.push('');

  const at = (name) => head.indexOf(name);
  // 점수 열: 기존 값과 비교해 더 큰 값
  Object.keys(cols).forEach(function (c) {
    const j = at(c), old = vals[j];
    vals[j] = (old === '' || old == null) ? cols[c] : Math.max(Number(old) || 0, cols[c]);
  });
  // 고정 열
  const scoreCols = head.slice(BEST_FIXED.length).map(function (_, k) { return vals[BEST_FIXED.length + k]; }).filter(function (v) { return v !== '' && v != null; });
  const totalTime = Object.keys(snap).reduce(function (s, k) { return s + (Number((snap[k] || {}).time) || 0); }, 0);
  vals[at('학번')] = "'" + sid;
  vals[at('이름')] = d.name || vals[at('이름')];
  vals[at('최근 전송 시각')] = new Date();
  vals[at('전송 횟수')] = (Number(vals[at('전송 횟수')]) || 0) + 1;
  vals[at('완료한 주제 수')] = Math.max(Number(vals[at('완료한 주제 수')]) || 0, Number(d.completedLessons) || 0);
  vals[at('완료한 소활동 수')] = Math.max(Number(vals[at('완료한 소활동 수')]) || 0, Object.keys(snap).length);
  vals[at('평균 점수(%)')] = scoreCols.length ? Math.round(scoreCols.reduce(function (s, v) { return s + Number(v); }, 0) / scoreCols.length) : '';
  vals[at('누적 학습 시간(초)')] = Math.max(Number(vals[at('누적 학습 시간(초)')]) || 0, totalTime);
  // 이어하기: 지금까지의 진행 상태와 이번 진행 상태를 합쳐(완료는 유지, 진도는 더 나간 쪽) 저장
  if (at('이어하기 데이터(JSON)') >= 0 && d.progress) {
    let old = null; try { old = JSON.parse(vals[at('이어하기 데이터(JSON)')] || 'null'); } catch (err) {}
    vals[at('이어하기 데이터(JSON)')] = JSON.stringify(mergeProgress_(old, d.progress));
  }

  if (row > 0) sh.getRange(row, 1, 1, head.length).setValues([vals]);
  else sh.appendRow(vals);
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
  const out = { completed: {}, scores: {}, steps: {}, seen: {}, subtabs: {}, results: {} };
  [a.completed, b.completed].forEach(function (o) { Object.keys(o || {}).forEach(function (k) { if (o[k]) out.completed[k] = true; }); });
  [a.scores, b.scores].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const v = o[k] || {}, c = out.scores[k];
    out.scores[k] = !c ? v : { first: c.first != null ? c.first : v.first, best: Math.max(c.best || 0, v.best || 0), total: Math.max(c.total || 0, v.total || 0), rounds: Math.max(c.rounds || 0, v.rounds || 0) };
  }); });
  [a.steps, b.steps].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
    const v = o[k] || {}, c = out.steps[k];
    out.steps[k] = !c ? { cur: v.cur || 0, reached: v.reached || 0 } : { cur: v.cur || 0, reached: Math.max(c.reached || 0, v.reached || 0) };
  }); });
  ['seen', 'subtabs'].forEach(function (f) { [a[f], b[f]].forEach(function (o) { Object.keys(o || {}).forEach(function (k) {
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
  return { ok: true };
}
function readConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CFG);
  if (!sh || sh.getLastRow() < 2) return null;
  try { return JSON.parse(sh.getRange(2, 1).getValue()); } catch (err) { return null; }
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
}
