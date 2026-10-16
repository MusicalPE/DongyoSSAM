/*************************************************************
 * 동요샘 노래상자 — 서버 (Collector.gs)
 * ------------------------------------------------------------
 * 동요샘 운영진 구글 시트 하나에만 설치합니다. 선생님들은 아무것도 설치하지 않고
 * 노래상자 화면에서 가입 신청 → 운영진 승인 → 로그인 해서 씁니다.
 * 모든 반의 도장·투표·릴레이·바꾼 가사가 이 시트에 반별로 쌓입니다.
 *
 * 배포 설정: 다음 사용자로 실행 = 나, 액세스 권한 = 모든 사용자
 * 배포 주소(…/exec)는 DongyoSSAM 저장소 national.json 의 collector 칸에 넣습니다.
 *
 * 운영진 비밀번호 정하기: 맨 아래쪽 setOperatorPassword 안의 "새비밀번호"를 바꾸고
 * 함수 목록에서 setOperatorPassword 를 골라 한 번 실행하세요. (정하기 전에는 1234)
 *
 * 운영 일은 노래상자 화면 맨 아래 '운영진' 메뉴에서 합니다.
 *  - 계정: 가입 신청 승인·거절, 비밀번호 잊은 선생님께 임시 비밀번호 만들어 주기
 *  - 곡 관리, 가사 확인(공개/숨김), 릴레이 곡, 참여 반 숨기기, 문의 이메일
 * 시트에서 직접 고쳐도 됩니다. 고친 내용은 2분 안에 화면에 반영됩니다.
 *************************************************************/

const SERVER_VERSION = 8;   // 8: 박자(4/4)가 날짜로 바뀌던 문제 고침  7: 엑셀로 곡 한꺼번에 올리기  6: 노래 파일(mp3)·악보 파일 올리기, 곡 종류(창작곡·편곡)·편곡자, 발표회 없는 개인 창작곡  5: 이번 달·학년도 집계, 로그인 유지, 알림 메일, 기록 내려받기, 공개 동의, 운영 기록  4: 곡 지우기·예시 곡 한꺼번에 지우기  3: 한 시트 방식 — 가입·승인·로그인, 계정 하나에 여러 반
const TZ_ = 'Asia/Seoul';
const BOARD_CACHE_SEC = 120;
const TOKEN_SEC = 21600;    // 로그인 6시간 (로그인 유지를 고르면 30일)
const KEEP_DAYS = 30;
const MAX_CLASSES = 20;     // 계정 하나에 만들 수 있는 반
const REGIONS_ = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주'];

const SH_ACC = '계정', SH_CLASSES = '반', SH_SONGS = '곡 목록', SH_SET = '설정', SH_LYRICS = '바꾼 가사', SH_RELAY = '릴레이 기록', SH_LOG = '운영 기록', SH_TODAY = '오늘의 동요';
const HEAD_ = {
  '계정': ['아이디', '비밀번호(암호화)', '선생님 이름', '학교', '지역', '신청한 날', '승인', '마지막 로그인', '메모'],
  '반': ['반 ID', '아이디', '반 이름', '지역', '도장', '투표 달', '투표 곡', '좋아요', '릴레이', '만든 날', '마지막 기록', '숨김'],
  '곡 목록': ['곡번호', '곡 제목', '작사', '작곡', '발표회', '발표 연도', '추천 학년군', '주제(쉼표로)', '계절', '조성', '박자', '음역', '가사', '소개', '수업 활용', '영상 주소', '악보 주소', '숨김', '공개 동의', '노래 파일', '곡 종류', '편곡'],
  '설정': ['항목', '값', '설명'],
  '바꾼 가사': ['가사 ID', '반 ID', '곡번호', '바꾼 가사', '반 이름', '지역', '받은 날', '공개'],
  '릴레이 기록': ['릴레이', '반 ID', '반 이름', '지역', '이어받은 때'],
  '운영 기록': ['때', '운영진', '한 일', '대상'],
  '오늘의 동요': ['날짜', '곡번호', '한마디']
};

/************ 입구 ************/
function doGet(e) {
  const q = (e && e.parameter) || {};
  let out;
  try { out = { ok: true, result: q.api === 'board' ? getBoard_() : { server: SERVER_VERSION, now: now_() } }; }
  catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return json_(out);
}
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const f = FN_[String(body.fn || '')];
    if (!f) throw new Error('알 수 없는 요청');
    out = { ok: true, result: f.apply(null, Array.isArray(body.args) ? body.args : []) };
  } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/************ 공통 ************/
function now_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd HH:mm:ss'); }
function month_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM'); }
function today_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd'); }
// 학년도 시작일 (3월 1일)
function yearStart_() { const t = today_(), y = Number(t.slice(0, 4)), m = Number(t.slice(5, 7)); return (m >= 3 ? y : y - 1) + '-03-01'; }
// 도장: [곡번호, 찍은 날] 목록 (예전 숫자만 있던 기록도 읽음)
function stampsOf_(cell) { return parse_(cell, []).map(function (x) { return Array.isArray(x) ? [num_(x[0]), String(x[1] || '')] : [num_(x), '']; }).filter(function (x) { return x[0]; }); }
// 시트가 날짜로 바꿔 버린 칸도 yyyy-MM-dd 글자로
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ_, 'yyyy-MM-dd') : String(v || '').trim().slice(0, 10); }
function tstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ_, 'yyyy-MM-dd HH:mm') : String(v || '').trim().slice(0, 16); }
function mstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ_, 'yyyy-MM') : String(v || '').trim().slice(0, 7); }
// 박자: 시트가 4/4를 날짜(4월 4일)로 바꿔 버린 칸도 '4/4'로 되돌려 읽기
function beatStr_(v) { return v instanceof Date ? Number(Utilities.formatDate(v, TZ_, 'MM')) + '/' + Number(Utilities.formatDate(v, TZ_, 'dd')) : String(v == null ? '' : v).trim(); }
// 곡 목록에서 글자 그대로 둘 칸(박자·조성·음역·발표 연도)을 '글자' 서식으로
function songTextCols_(sh, row, n) { sh.getRange(row, 10, n, 3).setNumberFormat('@'); }
function okMail_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e || '').trim()); }
const APP_URL = 'https://musicalpe.github.io/DongyoSSAM/';
// 운영진 알림 메일 (하루 40통까지)
function notify_(subject, body) {
  try {
    const to = String(settings_()['알림 이메일'] || '').trim(); if (!okMail_(to)) return;
    const p = PropertiesService.getScriptProperties(), k = 'MAIL_' + today_(), n = Number(p.getProperty(k) || 0);
    if (n >= 40) return;
    p.setProperty(k, String(n + 1));
    MailApp.sendEmail({ to: to, subject: '[동요샘 노래상자] ' + subject, body: body + '\n\n운영진 메뉴: ' + APP_URL + '?tab=op' });
  } catch (e) {}
}
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { const v = (b + 256) % 256; return (v < 16 ? '0' : '') + v.toString(16); }).join('');
}
function pwHash_(id, pw) { return hash_('dys-acc|' + id + '|' + pw); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[<>\u0000-\u0009\u000b-\u001f]/g, '').trim().slice(0, max); }
function num_(v) { const n = Math.round(Number(v)); return (isFinite(n) && n >= 1 && n <= 9999) ? n : 0; }
function isO_(v) { return /^[OoＯ○◯]$/.test(String(v || '').trim()); }
function parse_(s, d) { try { const v = JSON.parse(s || ''); return v == null ? d : v; } catch (e) { return d; } }
function lyricText_(t) { return String(t == null ? '' : t).replace(/\r/g, '').split('\n').map(function (x) { return clean_(x, 60); }).filter(String).slice(0, 12).join('\n'); }
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD_[name]); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD_[name].length).setFontWeight('bold');
    if (name === SH_SONGS) seedSongs_(sh);
    if (name === SH_TODAY) sh.getRange('A:A').setNumberFormat('@');
    if (name === SH_SET) sh.getRange(2, 1, 4, 3).setValues([
      ['릴레이 이름', '첫 번째 릴레이', '이름을 바꾸면 새 릴레이가 처음부터 시작돼요'],
      ['릴레이 곡번호', 3, '곡 목록의 곡번호'],
      ['문의 이메일', '', '비밀번호를 잊은 선생님이 연락할 주소 (화면에 보여요)'],
      ['알림 이메일', '', '가입 신청·새 가사가 오면 알림을 받을 운영진 주소 (화면에 안 보여요)']
    ]);
  }
  return sh;
}
function values_(name) { return sheet_(name).getDataRange().getValues(); }
function settings_() { const o = {}; values_(SH_SET).slice(1).forEach(function (r) { o[String(r[0]).trim()] = r[1]; }); return o; }
function putSetting_(k, val, note) {
  const sh = sheet_(SH_SET), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val, note || '']);
}
function bust_() { CacheService.getScriptCache().remove('board'); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('잠시 뒤 다시 눌러 주세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/************ 계정 ************/
function findAcc_(id) {
  const v = values_(SH_ACC);
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) return { row: i + 1, r: v[i] };
  return null;
}
function signup_(p) {
  p = p || {};
  const id = String(p.id || '').trim().toLowerCase();
  if (!/^[a-z0-9_]{4,20}$/.test(id)) throw new Error('아이디는 영어 소문자·숫자·밑줄로 4~20자로 정해 주세요.');
  const pw = String(p.pw || '');
  if (pw.length < 6) throw new Error('비밀번호는 6자 이상으로 정해 주세요.');
  const name = clean_(p.name, 20), school = clean_(p.school, 30), region = String(p.region || '');
  if (!name || !school) throw new Error('선생님 이름과 학교를 적어 주세요.');
  if (REGIONS_.indexOf(region) === -1) throw new Error('지역을 골라 주세요.');
  return withLock_(function () {
    const v = values_(SH_ACC);
    if (v.slice(1).some(function (r) { return String(r[0]) === id; })) throw new Error('이미 있는 아이디예요. 다른 아이디를 정해 주세요.');
    if (v.slice(1).filter(function (r) { return !String(r[6]).trim(); }).length >= 300) throw new Error('지금은 가입 신청이 너무 많아요. 잠시 뒤 다시 신청해 주세요.');
    sheet_(SH_ACC).appendRow([id, pwHash_(id, pw), name, school, region, now_(), '', '', '']);
    notify_('가입 신청: ' + name + ' 선생님 (' + school + ')', '새 가입 신청이 왔어요.\n\n아이디: ' + id + '\n이름: ' + name + '\n학교: ' + school + ' (' + region + ')\n\n운영진 메뉴의 계정 칸에서 승인해 주세요.');
    return { ok: true, message: '가입 신청을 받았어요. 운영진이 승인하면 로그인할 수 있어요.' };
  });
}
function login_(id, pw, keep) {
  id = String(id || '').trim().toLowerCase();
  const a = findAcc_(id);
  if (!a || String(a.r[1]) !== pwHash_(id, String(pw || ''))) return { ok: false, message: '아이디나 비밀번호가 맞지 않아요.' };
  const st = String(a.r[6]).trim();
  if (!isO_(st)) return { ok: false, message: st ? '가입이 승인되지 않은 계정이에요. 운영진에게 문의해 주세요.' : '아직 승인 전이에요. 운영진이 승인하면 로그인할 수 있어요.' };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('U_' + token, id, TOKEN_SEC);
  if (keep) {
    const p = PropertiesService.getScriptProperties(), all = p.getProperties(), nowMs = Date.now();
    Object.keys(all).forEach(function (k) { if (k.indexOf('T_') === 0 && Number(String(all[k]).split('|')[1]) < nowMs) p.deleteProperty(k); });
    p.setProperty('T_' + token, id + '|' + (nowMs + KEEP_DAYS * 86400000));
  }
  sheet_(SH_ACC).getRange(a.row, 8).setValue(now_());
  return { ok: true, token: token, me: me_(id) };
}
function logout_(token) { if (token) { CacheService.getScriptCache().remove('U_' + token); PropertiesService.getScriptProperties().deleteProperty('T_' + token); } return { ok: true }; }
function who_(token) {
  let id = token ? CacheService.getScriptCache().get('U_' + token) : null;
  if (!id && token) {
    const v = PropertiesService.getScriptProperties().getProperty('T_' + token);
    if (v && Number(v.split('|')[1]) > Date.now()) { id = v.split('|')[0]; CacheService.getScriptCache().put('U_' + token, id, TOKEN_SEC); }
  }
  if (!id) throw new Error('로그인이 필요해요. 다시 로그인해 주세요.');
  return id;
}
function changePw_(token, oldPw, newPw) {
  const id = who_(token), a = findAcc_(id);
  if (!a || String(a.r[1]) !== pwHash_(id, String(oldPw || ''))) return { ok: false, message: '지금 비밀번호가 맞지 않아요.' };
  if (String(newPw || '').length < 6) return { ok: false, message: '새 비밀번호는 6자 이상으로 정해 주세요.' };
  sheet_(SH_ACC).getRange(a.row, 2).setValue(pwHash_(id, String(newPw)));
  return { ok: true };
}

/************ 반 ************/
function classRow_(r) {
  const st = stampsOf_(r[4]);
  return { cid: String(r[0]), owner: String(r[1]), name: String(r[2]), region: String(r[3]), stamps: st.map(function (x) { return x[0]; }), stampDates: st,
    vote: mstr_(r[5]) && num_(r[6]) ? { month: mstr_(r[5]), song: num_(r[6]) } : null, likes: parse_(r[7], []), relays: parse_(r[8], []), hidden: isO_(r[11]) };
}
function me_(id) {
  const a = findAcc_(id);
  const lyr = {};
  values_(SH_LYRICS).slice(1).forEach(function (r) {
    const st = String(r[7] || '').trim();
    (lyr[String(r[1])] = lyr[String(r[1])] || []).push({ id: String(r[0]), songId: num_(r[2]), text: String(r[3]), status: isO_(st) ? 'O' : st ? 'X' : '' });
  });
  const classes = values_(SH_CLASSES).slice(1).filter(function (r) { return String(r[1]) === id; }).map(function (r) {
    const c = classRow_(r); delete c.owner; delete c.stampDates; c.lyrics = lyr[c.cid] || []; return c;
  });
  return { id: id, name: a ? String(a.r[2]) : '', school: a ? String(a.r[3]) : '', region: a ? String(a.r[4]) : '', classes: classes, month: month_() };
}
function meApi_(token) { return me_(who_(token)); }
function myClass_(id, cid) {
  const sh = sheet_(SH_CLASSES), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(cid)) {
    if (String(v[i][1]) !== id) throw new Error('내 반이 아니에요.');
    return { sh: sh, row: i + 1, r: v[i] };
  }
  throw new Error('반을 찾지 못했어요.');
}
function addClass_(token, name, region) {
  const id = who_(token), n = clean_(name, 20);
  if (!n) throw new Error('반 이름을 적어 주세요. 예: 서울햇살초 5-1');
  if (REGIONS_.indexOf(region) === -1) throw new Error('지역을 골라 주세요.');
  return withLock_(function () {
    const v = values_(SH_CLASSES);
    if (v.slice(1).filter(function (r) { return String(r[1]) === id; }).length >= MAX_CLASSES) throw new Error('반은 ' + MAX_CLASSES + '개까지 만들 수 있어요.');
    const cid = 'c' + hash_(id + '|' + Utilities.getUuid()).slice(0, 10);
    sheet_(SH_CLASSES).appendRow([cid, id, n, region, '[]', '', '', '[]', '[]', now_(), now_(), '']);
    bust_();
    const me = me_(id); me.newCid = cid; return me;
  });
}
function editClass_(token, cid, name, region) {
  const id = who_(token), n = clean_(name, 20);
  if (!n || REGIONS_.indexOf(region) === -1) throw new Error('반 이름과 지역을 정해 주세요.');
  return withLock_(function () {
    const c = myClass_(id, cid);
    c.sh.getRange(c.row, 3, 1, 2).setValues([[n, region]]);
    bust_(); return me_(id);
  });
}
function touch_(c) { c.sh.getRange(c.row, 11).setValue(now_()); }
function toggleStamp_(token, cid, songId) {
  const id = who_(token), s = num_(songId); if (!s) throw new Error('곡을 찾지 못했어요.');
  return withLock_(function () {
    const c = myClass_(id, cid), list = stampsOf_(c.r[4]), i = list.findIndex(function (x) { return x[0] === s; });
    if (i >= 0) list.splice(i, 1); else list.push([s, today_()]);
    c.sh.getRange(c.row, 5).setValue(JSON.stringify(list)); touch_(c); bust_(); return me_(id);
  });
}
function vote_(token, cid, songId) {
  const id = who_(token), s = num_(songId); if (!s) throw new Error('곡을 찾지 못했어요.');
  return withLock_(function () {
    const c = myClass_(id, cid);
    c.sh.getRange(c.row, 6).setNumberFormat('@'); c.sh.getRange(c.row, 6, 1, 2).setValues([[month_(), s]]); touch_(c); bust_(); return me_(id);
  });
}
function joinRelay_(token, cid, relayId) {
  const id = who_(token), rid = clean_(relayId, 30); if (!rid) throw new Error('릴레이를 찾지 못했어요.');
  return withLock_(function () {
    const c = myClass_(id, cid), list = parse_(c.r[8], []);
    if (list.indexOf(rid) === -1) {
      list.push(rid); c.sh.getRange(c.row, 9).setValue(JSON.stringify(list));
      sheet_(SH_RELAY).appendRow([rid, String(c.r[0]), String(c.r[2]), String(c.r[3]), now_()]);
    }
    touch_(c); bust_(); return me_(id);
  });
}
function addLyric_(token, cid, songId, text) {
  const id = who_(token), s = num_(songId), t = lyricText_(text);
  if (!s) throw new Error('곡을 찾지 못했어요.');
  if (!t) throw new Error('가사를 먼저 적어 주세요.');
  return withLock_(function () {
    const c = myClass_(id, cid), v = values_(SH_LYRICS);
    const mine = v.slice(1).filter(function (r) { return String(r[1]) === String(cid); });
    if (mine.filter(function (r) { return num_(r[2]) === s; }).length >= 3) throw new Error('한 곡에 우리 반 가사는 세 개까지 올릴 수 있어요.');
    sheet_(SH_LYRICS).appendRow([cid + '-' + (mine.length + 1), cid, s, t, String(c.r[2]), String(c.r[3]), now_(), '']);
    notify_('새 바꾼 가사: ' + String(c.r[2]), String(c.r[2]) + ' (' + String(c.r[3]) + ')에서 ' + s + '번 곡의 바꾼 가사를 올렸어요.\n\n' + t + '\n\n운영진 메뉴의 가사 확인 칸에서 공개하거나 숨겨 주세요.');
    touch_(c); return me_(id);
  });
}
function toggleLike_(token, cid, key) {
  const id = who_(token), k = clean_(key, 40);
  if (!k) throw new Error('가사를 찾지 못했어요.');
  if (k.indexOf(String(cid) + '-') === 0) throw new Error('우리 반 가사에는 좋아요를 누를 수 없어요.');
  return withLock_(function () {
    const c = myClass_(id, cid), list = parse_(c.r[7], []), i = list.indexOf(k);
    if (i >= 0) list.splice(i, 1); else list.push(k);
    c.sh.getRange(c.row, 8).setValue(JSON.stringify(list)); touch_(c); bust_(); return me_(id);
  });
}

/************ 전국 현황 ************/
function getBoard_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('board');
  if (hit) return JSON.parse(hit);
  const month = month_(), ok = {};
  values_(SH_ACC).slice(1).forEach(function (r) { if (isO_(r[6])) ok[String(r[0])] = 1; });
  const ys = yearStart_();
  const regions = {}, regionsYear = {}; REGIONS_.forEach(function (r) { regions[r] = { stamps: 0, classes: 0 }; regionsYear[r] = { stamps: 0, classes: 0 }; });
  const songClasses = {}, songMonth = {}, songYear = {}, votes = {}, likeCount = {}, classes = {};
  let classCount = 0, classYear = 0;
  values_(SH_CLASSES).slice(1).forEach(function (r) {
    const c = classRow_(r);
    if (!c.cid || c.hidden || !ok[c.owner] || !regions[c.region]) return;
    classes[c.cid] = c;
    const inMonth = c.stampDates.filter(function (x) { return x[1].slice(0, 7) === month; }).length;
    const inYear = c.stampDates.filter(function (x) { return !x[1] || x[1] >= ys; });
    if (inMonth) { regions[c.region].stamps += inMonth; regions[c.region].classes++; classCount++; }
    if (inYear.length) { regionsYear[c.region].stamps += inYear.length; regionsYear[c.region].classes++; classYear++; }
    c.stampDates.forEach(function (x) {
      songClasses[x[0]] = (songClasses[x[0]] || 0) + 1;
      if (!x[1] || x[1] >= ys) songYear[x[0]] = (songYear[x[0]] || 0) + 1;
      if (x[1].slice(0, 7) === month) songMonth[x[0]] = (songMonth[x[0]] || 0) + 1;
    });
    if (c.vote && c.vote.month === month) votes[c.vote.song] = (votes[c.vote.song] || 0) + 1;
    c.likes.forEach(function (k) { likeCount[k] = (likeCount[k] || 0) + 1; });
  });
  const relays = {};
  values_(SH_RELAY).slice(1).forEach(function (r) {
    const c = classes[String(r[1])]; if (!c) return;
    (relays[String(r[0])] = relays[String(r[0])] || []).push({ pid: c.cid, cls: c.name, region: c.region });
  });
  const lyrics = {};
  values_(SH_LYRICS).slice(1).forEach(function (r) {
    if (!isO_(r[7])) return;
    const c = classes[String(r[1])]; if (!c) return;
    const key = String(r[0]), sid = num_(r[2]);
    (lyrics[sid] = lyrics[sid] || []).push({ key: key, pid: c.cid, cls: c.name, region: c.region, text: String(r[3]), likes: likeCount[key] || 0 });
  });
  const songs = [];
  values_(SH_SONGS).slice(1).forEach(function (r) {
    const id = num_(r[0]);
    if (!id || !String(r[1]).trim() || isO_(r[17])) return;
    songs.push({ id: id, title: clean_(r[1], 40), lyricist: clean_(r[2], 40), composer: clean_(r[3], 40), concert: clean_(r[4], 40),
      year: num_(r[5]) || '', grade: clean_(r[6], 10), themes: String(r[7]).split(/[,，、]/).map(function (x) { return clean_(x, 10); }).filter(String),
      season: clean_(r[8], 20), key: clean_(r[9], 4), beat: clean_(beatStr_(r[10]), 6), range: clean_(r[11], 12),
      lyrics: String(r[12] || '').slice(0, 2000), intro: String(r[13] || '').slice(0, 600), tip: String(r[14] || '').slice(0, 600),
      video: /^https:\/\//.test(String(r[15])) ? String(r[15]).trim() : '', score: /^https:\/\//.test(String(r[16])) ? String(r[16]).trim() : '',
      audio: /^https:\/\//.test(String(r[19] || '')) ? String(r[19]).trim() : '', type: String(r[20] || '').trim() === '편곡' ? '편곡' : '창작곡', arranger: clean_(r[21], 40) });
  });
  const set = settings_();
  const relay = { id: clean_(set['릴레이 이름'], 30) || '첫 번째 릴레이', song: num_(set['릴레이 곡번호']) || (songs[0] ? songs[0].id : 0) };
  const contact = okMail_(set['문의 이메일']) ? String(set['문의 이메일']).trim() : '';
  const td = today_(), sched = {};
  values_(SH_TODAY).slice(1).forEach(function (r) { const d = dstr_(r[0]), id = num_(r[1]); if (id && d >= td && d <= td.slice(0, 7) + '-31') sched[d] = { song: id, note: clean_(r[2], 80) }; });
  const out = { month: month, today: td, yearStart: ys, classes: classCount, classesYear: classYear, regions: regions, regionsYear: regionsYear, songClasses: songClasses, songMonth: songMonth, songYear: songYear, todaySong: sched[td] || null, votes: votes, relays: relays, lyrics: lyrics, songs: songs, relay: relay, contact: contact, updated: now_() };
  try { cache.put('board', JSON.stringify(out), BOARD_CACHE_SEC); } catch (e) {}
  return out;
}
function clearBoardCache() { bust_(); }

/************ 운영진 ************/
function opHash_() { return PropertiesService.getScriptProperties().getProperty('OP_HASH') || hash_('1234'); }
function opLogin_(pw, name) {
  if (hash_(String(pw || '')) !== opHash_()) return { ok: false, message: '운영진 비밀번호가 맞지 않아요.' };
  const n = clean_(name, 20);
  if (!n) return { ok: false, message: '운영진 이름을 적어 주세요. 누가 처리했는지 기록에 남아요.' };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('OP_' + token, n, TOKEN_SEC);
  log_(n, '운영진 메뉴에 들어옴', '');
  return { ok: true, token: token, name: n };
}
function log_(who, what, target) { try { sheet_(SH_LOG).appendRow([now_(), who, what, String(target || '').slice(0, 80)]); } catch (e) {} }
function opWho_(token) { return CacheService.getScriptCache().get('OP_' + token) || '운영진'; }
function opLogout_(token) { if (token) CacheService.getScriptCache().remove('OP_' + token); return { ok: true }; }
function opNeed_(token) { if (!token || !CacheService.getScriptCache().get('OP_' + token)) throw new Error('운영진 로그인이 필요해요. 다시 로그인해 주세요.'); }
// 드라이브 권한 허용: 노래 파일·악보 올리기용. 한 번 실행하면 권한 창이 떠요 (폴더 '동요샘 노래상자 파일'이 생겨요)
function allowDrive() { Logger.log('파일 폴더: ' + folder_().getUrl()); }
// 알림 메일 권한 허용: 새 코드를 붙여 넣은 뒤 이 함수를 한 번 실행하면 메일 보내기 권한을 허용하는 창이 떠요 (메일은 보내지 않아요)
function allowMail() { Logger.log('오늘 보낼 수 있는 메일: ' + MailApp.getRemainingDailyQuota()); }
// 운영진 비밀번호 정하기: 아래 "새비밀번호"를 바꾸고 이 함수를 한 번 실행하세요
function setOperatorPassword() { PropertiesService.getScriptProperties().setProperty('OP_HASH', hash_('새비밀번호')); }

// 예전 시트에 새 칸 제목 채우기
function fixHeads_() {
  [SH_SONGS, SH_ACC, SH_CLASSES].forEach(function (n) {
    const sh = sheet_(n), h = HEAD_[n], cur = sh.getRange(1, 1, 1, h.length).getValues()[0];
    h.forEach(function (t, i) { if (String(cur[i]).trim() === '') sh.getRange(1, i + 1).setValue(t).setFontWeight('bold'); });
  });
}
function opData_(token) {
  opNeed_(token);
  fixHeads_();
  const songs = values_(SH_SONGS).slice(1).filter(function (r) { return num_(r[0]); }).map(function (r) {
    return { id: num_(r[0]), title: String(r[1]), lyricist: String(r[2]), composer: String(r[3]), concert: String(r[4]), year: r[5] === '' ? '' : String(r[5]),
      grade: String(r[6]), themes: String(r[7]), season: String(r[8]), key: String(r[9]), beat: beatStr_(r[10]), range: String(r[11]),
      lyrics: String(r[12]), intro: String(r[13]), tip: String(r[14]), video: String(r[15]), score: String(r[16]), hidden: isO_(r[17]), consent: isO_(r[18]), audio: String(r[19] || ''), type: String(r[20] || '').trim() === '편곡' ? '편곡' : '창작곡', arranger: String(r[21] || '') };
  });
  const lyrics = values_(SH_LYRICS).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const st = String(r[7] || '').trim();
    return { id: String(r[0]), songId: num_(r[2]), text: String(r[3]), cls: String(r[4]), region: String(r[5]), at: tstr_(r[6]), status: isO_(st) ? 'O' : st ? 'X' : '' };
  });
  const from = Utilities.formatDate(new Date(Date.now() - 62 * 86400000), TZ_, 'yyyy-MM') + '-01', schedule = {};
  values_(SH_TODAY).slice(1).forEach(function (r) { const d = dstr_(r[0]), id = num_(r[1]); if (id && d >= from) schedule[d] = { song: id, note: String(r[2] || '') }; });
  const logs = values_(SH_LOG).slice(1).slice(-60).reverse().map(function (r) { return { at: tstr_(r[0]), who: String(r[1]), what: String(r[2]), target: String(r[3]) }; });
  const accounts = values_(SH_ACC).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const st = String(r[6] || '').trim();
    return { id: String(r[0]), name: String(r[2]), school: String(r[3]), region: String(r[4]), at: tstr_(r[5]), last: tstr_(r[7]), status: isO_(st) ? 'O' : st ? 'X' : '' };
  });
  const classes = values_(SH_CLASSES).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { pid: String(r[0]), owner: String(r[1]), name: String(r[2]), region: String(r[3]), stamps: parse_(r[4], []).length, last: tstr_(r[10]), hidden: isO_(r[11]) };
  });
  const set = settings_();
  return { songs: songs, lyrics: lyrics, accounts: accounts, classes: classes, logs: logs, schedule: schedule, today: today_(), me: opWho_(token), contact: String(set['문의 이메일'] || ''), notify: String(set['알림 이메일'] || ''),
    relay: { id: String(set['릴레이 이름'] || ''), song: num_(set['릴레이 곡번호']) } };
}
function opAccount_(token, id, status) {
  opNeed_(token);
  const a = findAcc_(String(id)); if (!a) throw new Error('계정을 찾지 못했어요.');
  sheet_(SH_ACC).getRange(a.row, 7).setValue(status === 'O' ? 'O' : status === 'X' ? 'X' : '');
  sheet_(SH_ACC).getRange(a.row, 9).setValue((status === 'O' ? '승인' : status === 'X' ? '거절' : '다시 검토') + ' · ' + opWho_(token) + ' · ' + now_().slice(0, 16));
  log_(opWho_(token), status === 'O' ? '계정 승인' : status === 'X' ? '계정 거절' : '계정 다시 검토', id + ' ' + String(a.r[2]));
  bust_(); return opData_(token);
}
function opResetPw_(token, id) {
  opNeed_(token);
  const a = findAcc_(String(id)); if (!a) throw new Error('계정을 찾지 못했어요.');
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let pw = ''; for (let i = 0; i < 8; i++) pw += chars.charAt(Math.floor(Math.random() * chars.length));
  sheet_(SH_ACC).getRange(a.row, 2).setValue(pwHash_(String(id), pw));
  log_(opWho_(token), '임시 비밀번호 만듦', id);
  const d = opData_(token); d.tempPw = { id: String(id), pw: pw }; return d;
}
function opContact_(token, email) {
  opNeed_(token);
  const e = String(email || '').trim();
  if (e && !okMail_(e)) throw new Error('이메일 주소 모양이 아니에요.');
  putSetting_('문의 이메일', e, '비밀번호를 잊은 선생님이 연락할 주소 (화면에 보여요)');
  log_(opWho_(token), '문의 이메일 바꿈', e);
  bust_(); return opData_(token);
}
function opNotify_(token, email) {
  opNeed_(token);
  const e = String(email || '').trim();
  if (e && !okMail_(e)) throw new Error('이메일 주소 모양이 아니에요.');
  putSetting_('알림 이메일', e, '가입 신청·새 가사가 오면 알림을 받을 운영진 주소 (화면에 안 보여요)');
  log_(opWho_(token), '알림 이메일 바꿈', e);
  if (e) { try { MailApp.sendEmail({ to: e, subject: '[동요샘 노래상자] 알림 메일 확인', body: '이제 가입 신청과 새 바꾼 가사가 오면 이 주소로 알려 드려요.\n\n운영진 메뉴: ' + APP_URL + '?tab=op' }); } catch (err) { throw new Error('알림 메일을 보내지 못했어요. 서버를 배포할 때 메일 보내기 권한을 허용했는지 확인해 주세요.'); } }
  return opData_(token);
}
// 오늘의 동요 달력: 한 달치를 통째로 바꿔 저장 (entries: [{date, song, note}])
function opToday_(token, month, entries) {
  opNeed_(token);
  const m = String(month || '');
  if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('달을 알 수 없어요.');
  return withLock_(function () {
    const sh = sheet_(SH_TODAY), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (dstr_(v[i][0]).slice(0, 7) === m) sh.deleteRow(i + 1);
    const rows = (Array.isArray(entries) ? entries : []).filter(function (e) { return e && String(e.date).slice(0, 7) === m && /^\d{4}-\d{2}-\d{2}$/.test(String(e.date)) && num_(e.song); })
      .map(function (e) { return [String(e.date), num_(e.song), clean_(e.note, 80)]; });
    if (rows.length) { sh.getRange('A:A').setNumberFormat('@'); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows); }
    log_(opWho_(token), '오늘의 동요 달력 저장', m + ' · ' + rows.length + '일');
    bust_(); return opData_(token);
  });
}
// 엑셀로 여러 곡 한꺼번에 올리기: 같은 제목이 이미 있으면 건너뜀, 잘못된 줄은 이유와 함께 알려 줌
function opBulkSongs_(token, list) {
  opNeed_(token);
  if (!Array.isArray(list) || !list.length) throw new Error('올릴 곡이 없어요.');
  if (list.length > 200) throw new Error('한 번에 200곡까지 올릴 수 있어요.');
  return withLock_(function () {
    const sh = sheet_(SH_SONGS), v = sh.getDataRange().getValues(), have = {};
    let max = 0;
    for (let i = 1; i < v.length; i++) { max = Math.max(max, num_(v[i][0])); have[String(v[i][1]).trim()] = 1; }
    const rows = [], skipped = [];
    list.forEach(function (s, k) {
      const t = clean_(s && s.title, 40), at = (s && s.line) || (k + 1);
      if (t && have[t]) { skipped.push(at + '번째 줄 「' + t + '」: 같은 제목이 이미 있어요'); return; }
      try { rows.push(songData_(s, max + rows.length + 1)); have[t] = 1; }
      catch (e) { skipped.push(at + '번째 줄' + (t ? ' 「' + t + '」' : '') + ': ' + e.message); }
    });
    if (rows.length) {
      const start = sh.getLastRow() + 1;
      songTextCols_(sh, start, rows.length);
      sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
      sh.getRange(start, 13, rows.length, 3).setWrap(true);
    }
    log_(opWho_(token), '엑셀로 곡 올림', rows.length + '곡' + (skipped.length ? ' · 건너뜀 ' + skipped.length : ''));
    bust_();
    const d = opData_(token); d.bulk = { added: rows.length, skipped: skipped }; return d;
  });
}
// 노래 파일(mp3)·악보(PDF·그림) 올리기: 운영진 구글 드라이브의 '동요샘 노래상자 파일' 폴더에 저장하고 링크로 공유
const UPLOAD_MAX = 20 * 1024 * 1024;
function folder_() {
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('DRIVE_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder('동요샘 노래상자 파일');
  p.setProperty('DRIVE_FOLDER', f.getId());
  return f;
}
function opUpload_(token, kind, name, mime, data) {
  opNeed_(token);
  mime = String(mime || '');
  const okType = kind === 'audio' ? /^audio\//.test(mime) || mime === 'video/mp4' : /^image\//.test(mime) || mime === 'application/pdf';
  if (!okType) throw new Error(kind === 'audio' ? 'mp3 같은 노래 파일만 올릴 수 있어요.' : '악보는 PDF나 그림 파일만 올릴 수 있어요.');
  const bytes = Utilities.base64Decode(String(data || ''));
  if (!bytes.length) throw new Error('파일이 비어 있어요.');
  if (bytes.length > UPLOAD_MAX) throw new Error('파일이 너무 커요. 20MB까지 올릴 수 있어요.');
  let file;
  try { file = folder_().createFile(Utilities.newBlob(bytes, mime, clean_(name, 80) || '파일')); }
  catch (e) { throw new Error('드라이브에 저장하지 못했어요. 편집기에서 allowDrive 를 한 번 실행해 권한을 허용해 주세요.'); }
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { throw new Error('파일은 올렸지만 링크 공유를 켜지 못했어요. 드라이브에서 직접 "링크가 있는 모든 사용자"로 바꿔 주세요.'); }
  log_(opWho_(token), kind === 'audio' ? '노래 파일 올림' : '악보 파일 올림', file.getName());
  return { id: file.getId(), name: file.getName(), url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
}
// 기록 내려받기: 반별·곡별
function opExport_(token) {
  opNeed_(token);
  const month = month_(), ys = yearStart_(), acc = {};
  values_(SH_ACC).slice(1).forEach(function (r) { acc[String(r[0])] = { name: String(r[2]), school: String(r[3]) }; });
  const lyrCount = {}, lyrSong = {};
  values_(SH_LYRICS).slice(1).forEach(function (r) { lyrCount[String(r[1])] = (lyrCount[String(r[1])] || 0) + 1; if (isO_(r[7])) lyrSong[num_(r[2])] = (lyrSong[num_(r[2])] || 0) + 1; });
  const songStat = {};
  const classes = values_(SH_CLASSES).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const c = classRow_(r), a = acc[c.owner] || { name: '', school: '' };
    const m = c.stampDates.filter(function (x) { return x[1].slice(0, 7) === month; }).length, y = c.stampDates.filter(function (x) { return !x[1] || x[1] >= ys; }).length;
    c.stampDates.forEach(function (x) { const st = songStat[x[0]] = songStat[x[0]] || { all: 0, year: 0, month: 0 }; st.all++; if (!x[1] || x[1] >= ys) st.year++; if (x[1].slice(0, 7) === month) st.month++; });
    return [c.name, c.region, a.school, a.name, c.owner, c.stamps.length, y, m, c.vote && c.vote.month === month ? c.vote.song : '', lyrCount[c.cid] || 0, c.relays.length, dstr_(r[9]), tstr_(r[10]), c.hidden ? '숨김' : ''];
  });
  const votes = {};
  values_(SH_CLASSES).slice(1).forEach(function (r) { if (mstr_(r[5]) === month && num_(r[6])) votes[num_(r[6])] = (votes[num_(r[6])] || 0) + 1; });
  const songs = values_(SH_SONGS).slice(1).filter(function (r) { return num_(r[0]); }).map(function (r) {
    const id = num_(r[0]), st = songStat[id] || { all: 0, year: 0, month: 0 };
    return [id, String(r[1]), String(r[20] || '').trim() === '편곡' ? '편곡' : '창작곡', String(r[2]), String(r[3]), String(r[21] || ''), String(r[4]) || '발표회 없음', st.all, st.year, st.month, votes[id] || 0, lyrSong[id] || 0, isO_(r[17]) ? '숨김' : '', isO_(r[18]) ? 'O' : ''];
  });
  log_(opWho_(token), '기록 내려받기', '');
  return { month: month, yearStart: ys,
    classHead: ['반 이름', '지역', '학교', '선생님', '아이디', '스티커(모두)', '스티커(올해)', '스티커(이번 달)', '이번 달 투표 곡번호', '올린 가사', '릴레이 참여', '만든 날', '마지막 기록', '숨김'], classes: classes,
    songHead: ['곡번호', '곡 제목', '곡 종류', '작사', '작곡', '편곡', '발표회', '부른 반(모두)', '부른 반(올해)', '부른 반(이번 달)', '이번 달 투표', '공개된 바꾼 가사', '숨김', '공개 동의'], songs: songs };
}
// 곡 한 줄 만들기 (확인할 것을 확인하고, 시트에 들어갈 칸 순서대로)
function songData_(s, id) {
  s = s || {};
  const title = clean_(s.title, 40);
  if (!title) throw new Error('곡 제목을 적어 주세요.');
  if (!s.consent) throw new Error('작사·작곡·편곡 선생님께 공개 동의를 받았는지 확인해 주세요.');
  if (s.type === '편곡' && !clean_(s.arranger, 40)) throw new Error('편곡한 선생님 이름을 적어 주세요.');
  if (/^[-–—\s]+$/.test(String(s.concert || ''))) s.concert = '';
  const link = function (u) { u = String(u || '').trim(); if (u && !/^https:\/\//.test(u)) throw new Error('영상·노래·악보 주소는 https:// 로 시작해야 해요.'); return u.slice(0, 500); };
  const long = function (t, n) { return String(t == null ? '' : t).replace(/\r/g, '').replace(/[<>]/g, '').slice(0, n); };
  return [id, title, clean_(s.lyricist, 40), clean_(s.composer, 40), clean_(s.concert, 40), num_(s.year) || '', clean_(s.grade, 10),
    clean_(s.themes, 60), clean_(s.season, 20), clean_(s.key, 4), clean_(s.beat, 6), clean_(s.range, 12),
    long(s.lyrics, 2000), long(s.intro, 600), long(s.tip, 600), link(s.video), link(s.score), s.hidden ? 'O' : '', 'O', link(s.audio), s.type === '편곡' ? '편곡' : '창작곡', s.type === '편곡' ? clean_(s.arranger, 40) : ''];
}
function opSaveSong_(token, s) {
  opNeed_(token);
  s = s || {};
  songData_(s, 0);
  const title = clean_(s.title, 40);
  return withLock_(function () {
    const sh = sheet_(SH_SONGS), v = sh.getDataRange().getValues();
    let id = num_(s.id), row = 0, max = 0;
    for (let i = 1; i < v.length; i++) { const n = num_(v[i][0]); max = Math.max(max, n); if (id && n === id) row = i + 1; }
    if (!id) id = max + 1;
    const data = songData_(s, id);
    if (!row) row = sh.getLastRow() + 1;
    songTextCols_(sh, row, 1);
    sh.getRange(row, 1, 1, data.length).setValues([data]);
    sh.getRange(row, 13, 1, 3).setWrap(true);
    log_(opWho_(token), num_(s.id) ? '곡 고침' : '새 곡 올림', id + '. ' + title);
    bust_(); return opData_(token);
  });
}
function opDeleteSong_(token, id) {
  opNeed_(token);
  const n = num_(id); if (!n) throw new Error('곡을 찾지 못했어요.');
  return withLock_(function () {
    const sh = sheet_(SH_SONGS), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (num_(v[i][0]) === n) { log_(opWho_(token), '곡 지움', n + '. ' + String(v[i][1])); sh.deleteRow(i + 1); bust_(); return opData_(token); }
    throw new Error('곡을 찾지 못했어요.');
  });
}
// 처음 설치 때 들어간 예시 곡(작사·작곡이 "예시"로 시작)만 지움
function opClearExamples_(token) {
  opNeed_(token);
  return withLock_(function () {
    const sh = sheet_(SH_SONGS), v = sh.getDataRange().getValues();
    let n = 0;
    for (let i = v.length - 1; i >= 1; i--) if (/^예시/.test(String(v[i][2])) || /^예시/.test(String(v[i][3]))) { sh.deleteRow(i + 1); n++; }
    log_(opWho_(token), '예시 곡 모두 지움', n + '곡');
    bust_(); return opData_(token);
  });
}
function opLyric_(token, id, status) {
  opNeed_(token);
  const sh = sheet_(SH_LYRICS), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    sh.getRange(i + 1, 8).setValue(status === 'O' ? 'O' : status === 'X' ? 'X' : '');
    log_(opWho_(token), status === 'O' ? '가사 공개' : status === 'X' ? '가사 숨김' : '가사 확인 전으로', String(v[i][4]) + ' · ' + String(v[i][3]).split('\n')[0]);
    bust_(); return opData_(token);
  }
  throw new Error('그 가사를 찾지 못했어요.');
}
function opRelay_(token, name, songId) {
  opNeed_(token);
  const n = clean_(name, 30), id = num_(songId);
  if (!n || !id) throw new Error('릴레이 이름과 곡을 정해 주세요.');
  putSetting_('릴레이 이름', n); putSetting_('릴레이 곡번호', id);
  log_(opWho_(token), '릴레이 정함', n + ' · ' + id + '번 곡');
  bust_(); return opData_(token);
}
function opClass_(token, cid, hidden) {
  opNeed_(token);
  const sh = sheet_(SH_CLASSES), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(cid)) {
    sh.getRange(i + 1, 12).setValue(hidden ? 'O' : '');
    log_(opWho_(token), hidden ? '반 숨김' : '반 다시 보이기', String(v[i][2]));
    bust_(); return opData_(token);
  }
  throw new Error('그 반을 찾지 못했어요.');
}

/* 요청 이름 → 함수 (맨 아래에 두어야 위 함수들이 모두 준비된 뒤 묶임) */
const FN_ = {
  signup: signup_, login: login_, logout: logout_, me: meApi_, changePw: changePw_,
  addClass: addClass_, editClass: editClass_,
  toggleStamp: toggleStamp_, vote: vote_, joinRelay: joinRelay_, addLyric: addLyric_, toggleLike: toggleLike_,
  opLogin: opLogin_, opLogout: opLogout_, opData: opData_, opSaveSong: opSaveSong_, opLyric: opLyric_, opRelay: opRelay_,
  opClass: opClass_, opAccount: opAccount_, opResetPw: opResetPw_, opContact: opContact_,
  opDeleteSong: opDeleteSong_, opClearExamples: opClearExamples_, opNotify: opNotify_, opExport: opExport_, opToday: opToday_, opUpload: opUpload_, opBulkSongs: opBulkSongs_
};

/************ 곡 목록 예시 (처음 한 번) ************/
function seedSongs_(sh) {
  const ex = [
    [1, '운동장 한 바퀴', '예시 회원 A', '예시 회원 A', '제3회 동요샘 창작동요 발표회', 2026, '1~2학년', '학교생활, 몸놀이', '가을', 'D', '2/4', '레~레',
      '운동장 한 바퀴 뛰어 볼까\n발소리 콩콩 바람도 콩콩\n숨이 차도 괜찮아\n친구 손 잡고 또 한 바퀴', '아이들이 달리기 전에 부르며 몸을 깨우는 곡입니다.', "'콩콩'에서 제자리 뛰기, '한 바퀴'에서 짝과 손잡고 한 바퀴 돌기."],
    [2, '은행잎 편지', '예시 회원 B', '예시 회원 C', '제3회 동요샘 창작동요 발표회', 2026, '3~4학년', '계절, 자연', '가을', 'F', '3/4', '도~레',
      '노란 은행잎에 편지를 써요\n바람아 바람아 전해 주렴\n멀리 사는 할머니께\n가을이 왔다고', '가을 길을 걸으며 떠오른 이야기를 담은 3박자 노래입니다.', "국어 '마음을 전하는 글' 단원과 엮어 은행잎 카드에 편지 쓰기."],
    [3, '우리 반 무지개', '예시 회원 D', '예시 회원 D', '제2회 동요샘 창작동요 발표회', 2025, '전 학년', '친구, 학교생활', '봄', 'G', '4/4', '레~미',
      '빨간 마음 주황 웃음\n노란 목소리 모여서\n우리 반은 무지개\n하나하나 달라서 예뻐', '3월 첫 주 학급 세우기에 많이 쓰인 곡입니다.', '일곱 색 모둠으로 나누어 한 소절씩 이어 부르기.'],
    [4, '할머니 손 칼국수', '예시 회원 E', '예시 회원 A', '제2회 동요샘 창작동요 발표회', 2025, '3~4학년', '가족', '겨울', 'C', '6/8', '도~도',
      '쓱쓱 밀고 톡톡 썰고\n할머니 손은 마술 손\n김이 모락모락\n우리 집이 따뜻해져요', '6/8박자의 흔들리는 느낌이 부엌 소리와 잘 어울리는 곡입니다.', '흉내말에 맞춰 몸짓 만들기.'],
    [5, '빗방울 합창단', '예시 회원 F', '예시 회원 F', '제1회 동요샘 창작동요 발표회', 2024, '1~2학년', '자연, 계절', '여름', 'F', '2/4', '도~라',
      '똑 똑 똑 지붕 위에서\n또 똑 똑 창문 위에서\n빗방울 합창단이\n여름 노래 불러요', '음역이 좁아 누구나 편하게 부를 수 있습니다.', "컵으로 '똑 똑 똑' 리듬 치기."],
    [6, '내일의 나에게', '예시 회원 G', '예시 회원 D', '제1회 동요샘 창작동요 발표회', 2024, '5~6학년', '꿈, 감정', '겨울', 'A♭', '4/4', '시~파',
      '오늘 넘어진 나를 안아 줘\n괜찮다고 말해 줘\n내일의 나에게 보내는\n작은 노래 하나', '졸업을 앞둔 6학년을 생각하며 만든 곡입니다.', '졸업식 합창곡이나 진로 수업 마무리 곡.'],
    [7, '줄넘기 리듬송', '예시 회원 A', '예시 회원 A', '제3회 동요샘 창작동요 발표회', 2026, '3~4학년', '몸놀이, 학교생활', '봄', 'E', '2/4', '미~미',
      '하나 둘 넘고 셋 넷 넘고\n줄이 돌면 나도 돌아\n걸려도 다시 하나부터\n우리 같이 뛰어 보자', '줄넘기 속도에 맞춘 빠르기로 만든 곡입니다.', '긴 줄넘기 들어가기 연습에 맞춰 부르기.'],
    [8, '별에게 물어봐', '예시 회원 H', '예시 회원 F', '제2회 동요샘 창작동요 발표회', 2025, '5~6학년', '자연, 꿈', '여름', 'D', '3/4', '라~레',
      '밤하늘 별에게 물어봐\n너는 얼마나 멀리서 왔니\n반짝이는 대답 대신\n가만히 나를 비춰 줘', '과학 별 단원과 연결해 쓰려고 만든 곡입니다.', '별자리 관찰 뒤 감상곡으로, 2부 합창으로 화음 쌓기.']
  ];
  sh.getRange(2, 1, ex.length, 15).setValues(ex);
  sh.getRange(2, 13, ex.length, 3).setWrap(true);
}
