/*************************************************************
 * 동요샘 노래상자 — 학교 시트용 껍데기 (Shell.gs)
 * ------------------------------------------------------------
 * 이 파일 하나만 스프레드시트의 Apps Script 에 붙여넣고 웹앱으로 배포하면 됩니다.
 * 화면은 GitHub(musicalpe.github.io/DongyoSSAM)에서 열리고, 이 파일은 뒤에서
 * 우리 반 도장·투표·릴레이·바꾼 가사·좋아요를 시트에 읽고 쓰는 일만 합니다.
 * 화면이 새 판으로 바뀌어도 이 파일은 그대로 두면 됩니다.
 *
 * 배포 설정: 다음 사용자로 실행 = 나, 액세스 권한 = 모든 사용자
 * 처음 선생님 비밀번호: 1234 (화면의 '우리 반' 탭에서 바꿀 수 있어요)
 *************************************************************/

const SHELL_VERSION = 1;
const APP_URL = 'https://musicalpe.github.io/DongyoSSAM/';
const DEFAULT_COLLECTOR_URL = '';            // 비어 있으면 화면 저장소의 national.json 에서 읽음
const TZ_ = 'Asia/Seoul';
const TOKEN_TTL_SEC = 21600;                 // 선생님 로그인 유지: 6시간
const DEFAULT_ADMIN_PASSWORD = '1234';
const REGIONS_ = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주'];

const SH_STAMPS = '도장';
const SH_VOTES  = '투표';
const SH_RELAY  = '릴레이';
const SH_LYRICS = '바꾼 가사';
const SH_LIKES  = '좋아요';
const HEAD_ = {
  '도장':     ['곡번호', '곡 제목', '찍은 날'],
  '투표':     ['달', '곡번호', '곡 제목', '투표한 날'],
  '릴레이':   ['릴레이', '곡번호', '이어받은 날'],
  '바꾼 가사': ['번호', '곡번호', '곡 제목', '바꾼 가사', '올린 날'],
  '좋아요':   ['가사 ID', '누른 날']
};

/************ 입구 ************/
function doGet() {
  const self = ScriptApp.getService().getUrl();
  const app = APP_URL + '?s=' + encodeURIComponent(self);
  const html =
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:"Malgun Gothic",sans-serif;background:#F3FAFF;margin:0;padding:24px;color:#2B3A55}' +
    '.c{max-width:560px;margin:30px auto;background:#fff;border:2px solid #C9D8E8;border-radius:22px;overflow:hidden}' +
    '.h{background:#DDF0FF;padding:22px 24px}.h h1{margin:0 0 4px;font-size:22px}.h p{margin:0;color:#61708A}' +
    '.b{padding:22px 24px;line-height:1.6}a.btn{display:block;text-align:center;background:#4A8FE0;color:#fff;text-decoration:none;font-weight:800;padding:14px;border-radius:14px;font-size:17px;margin:14px 0}' +
    'code{display:block;background:#F3F4F6;padding:10px;border-radius:8px;word-break:break-all;font-size:12px}small{color:#61708A}</style></head><body>' +
    '<div class="c"><div class="h"><h1>동요샘 노래상자</h1><p>껍데기 설치 완료 · 껍데기 판 ' + SHELL_VERSION + '</p></div><div class="b">' +
    '<p>이 주소는 <b>설치 확인용</b>이에요. 프로그램은 아래 버튼으로 엽니다. 열린 주소를 즐겨찾기해 두세요.</p>' +
    '<a class="btn" href="' + app + '" target="_top">노래상자 열기</a>' +
    '<small>프로그램 주소</small><code>' + app + '</code>' +
    '<p><small>처음 선생님 비밀번호는 1234 입니다. 노래상자의 우리 반 탭에서 반 이름·지역을 정하고 비밀번호를 바꿔 주세요.</small></p>' +
    '</div></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('동요샘 노래상자 — 설치 확인')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const RPC_ALLOW_ = ['ping', 'getState', 'login', 'logout', 'changePassword', 'saveSettings',
  'toggleStamp', 'vote', 'joinRelay', 'addLyric', 'toggleLike', 'resend'];
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(body.fn || '');
    if (RPC_ALLOW_.indexOf(fn) === -1) throw new Error('알 수 없는 요청: ' + fn);
    const g = (typeof globalThis !== 'undefined') ? globalThis : this;
    const args = Array.isArray(body.args) ? body.args : [];
    out = { ok: true, result: g[fn].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/************ 공통 ************/
function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(HEAD_[name]); sh.setFrozenRows(1); sh.getRange(1, 1, 1, HEAD_[name].length).setFontWeight('bold'); }
  return sh;
}
function rows_(name) { return sheet_(name).getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ''; }); }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { const v = (b + 256) % 256; return (v < 16 ? '0' : '') + v.toString(16); }).join('');
}
function now_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd HH:mm'); }
function month_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM'); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[<>\u0000-\u0009\u000b-\u001f]/g, '').trim().slice(0, max); }
function int_(v) { const n = Math.round(Number(v)); if (!isFinite(n) || n < 1 || n > 9999) throw new Error('곡 번호가 올바르지 않습니다.'); return n; }

// 반 고유 열쇠: 처음 한 번 만들고, 시트를 복사해 가면(시트 ID가 바뀌면) 새로 만듦
function key_() {
  const p = props_(), id = ss_().getId();
  let k = p.getProperty('KEY');
  if (!k || p.getProperty('KEY_SS') !== id) { k = Utilities.getUuid(); p.setProperties({ KEY: k, KEY_SS: id }); }
  return k;
}
function pid_() { return hash_('dys|' + key_()).slice(0, 12); }

/************ 선생님 로그인 ************/
function adminHash_() { return props_().getProperty('ADMIN_HASH') || hash_(DEFAULT_ADMIN_PASSWORD); }
function login(pw) {
  if (hash_(String(pw || '')) !== adminHash_()) return { ok: false, message: '비밀번호가 맞지 않아요.' };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_' + token, '1', TOKEN_TTL_SEC);
  return { ok: true, token: token, defaultPw: !props_().getProperty('ADMIN_HASH') };
}
function logout(token) { if (token) CacheService.getScriptCache().remove('ADMIN_' + token); return { ok: true }; }
function need_(token) { if (!token || !CacheService.getScriptCache().get('ADMIN_' + token)) throw new Error('선생님 로그인이 필요해요. 다시 로그인해 주세요.'); }
function changePassword(token, pw) {
  need_(token);
  const v = String(pw || '').trim();
  if (v.length < 4) return { ok: false, message: '비밀번호는 4자 이상으로 정해 주세요.' };
  props_().setProperty('ADMIN_HASH', hash_(v));
  return { ok: true };
}
// (선택) 편집기에서 직접 비밀번호를 정할 때: setAdminPassword("새비밀번호")
function setAdminPassword(pw) { props_().setProperty('ADMIN_HASH', hash_(String(pw))); }

/************ 읽기 ************/
function ping() { return { shell: SHELL_VERSION, title: ss_().getName() }; }

function getState() {
  const p = props_(), pid = pid_(), m = month_();
  const vote = rows_(SH_VOTES).filter(function (r) { return String(r[0]) === m; })[0];
  return {
    shell: SHELL_VERSION, pid: pid, month: m,
    name: p.getProperty('CLASS_NAME') || '', region: p.getProperty('REGION') || '',
    configured: !!(p.getProperty('CLASS_NAME') && p.getProperty('REGION')),
    stamps: rows_(SH_STAMPS).map(function (r) { return Number(r[0]); }),
    vote: vote ? { month: m, song: Number(vote[1]) } : null,
    relays: rows_(SH_RELAY).map(function (r) { return String(r[0]); }),
    lyrics: rows_(SH_LYRICS).map(function (r) { return { id: pid + '-' + r[0], no: Number(r[0]), songId: Number(r[1]), text: String(r[3]), date: String(r[4]) }; }),
    likes: rows_(SH_LIKES).map(function (r) { return String(r[0]); })
  };
}

/************ 쓰기 (선생님만) ************/
function saveSettings(token, name, region) {
  need_(token);
  const n = clean_(name, 20);
  if (!n) throw new Error('반 이름을 적어 주세요. 예: 서울햇살초 5-1');
  if (REGIONS_.indexOf(region) === -1) throw new Error('지역을 골라 주세요.');
  props_().setProperties({ CLASS_NAME: n, REGION: region });
  return done_();
}

function toggleStamp(token, songId, title) {
  need_(token);
  const id = int_(songId), sh = sheet_(SH_STAMPS), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (Number(v[i][0]) === id) { sh.deleteRow(i + 1); return done_(); }
  sh.appendRow([id, clean_(title, 40), now_()]);
  return done_();
}

function vote(token, songId, title) {
  need_(token);
  const id = int_(songId), m = month_(), sh = sheet_(SH_VOTES), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === m) sh.deleteRow(i + 1);
  sh.appendRow([m, id, clean_(title, 40), now_()]);
  return done_();
}

function joinRelay(token, relayId, songId) {
  need_(token);
  const r = clean_(relayId, 30);
  if (!r) throw new Error('릴레이 이름이 없습니다.');
  if (rows_(SH_RELAY).some(function (x) { return String(x[0]) === r; })) return done_();
  sheet_(SH_RELAY).appendRow([r, int_(songId), now_()]);
  return done_();
}

function addLyric(token, songId, title, text) {
  need_(token);
  const id = int_(songId);
  const t = String(text == null ? '' : text).replace(/\r/g, '').split('\n').map(function (l) { return clean_(l, 60); }).filter(String).slice(0, 12).join('\n');
  if (!t) throw new Error('가사를 먼저 적어 주세요.');
  const all = rows_(SH_LYRICS);
  if (all.filter(function (r) { return Number(r[1]) === id; }).length >= 3) throw new Error('한 곡에 우리 반 가사는 세 개까지 올릴 수 있어요.');
  const no = all.reduce(function (a, r) { return Math.max(a, Number(r[0]) || 0); }, 0) + 1;
  sheet_(SH_LYRICS).appendRow([no, id, clean_(title, 40), t, now_()]);
  return done_();
}

function toggleLike(token, lyricKey) {
  need_(token);
  const k = clean_(lyricKey, 40);
  if (!k) throw new Error('가사를 찾지 못했어요.');
  if (k.indexOf(pid_() + '-') === 0) throw new Error('우리 반 가사에는 좋아요를 누를 수 없어요.');
  const sh = sheet_(SH_LIKES), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === k) { sh.deleteRow(i + 1); return done_(); }
  sh.appendRow([k, now_()]);
  return done_();
}

function resend(token) { need_(token); return done_(); }

function done_() { return { state: getState(), nat: report_() }; }

/************ 전국 수집기로 보내기 ************/
function collector_() {
  if (DEFAULT_COLLECTOR_URL) return DEFAULT_COLLECTOR_URL;
  const cache = CacheService.getScriptCache(), hit = cache.get('dys_collector');
  if (hit !== null) return hit;
  let url = '';
  try {
    const j = JSON.parse(UrlFetchApp.fetch(APP_URL + 'national.json', { muteHttpExceptions: true }).getContentText() || '{}');
    if (j && /^https:\/\/script\.google\.com\//.test(String(j.collector || ''))) url = String(j.collector);
  } catch (e) {}
  cache.put('dys_collector', url, url ? 21600 : 600);
  return url;
}

function report_() {
  const st = getState();
  if (!st.configured) return { skipped: 'settings' };
  const url = collector_();
  if (!url) return { skipped: 'nocollector' };
  const payload = {
    key: key_(), name: st.name, region: st.region, shell: SHELL_VERSION,
    stamps: st.stamps, vote: st.vote, relays: st.relays, likes: st.likes,
    lyrics: st.lyrics.map(function (l) { return { no: l.no, songId: l.songId, text: l.text }; })
  };
  try {
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain;charset=utf-8',
      payload: JSON.stringify({ fn: 'report', args: [payload] }), muteHttpExceptions: true, followRedirects: true });
    const j = JSON.parse(res.getContentText() || '{}');
    return j.ok ? { ok: true } : { error: j.error || '전국 수집기 응답 오류' };
  } catch (e) {
    return { error: '전국 수집기에 연결하지 못했어요.' };
  }
}
