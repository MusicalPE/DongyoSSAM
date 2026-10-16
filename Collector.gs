/*************************************************************
 * 동요샘 노래상자 — 전국 수집기 (Collector.gs)
 * ------------------------------------------------------------
 * 동요샘 운영진 시트 하나에만 설치합니다. 학교 껍데기들이 보낸 우리 반 기록을
 * 모아 전국 무대·동요샘 차트·이어 부르기 릴레이·바꾼 가사 모음을 만듭니다.
 *
 * 배포 설정: 다음 사용자로 실행 = 나, 액세스 권한 = 모든 사용자
 * 배포 주소는 DongyoSSAM 저장소 national.json 의 collector 칸에 넣습니다.
 *
 * 운영진이 시트에서 하는 일
 *  - 곡 목록 탭: 한 줄에 한 곡. 고치면 2분 안에 모든 학교 화면에 반영됩니다.
 *    (처음 열 때 예시 곡 8개가 들어가 있어요. 실제 곡으로 바꿔 주세요. 곡번호는 바꾸지 마세요)
 *  - 설정 탭: 지금 진행하는 릴레이 이름과 곡번호
 *  - 바꾼 가사 탭: '공개' 칸에 O 를 적은 가사만 전국에 보입니다
 *  - 반 탭: '숨김' 칸에 O 를 적으면 그 반 기록을 전국 무대에서 뺍니다
 *  위 일은 모두 노래상자 화면 맨 아래 '운영진' 메뉴에서도 할 수 있습니다.
 *  처음 운영진 비밀번호: 1234 (운영진 메뉴에서 꼭 바꿔 주세요)
 *************************************************************/

const COLLECTOR_VERSION = 2;   // 2: 운영진 메뉴(곡 관리·가사 확인·릴레이·참여 반)
const TZ_ = 'Asia/Seoul';
const BOARD_CACHE_SEC = 120;
const REGIONS_ = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주'];

const SH_SONGS = '곡 목록', SH_SET = '설정', SH_CLASSES = '반', SH_LYRICS = '바꾼 가사', SH_RELAY = '릴레이 기록';
const HEAD_ = {
  '곡 목록': ['곡번호', '곡 제목', '작사', '작곡', '발표회', '발표 연도', '추천 학년군', '주제(쉼표로)', '계절', '조성', '박자', '음역', '가사', '소개', '수업 활용', '영상 주소', '악보 주소', '숨김'],
  '설정': ['항목', '값', '설명'],
  '반': ['반 ID', '반 이름', '지역', '도장', '투표 달', '투표 곡', '좋아요', '처음 받은 날', '마지막 받은 날', '숨김'],
  '바꾼 가사': ['가사 ID', '반 ID', '곡번호', '바꾼 가사', '반 이름', '지역', '받은 날', '공개'],
  '릴레이 기록': ['릴레이', '반 ID', '반 이름', '지역', '이어받은 때']
};

/************ 입구 ************/
function doGet(e) {
  const q = (e && e.parameter) || {};
  let out;
  try {
    if (q.api === 'board') out = getBoard_();
    else out = { collector: COLLECTOR_VERSION, now: now_() };
    out = { ok: true, result: out };
  } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return json_(out);
}
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = Array.isArray(body.args) ? body.args : [];
    if (body.fn === 'report') out = report_(args[0] || {});
    else if (OP_FN_[body.fn]) out = OP_FN_[body.fn].apply(null, args);
    else throw new Error('알 수 없는 요청');
    out = { ok: true, result: out };
  } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/************ 공통 ************/
function now_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd HH:mm:ss'); }
function month_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM'); }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { const v = (b + 256) % 256; return (v < 16 ? '0' : '') + v.toString(16); }).join('');
}
function clean_(s, max) { return String(s == null ? '' : s).replace(/[<>\u0000-\u0009\u000b-\u001f]/g, '').trim().slice(0, max); }
function num_(v) { const n = Math.round(Number(v)); return (isFinite(n) && n >= 1 && n <= 9999) ? n : 0; }
function isO_(v) { return /^[OoＯ○◯]$/.test(String(v || '').trim()); }
function parse_(s, d) { try { const v = JSON.parse(s || ''); return v == null ? d : v; } catch (e) { return d; } }

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD_[name]); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD_[name].length).setFontWeight('bold');
    if (name === SH_SONGS) seedSongs_(sh);
    if (name === SH_SET) sh.getRange(2, 1, 2, 3).setValues([
      ['릴레이 이름', '첫 번째 릴레이', '이름을 바꾸면 새 릴레이가 처음부터 시작돼요'],
      ['릴레이 곡번호', 3, '곡 목록의 곡번호']
    ]);
  }
  return sh;
}
function values_(name) { return sheet_(name).getDataRange().getValues(); }

/************ 학교 껍데기에서 오는 보고 (우리 반 전체 상태를 통째로) ************/
function report_(p) {
  const key = String(p.key || '');
  if (key.length < 20) throw new Error('반 열쇠가 올바르지 않습니다.');
  const pid = hash_('dys|' + key).slice(0, 12);
  const name = clean_(p.name, 20), region = String(p.region || '');
  if (!name || REGIONS_.indexOf(region) === -1) throw new Error('반 이름과 지역이 필요합니다.');

  const uniq = function (a) { return a.filter(function (x, i) { return x && a.indexOf(x) === i; }); };
  const stamps = uniq((Array.isArray(p.stamps) ? p.stamps : []).slice(0, 500).map(num_));
  const vote = p.vote && /^\d{4}-\d{2}$/.test(String(p.vote.month)) && num_(p.vote.song) ? [String(p.vote.month), num_(p.vote.song)] : ['', ''];
  const likes = uniq((Array.isArray(p.likes) ? p.likes : []).slice(0, 1000).map(function (k) { return clean_(k, 40); }));
  const relays = uniq((Array.isArray(p.relays) ? p.relays : []).slice(0, 100).map(function (k) { return clean_(k, 30); }));
  const lyrics = (Array.isArray(p.lyrics) ? p.lyrics : []).slice(0, 300).map(function (l) {
    return { id: pid + '-' + num_(l.no), songId: num_(l.songId), text: String(l.text == null ? '' : l.text).replace(/\r/g, '').split('\n').map(function (x) { return clean_(x, 60); }).filter(String).slice(0, 12).join('\n') };
  }).filter(function (l) { return l.songId && l.text && !/-0$/.test(l.id); });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('잠시 뒤 다시 보내 주세요.');
  try {
    const t = now_();
    // 반
    const cs = sheet_(SH_CLASSES), cv = cs.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < cv.length; i++) if (String(cv[i][0]) === pid) { row = i + 1; break; }
    const data = [name, region, JSON.stringify(stamps), vote[0], vote[1], JSON.stringify(likes)];
    if (row) { cs.getRange(row, 2, 1, 6).setValues([data]); cs.getRange(row, 9).setValue(t); }
    else cs.appendRow([pid].concat(data, [t, t, '']));
    // 릴레이: 수집기가 처음 받은 순서대로 줄을 정함
    const rs = sheet_(SH_RELAY), rv = rs.getDataRange().getValues(), have = {};
    for (let i = 1; i < rv.length; i++) if (String(rv[i][1]) === pid) have[String(rv[i][0])] = 1;
    relays.forEach(function (r) { if (!have[r]) rs.appendRow([r, pid, name, region, t]); });
    // 바꾼 가사: 새 가사만 추가 (공개 칸은 운영진이 O 로)
    const ls = sheet_(SH_LYRICS), lv = ls.getDataRange().getValues(), seen = {};
    for (let i = 1; i < lv.length; i++) seen[String(lv[i][0])] = 1;
    lyrics.forEach(function (l) { if (!seen[l.id]) ls.appendRow([l.id, pid, l.songId, l.text, name, region, t, '']); });
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('board');
  return { pid: pid };
}

/************ 전국 현황 ************/
function getBoard_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('board');
  if (hit) return JSON.parse(hit);
  const month = month_();
  const regions = {}; REGIONS_.forEach(function (r) { regions[r] = { stamps: 0, classes: 0 }; });
  const songClasses = {}, votes = {}, likeCount = {}, classes = {};
  let classCount = 0;

  const cv = values_(SH_CLASSES);
  for (let i = 1; i < cv.length; i++) {
    const pid = String(cv[i][0]); if (!pid || isO_(cv[i][9])) continue;
    const region = String(cv[i][2]); if (!regions[region]) continue;
    classes[pid] = { name: String(cv[i][1]), region: region };
    const stamps = parse_(cv[i][3], []);
    if (stamps.length) { regions[region].stamps += stamps.length; regions[region].classes++; classCount++; }
    stamps.forEach(function (id) { songClasses[id] = (songClasses[id] || 0) + 1; });
    if (String(cv[i][4]) === month && num_(cv[i][5])) votes[num_(cv[i][5])] = (votes[num_(cv[i][5])] || 0) + 1;
    parse_(cv[i][6], []).forEach(function (k) { likeCount[k] = (likeCount[k] || 0) + 1; });
  }

  const relays = {};
  const rv = values_(SH_RELAY);
  for (let i = 1; i < rv.length; i++) {
    const c = classes[String(rv[i][1])]; if (!c) continue;
    (relays[String(rv[i][0])] = relays[String(rv[i][0])] || []).push({ pid: String(rv[i][1]), cls: c.name, region: c.region });
  }

  const lyrics = {};
  const lv = values_(SH_LYRICS);
  for (let i = 1; i < lv.length; i++) {
    if (!isO_(lv[i][7])) continue;
    const c = classes[String(lv[i][1])]; if (!c) continue;
    const key = String(lv[i][0]), sid = num_(lv[i][2]);
    (lyrics[sid] = lyrics[sid] || []).push({ key: key, pid: String(lv[i][1]), cls: c.name, region: c.region, text: String(lv[i][3]), likes: likeCount[key] || 0 });
  }

  const songs = [];
  const sv = values_(SH_SONGS);
  for (let i = 1; i < sv.length; i++) {
    const r = sv[i], id = num_(r[0]);
    if (!id || !String(r[1]).trim() || isO_(r[17])) continue;
    songs.push({ id: id, title: clean_(r[1], 40), lyricist: clean_(r[2], 40), composer: clean_(r[3], 40), concert: clean_(r[4], 40),
      year: num_(r[5]) || '', grade: clean_(r[6], 10), themes: String(r[7]).split(/[,，、]/).map(function (x) { return clean_(x, 10); }).filter(String),
      season: clean_(r[8], 4), key: clean_(r[9], 4), beat: clean_(r[10], 6), range: clean_(r[11], 12),
      lyrics: String(r[12] || '').slice(0, 2000), intro: String(r[13] || '').slice(0, 600), tip: String(r[14] || '').slice(0, 600),
      video: /^https:\/\//.test(String(r[15])) ? String(r[15]).trim() : '', score: /^https:\/\//.test(String(r[16])) ? String(r[16]).trim() : '' });
  }

  const set = {}; values_(SH_SET).slice(1).forEach(function (r) { set[String(r[0]).trim()] = r[1]; });
  const relay = { id: clean_(set['릴레이 이름'], 30) || '첫 번째 릴레이', song: num_(set['릴레이 곡번호']) || (songs[0] ? songs[0].id : 0) };

  const out = { month: month, classes: classCount, regions: regions, songClasses: songClasses, votes: votes, relays: relays, lyrics: lyrics, songs: songs, relay: relay, updated: now_() };
  try { cache.put('board', JSON.stringify(out), BOARD_CACHE_SEC); } catch (e) {}
  return out;
}

// 시트를 고친 뒤 바로 보고 싶을 때 편집기에서 실행
function clearBoardCache() { CacheService.getScriptCache().remove('board'); }

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

/************ 운영진 메뉴 (노래상자 화면에서 운영진 비밀번호로) ************/
const OP_TTL_SEC = 21600;
const OP_FN_ = { opLogin: opLogin_, opLogout: opLogout_, opData: opData_, opSaveSong: opSaveSong_, opLyric: opLyric_, opRelay: opRelay_, opClass: opClass_, opPassword: opPassword_ };
function opHash_() { return PropertiesService.getScriptProperties().getProperty('OP_HASH') || hash_('1234'); }
function opLogin_(pw) {
  if (hash_(String(pw || '')) !== opHash_()) return { ok: false, message: '운영진 비밀번호가 맞지 않아요.' };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('OP_' + token, '1', OP_TTL_SEC);
  return { ok: true, token: token, defaultPw: !PropertiesService.getScriptProperties().getProperty('OP_HASH') };
}
function opLogout_(token) { if (token) CacheService.getScriptCache().remove('OP_' + token); return { ok: true }; }
function opNeed_(token) { if (!token || !CacheService.getScriptCache().get('OP_' + token)) throw new Error('운영진 로그인이 필요해요. 다시 로그인해 주세요.'); }
function opPassword_(token, pw) {
  opNeed_(token);
  const v = String(pw || '').trim();
  if (v.length < 6) return { ok: false, message: '운영진 비밀번호는 6자 이상으로 정해 주세요.' };
  PropertiesService.getScriptProperties().setProperty('OP_HASH', hash_(v));
  return { ok: true };
}
// (선택) 편집기에서 직접 정할 때: setOperatorPassword("새비밀번호")
function setOperatorPassword(pw) { PropertiesService.getScriptProperties().setProperty('OP_HASH', hash_(String(pw))); }

function opData_(token) {
  opNeed_(token);
  const songs = values_(SH_SONGS).slice(1).filter(function (r) { return num_(r[0]); }).map(function (r) {
    return { id: num_(r[0]), title: String(r[1]), lyricist: String(r[2]), composer: String(r[3]), concert: String(r[4]), year: r[5] === '' ? '' : String(r[5]),
      grade: String(r[6]), themes: String(r[7]), season: String(r[8]), key: String(r[9]), beat: String(r[10]), range: String(r[11]),
      lyrics: String(r[12]), intro: String(r[13]), tip: String(r[14]), video: String(r[15]), score: String(r[16]), hidden: isO_(r[17]) };
  });
  const lyrics = values_(SH_LYRICS).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const st = String(r[7] || '').trim();
    return { id: String(r[0]), songId: num_(r[2]), text: String(r[3]), cls: String(r[4]), region: String(r[5]), at: String(r[6]).slice(0, 16), status: isO_(st) ? 'O' : st ? 'X' : '' };
  });
  const classes = values_(SH_CLASSES).slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { pid: String(r[0]), name: String(r[1]), region: String(r[2]), stamps: parse_(r[3], []).length, last: String(r[8]).slice(0, 16), hidden: isO_(r[9]) };
  });
  const set = {}; values_(SH_SET).slice(1).forEach(function (r) { set[String(r[0]).trim()] = r[1]; });
  return { songs: songs, lyrics: lyrics, classes: classes, relay: { id: String(set['릴레이 이름'] || ''), song: num_(set['릴레이 곡번호']) },
    defaultPw: !PropertiesService.getScriptProperties().getProperty('OP_HASH') };
}

function opSaveSong_(token, s) {
  opNeed_(token);
  s = s || {};
  const title = clean_(s.title, 40);
  if (!title) throw new Error('곡 제목을 적어 주세요.');
  const link = function (u) { u = String(u || '').trim(); if (u && !/^https:\/\//.test(u)) throw new Error('영상·악보 주소는 https:// 로 시작해야 해요.'); return u.slice(0, 500); };
  const long = function (t, n) { return String(t == null ? '' : t).replace(/\r/g, '').replace(/[<>]/g, '').slice(0, n); };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH_SONGS), v = sh.getDataRange().getValues();
    let id = num_(s.id), row = 0, max = 0;
    for (let i = 1; i < v.length; i++) { const n = num_(v[i][0]); max = Math.max(max, n); if (id && n === id) row = i + 1; }
    if (!id) id = max + 1;
    const data = [id, title, clean_(s.lyricist, 40), clean_(s.composer, 40), clean_(s.concert, 40), num_(s.year) || '', clean_(s.grade, 10),
      clean_(s.themes, 60), clean_(s.season, 4), clean_(s.key, 4), clean_(s.beat, 6), clean_(s.range, 12),
      long(s.lyrics, 2000), long(s.intro, 600), long(s.tip, 600), link(s.video), link(s.score), s.hidden ? 'O' : ''];
    if (row) sh.getRange(row, 1, 1, data.length).setValues([data]);
    else { sh.appendRow(data); row = sh.getLastRow(); }
    sh.getRange(row, 13, 1, 3).setWrap(true);
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('board');
  return opData_(token);
}

function opLyric_(token, id, status) {
  opNeed_(token);
  const sh = sheet_(SH_LYRICS), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    sh.getRange(i + 1, 8).setValue(status === 'O' ? 'O' : status === 'X' ? 'X' : '');
    CacheService.getScriptCache().remove('board');
    return opData_(token);
  }
  throw new Error('그 가사를 찾지 못했어요.');
}

function opRelay_(token, name, songId) {
  opNeed_(token);
  const n = clean_(name, 30), id = num_(songId);
  if (!n || !id) throw new Error('릴레이 이름과 곡을 정해 주세요.');
  const sh = sheet_(SH_SET), v = sh.getDataRange().getValues();
  const put = function (k, val) {
    for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === k) { sh.getRange(i + 1, 2).setValue(val); return; }
    sh.appendRow([k, val, '']);
  };
  put('릴레이 이름', n); put('릴레이 곡번호', id);
  CacheService.getScriptCache().remove('board');
  return opData_(token);
}

function opClass_(token, pid, hidden) {
  opNeed_(token);
  const sh = sheet_(SH_CLASSES), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(pid)) {
    sh.getRange(i + 1, 10).setValue(hidden ? 'O' : '');
    CacheService.getScriptCache().remove('board');
    return opData_(token);
  }
  throw new Error('그 반을 찾지 못했어요.');
}
