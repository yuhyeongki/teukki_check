/**
 * 활동 자기평가 · 확인 시스템 (백엔드)
 * 구글 스프레드시트의 [확장 프로그램 → Apps Script]에 이 파일 전체를 붙여넣어 사용합니다.
 */

// ▼▼▼ 이 두 값만 바꾸세요 ▼▼▼
const CONFIG = {
  CLIENT_ID: '여기에_OAuth_클라이언트_ID.apps.googleusercontent.com',
  ADMIN_EMAIL: '선생님_구글계정@gmail.com',
};
// ▲▲▲ 여기까지 ▲▲▲

const SHEETS = {
  설정: ['key', 'value'],
  탭: ['id', 'name', 'guide', 'questions', 'groups', 'open', 'byteLimit', 'order', 'deleted'],
  명단: ['email', 'sid', 'name', 'groups', 'status', 'createdAt'],
  응답: ['tabId', 'email', 'sid', 'name', 'answers', 'status', 'updatedAt'],
  검토: ['tabId', 'email', 'sid', 'name', 'text', 'published', 'version', 'confirm', 'comment', 'confirmedAt'],
};

const DEFAULT_TEXTS = {
  siteTitle: '활동 자기평가',
  loginGuide: '본인의 구글 계정으로 로그인하세요.\n카카오톡 등 앱 안의 브라우저에서는 로그인이 되지 않으니 Chrome이나 Safari로 열어 주세요.',
  registerGuide: '처음 접속했습니다. 학번과 이름을 정확히 입력해 등록을 신청하세요.',
  pendingGuide: '등록 신청이 접수되었습니다. 선생님이 승인하면 이용할 수 있습니다.',
  writeGuide: '내가 실제로 한 활동, 맡은 역할, 배우고 느낀 점을 사실대로 적으세요.',
  resultGuide: '활동 내역이 사실과 다르거나 오탈자가 있으면 수정 의견을 남겨 주세요.',
  submitDone: '제출했습니다. 마감 전까지는 수정할 수 있습니다.',
  closedNotice: '작성이 마감되었습니다.',
  registrationOpen: '1',
};

const STATUSES = ['대기', '승인', '차단'];
const MAX_ANSWER = 3000;
const MAX_REVIEW = 5000;

/** 처음 한 번 편집기에서 실행하세요. 여러 번 실행해도 기존 데이터는 지워지지 않습니다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[name]);
      sh.setFrozenRows(1);
    }
  });
  const have = readTable_('설정').map(function (r) { return r.key; });
  Object.keys(DEFAULT_TEXTS).forEach(function (k) {
    if (have.indexOf(k) === -1) writeRow_('설정', { key: k, value: DEFAULT_TEXTS[k] });
  });
  if (readTable_('탭').length === 0) {
    writeRow_('탭', {
      id: newId_(), name: '예시 과목', guide: '이 안내와 문항은 설정 화면에서 바꿀 수 있습니다.',
      questions: JSON.stringify([
        { id: newId_(), text: '어떤 활동을 했나요?' },
        { id: newId_(), text: '맡은 역할과 배우고 느낀 점을 적어 주세요.' },
      ]),
      groups: '', open: '1', byteLimit: '1500', order: '1', deleted: '',
    });
  }
  // 권한 승인 화면이 한 번에 뜨도록 외부 요청도 호출해 둡니다.
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });
}

function doGet() {
  return json_({ ok: true, message: 'API 작동 중' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'public') {
      const t = texts_();
      return json_({ ok: true, data: { siteTitle: t.siteTitle, loginGuide: t.loginGuide } });
    }
    const user = verify_(req.token);
    const isAdmin = user.email === CONFIG.ADMIN_EMAIL.trim().toLowerCase();

    const studentActions = { boot: boot_, register: register_, saveResponse: saveResponse_, confirmResult: confirmResult_ };
    const adminActions = { adminData: adminData_, saveSettings: saveSettings_, saveStudents: saveStudents_, reviewData: reviewData_, saveReviews: saveReviews_ };
    const writes = ['register', 'saveResponse', 'confirmResult', 'saveSettings', 'saveStudents', 'saveReviews'];

    let fn;
    if (Object.prototype.hasOwnProperty.call(studentActions, req.action)) fn = studentActions[req.action];
    else if (Object.prototype.hasOwnProperty.call(adminActions, req.action)) {
      if (!isAdmin) fail_('FORBIDDEN', '교사 계정만 사용할 수 있습니다.');
      fn = adminActions[req.action];
    } else fail_('BAD', '알 수 없는 요청입니다.');

    const run = function () { return fn(req, user, isAdmin); };
    const data = writes.indexOf(req.action) === -1 ? run() : withLock_(run);
    return json_({ ok: true, data: data });
  } catch (err) {
    if (err && err.appCode) return json_({ ok: false, code: err.appCode, message: err.message });
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, code: 'SERVER', message: '서버 오류가 발생했습니다.' });
  }
}

/* ───────── 학생 ───────── */

function boot_(req, user, isAdmin) {
  const t = texts_();
  if (isAdmin) return { role: 'admin', email: user.email, texts: t };
  const me = findStudent_(user.email);
  if (!me) return { role: 'unregistered', email: user.email, texts: t };
  if (me.status === '차단') return { role: 'blocked', email: user.email, texts: t };
  if (me.status !== '승인') return { role: 'pending', email: user.email, texts: t };

  const mine = function (r) { return r.email.toLowerCase() === user.email; };
  const responses = readTable_('응답').filter(mine);
  const reviews = readTable_('검토').filter(mine);
  const tabs = tabs_().filter(function (tab) { return visible_(tab, me); }).map(function (tab) {
    const res = responses.filter(function (r) { return r.tabId === tab.id; })[0];
    const rev = reviews.filter(function (r) { return r.tabId === tab.id; })[0];
    const saved = res ? parse_(res.answers, {}) : {};
    const answers = {};
    tab.questions.forEach(function (q) { answers[q.id] = saved[q.id] ? saved[q.id].a : ''; });
    return {
      id: tab.id, name: tab.name, guide: tab.guide, questions: tab.questions, open: tab.open,
      answers: answers, status: res ? res.status : '', updatedAt: res ? res.updatedAt : '',
      result: rev && rev.published === '1' && rev.text
        ? { text: rev.text, version: rev.version, confirm: rev.confirm, comment: rev.comment }
        : null,
    };
  });
  return { role: 'student', email: user.email, sid: me.sid, name: me.name, texts: t, tabs: tabs };
}

function register_(req, user, isAdmin) {
  if (isAdmin) fail_('BAD', '교사 계정은 등록할 필요가 없습니다.');
  if (findStudent_(user.email)) return { done: true };
  if (texts_().registrationOpen !== '1') fail_('CLOSED', '지금은 등록 신청을 받지 않습니다.');
  const sid = clean_(req.sid, 20);
  const name = clean_(req.name, 30);
  if (!sid || !name) fail_('BAD', '학번과 이름을 모두 입력하세요.');
  writeRow_('명단', { email: user.email, sid: sid, name: name, groups: '', status: '대기', createdAt: now_() });
  return { done: true };
}

function saveResponse_(req, user) {
  const me = approved_(user.email);
  const tab = tabs_().filter(function (t) { return t.id === req.tabId; })[0];
  if (!tab || !visible_(tab, me)) fail_('FORBIDDEN', '작성할 수 없는 항목입니다.');
  if (!tab.open) fail_('CLOSED', texts_().closedNotice);

  const prev = findRow_('응답', tab.id, user.email);
  const answers = prev ? parse_(prev.answers, {}) : {}; // 삭제된 문항의 답변도 보존
  const input = req.answers || {};
  let filled = false;
  tab.questions.forEach(function (q) {
    const a = String(input[q.id] == null ? '' : input[q.id]).replace(/\r/g, '');
    if (a.length > MAX_ANSWER) fail_('BAD', '답변은 문항당 ' + MAX_ANSWER + '자까지 입력할 수 있습니다.');
    if (a.trim()) filled = true;
    answers[q.id] = { q: q.text, a: a };
  });
  const status = req.submit || (prev && prev.status === '제출') ? '제출' : '임시';
  if (status === '제출' && !filled) fail_('BAD', '내용을 입력한 뒤 제출하세요.');
  const packed = JSON.stringify(answers);
  if (packed.length > 45000) fail_('BAD', '전체 답변이 너무 깁니다. 내용을 줄여 주세요.');
  const row = { tabId: tab.id, email: user.email, sid: me.sid, name: me.name, answers: packed, status: status, updatedAt: now_() };
  writeRow_('응답', row, prev && prev._row);
  return { status: row.status, updatedAt: row.updatedAt };
}

function confirmResult_(req, user) {
  const me = approved_(user.email);
  const tab = tabs_().filter(function (t) { return t.id === req.tabId; })[0];
  const rev = tab && visible_(tab, me) ? findRow_('검토', tab.id, user.email) : null;
  if (!rev || rev.published !== '1' || !rev.text) fail_('FORBIDDEN', '확인할 내용이 없습니다.');
  if (rev.version !== String(req.version)) fail_('STALE', '내용이 바뀌었습니다. 새로고침한 뒤 다시 확인하세요.');
  if (req.confirm !== '이상 없음' && req.confirm !== '수정 의견') fail_('BAD', '확인 결과를 선택하세요.');
  const comment = req.confirm === '수정 의견' ? clean_(req.comment, 1000) : '';
  if (req.confirm === '수정 의견' && !comment) fail_('BAD', '수정이 필요한 부분을 적어 주세요.');
  rev.confirm = req.confirm;
  rev.comment = comment;
  rev.confirmedAt = now_();
  writeRow_('검토', rev, rev._row);
  return { confirm: rev.confirm, comment: rev.comment };
}

/* ───────── 교사 ───────── */

function adminData_() {
  return {
    texts: texts_(),
    tabs: tabs_(),
    students: readTable_('명단').map(function (s) {
      return { email: s.email.toLowerCase(), sid: s.sid, name: s.name, groups: s.groups, status: s.status };
    }),
  };
}

function saveSettings_(req) {
  const incoming = req.tabs || [];
  incoming.forEach(function (t) {
    if (!t.deleted && !clean_(t.name, 50)) fail_('BAD', '탭 이름을 입력하세요.');
  });

  const rows = readTable_('설정');
  Object.keys(DEFAULT_TEXTS).forEach(function (k) {
    if (!req.texts || req.texts[k] == null) return;
    const row = rows.filter(function (r) { return r.key === k; })[0];
    writeRow_('설정', { key: k, value: String(req.texts[k]).slice(0, 3000) }, row && row._row);
  });

  incoming.forEach(function (t, i) {
    const prev = t.id ? readTable_('탭').filter(function (r) { return r.id === t.id; })[0] : null;
    if (!prev && t.deleted) return;
    const questions = (t.questions || []).map(function (q) {
      return { id: validId_(q.id), text: clean_(q.text, 500) };
    }).filter(function (q) { return q.text; }).slice(0, 12);
    writeRow_('탭', {
      id: prev ? prev.id : validId_(t.id),
      name: clean_(t.name, 50),
      guide: String(t.guide || '').slice(0, 3000),
      questions: JSON.stringify(questions),
      groups: splitGroups_(t.groups).join(', '),
      open: t.open ? '1' : '',
      byteLimit: String(Math.max(0, Math.min(10000, parseInt(t.byteLimit, 10) || 0))),
      order: String(i + 1),
      deleted: t.deleted ? '1' : '',
    }, prev && prev._row);
  });
  return adminData_();
}

function saveStudents_(req) {
  const known = byEmail_(readTable_('명단'));
  (req.students || []).forEach(function (s) {
    const email = String(s.email || '').trim().toLowerCase();
    if (!email) return;
    if (STATUSES.indexOf(s.status) === -1) fail_('BAD', '상태 값이 올바르지 않습니다.');
    const prev = known[email];
    writeRow_('명단', {
      email: email, sid: clean_(s.sid, 20), name: clean_(s.name, 30),
      groups: splitGroups_(s.groups).join(', '), status: s.status,
      createdAt: prev ? prev.createdAt : now_(),
    }, prev && prev._row);
  });
  return adminData_();
}

function reviewData_(req) {
  const tab = tabs_().filter(function (t) { return t.id === req.tabId; })[0];
  if (!tab) fail_('BAD', '탭을 찾을 수 없습니다.');
  const inTab = function (r) { return r.tabId === tab.id; };
  const responses = byEmail_(readTable_('응답').filter(inTab));
  const reviews = byEmail_(readTable_('검토').filter(inTab));
  const list = readTable_('명단').filter(function (s) {
    const email = s.email.toLowerCase();
    return (s.status === '승인' && visible_(tab, s)) || responses[email] || reviews[email];
  }).map(function (s) {
    const email = s.email.toLowerCase();
    const res = responses[email];
    const rev = reviews[email];
    const saved = res ? parse_(res.answers, {}) : {};
    const order = tab.questions.map(function (q) { return q.id; });
    Object.keys(saved).forEach(function (id) { if (order.indexOf(id) === -1) order.push(id); });
    return {
      email: email, sid: s.sid, name: s.name,
      status: res ? res.status : '미작성',
      answers: order.filter(function (id) { return saved[id]; }).map(function (id) { return saved[id]; }),
      text: rev ? rev.text : '', published: !!rev && rev.published === '1',
      confirm: rev ? rev.confirm : '', comment: rev ? rev.comment : '',
    };
  });
  list.sort(function (a, b) { return a.sid.localeCompare(b.sid, 'ko', { numeric: true }); });
  return { tabId: tab.id, students: list };
}

function saveReviews_(req) {
  const tab = tabs_().filter(function (t) { return t.id === req.tabId; })[0];
  if (!tab) fail_('BAD', '탭을 찾을 수 없습니다.');
  const items = req.items || [];
  // 하나라도 문제가 있으면 아무것도 저장하지 않도록 먼저 검사
  const students = byEmail_(readTable_('명단'));
  const existing = byEmail_(readTable_('검토').filter(function (r) { return r.tabId === tab.id; }));
  items.forEach(function (it) {
    if (String(it.text || '').length > MAX_REVIEW) fail_('BAD', '내용이 너무 깁니다.');
    if (!students[String(it.email || '').toLowerCase()]) fail_('BAD', '명단에 없는 학생이 있습니다.');
  });
  items.forEach(function (it) {
    const email = String(it.email).toLowerCase();
    const s = students[email];
    const prev = existing[email];
    const text = String(it.text || '').replace(/\r/g, '');
    const changed = !prev || prev.text !== text;
    writeRow_('검토', {
      tabId: tab.id, email: email, sid: s.sid, name: s.name, text: text,
      published: it.published && text.trim() ? '1' : '',
      version: changed ? newId_() : prev.version,
      // 문구가 바뀌면 이전 확인 결과는 초기화
      confirm: changed ? '' : prev.confirm,
      comment: changed ? '' : prev.comment,
      confirmedAt: changed ? '' : prev.confirmedAt,
    }, prev && prev._row);
  });
  return reviewData_(req);
}

/* ───────── 공통 ───────── */

/** 구글이 발급한 로그인 토큰인지 서버에서 확인합니다. */
function verify_(token) {
  if (!token || typeof token !== 'string') fail_('AUTH', '로그인이 필요합니다.');
  const cache = CacheService.getScriptCache();
  const key = 'tk_' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const nowSec = Math.floor(Date.now() / 1000);
  const hit = cache.get(key);
  if (hit) {
    const u = JSON.parse(hit);
    if (u.exp > nowSec) return u;
  }
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) fail_('AUTH', '로그인이 만료되었습니다. 다시 로그인하세요.');
  const p = JSON.parse(res.getContentText());
  const issuerOk = p.iss === 'accounts.google.com' || p.iss === 'https://accounts.google.com';
  if (p.aud !== CONFIG.CLIENT_ID || !issuerOk || String(p.email_verified) !== 'true' || !p.email || !(Number(p.exp) > nowSec)) {
    fail_('AUTH', '로그인 정보를 확인할 수 없습니다.');
  }
  const user = { email: String(p.email).toLowerCase(), exp: Number(p.exp) };
  const ttl = Math.min(21600, user.exp - nowSec);
  if (ttl > 0) cache.put(key, JSON.stringify(user), ttl);
  return user;
}

function texts_() {
  const out = {};
  Object.keys(DEFAULT_TEXTS).forEach(function (k) { out[k] = DEFAULT_TEXTS[k]; });
  readTable_('설정').forEach(function (r) { if (r.key in out) out[r.key] = r.value; });
  return out;
}

function tabs_() {
  return readTable_('탭')
    .filter(function (r) { return r.id && r.deleted !== '1'; })
    .sort(function (a, b) { return Number(a.order) - Number(b.order); })
    .map(function (r) {
      return {
        id: r.id, name: r.name, guide: r.guide, questions: parse_(r.questions, []),
        groups: r.groups, open: r.open === '1', byteLimit: Number(r.byteLimit) || 0,
      };
    });
}

function splitGroups_(s) {
  return String(s || '').split(/[,，]/).map(function (x) { return x.trim(); }).filter(String);
}

/** 탭에 대상 그룹이 없으면 전체 공개, 있으면 그룹이 하나라도 겹치는 학생만 */
function visible_(tab, student) {
  const need = splitGroups_(tab.groups);
  if (!need.length) return true;
  const have = splitGroups_(student.groups);
  return need.some(function (g) { return have.indexOf(g) !== -1; });
}

function findStudent_(email) {
  return readTable_('명단').filter(function (s) { return s.email.trim().toLowerCase() === email; })[0] || null;
}

function approved_(email) {
  const s = findStudent_(email);
  if (!s || s.status !== '승인') fail_('FORBIDDEN', '승인된 학생만 사용할 수 있습니다.');
  return s;
}

function findRow_(sheet, tabId, email) {
  return readTable_(sheet).filter(function (r) { return r.tabId === tabId && r.email.toLowerCase() === email; })[0] || null;
}

function byEmail_(rows) {
  const m = Object.create(null);
  rows.forEach(function (r) { m[r.email.toLowerCase()] = r; });
  return m;
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) fail_('SETUP', '시트가 없습니다. 편집기에서 setup 함수를 먼저 실행하세요.');
  return sh;
}

function readTable_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const head = values[0];
  return values.slice(1).map(function (row, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, c) { o[h] = String(row[c] == null ? '' : row[c]); });
    return o;
  });
}

/** 모든 값을 앞에 ' 를 붙여 글자로 저장합니다. 학생이 입력한 내용이 수식으로 실행되는 것을 막습니다. */
function writeRow_(name, obj, rowNumber) {
  const sh = sheet_(name);
  const values = SHEETS[name].map(function (h) {
    const v = String(obj[h] == null ? '' : obj[h]);
    return v === '' ? '' : "'" + v;
  });
  if (rowNumber) sh.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values); // 행이 모자라면 자동으로 늘어남
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_('BUSY', '접속이 몰리고 있습니다. 잠시 후 다시 시도하세요.');
  try {
    const out = fn();
    SpreadsheetApp.flush();
    return out;
  } finally {
    lock.releaseLock();
  }
}

function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function parse_(s, fallback) { try { return JSON.parse(s) || fallback; } catch (e) { return fallback; } }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
/** 화면에서 만든 ID를 그대로 쓰면 저장 요청이 재시도되어도 탭이 중복 생성되지 않습니다. */
function validId_(id) { return /^[0-9a-f]{12}$/.test(id) ? id : newId_(); }
function fail_(code, message) { const e = new Error(message); e.appCode = code; throw e; }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
