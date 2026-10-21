/**
 * filego — 교사 파일 취합 백엔드 (Google Apps Script)
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기
 *       → 함수 'setup' 한 번 실행 → 배포(웹 앱, 실행: 나, 액세스: 모든 사용자)
 * 자세한 방법은 설정방법.md 참고
 */

// ★ 관리자 비밀번호 — 바꾼 뒤에는 '배포 관리 → 새 버전'으로 다시 배포해야 적용됩니다
const ADMIN_PASSWORD = '1234';

// 드라이브 최상위 폴더 이름
const ROOT_FOLDER_NAME = 'filego 취합';

const SHEET_COLLECTIONS = '취합목록';
const SHEET_ROSTER = '교사명단';
const SHEET_SUBMISSIONS = '제출기록';

const HEADERS = {
  [SHEET_COLLECTIONS]: ['ID', '제목', '안내문', '마감일', '상태', '폴더ID', '생성일', '대상명단', '대상방식'],
  [SHEET_ROSTER]: ['이름', '소속'],
  [SHEET_SUBMISSIONS]: ['제출시각', '취합ID', '이름', '파일명', '파일ID'],
};

const STATUS_OPEN = '진행';
const STATUS_CLOSED = '마감';

// 대상방식: 전체 = 교사 명단 전체(인사발령 즉시 반영), 선택 = 취합마다 고른 사람만
const MODE_ALL = '전체';
const MODE_PICK = '선택';

// ───────────────────────── 초기 설정 ─────────────────────────

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  // 날짜 자동 변환 방지
  ss.getSheetByName(SHEET_COLLECTIONS).getRange('D:D').setNumberFormat('@');
  const def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  getRootFolder_();
  Logger.log('설정 완료. 이제 배포 → 새 배포 → 웹 앱으로 배포하세요.');
}

// ───────────────────────── 요청 처리 ─────────────────────────

function doGet() {
  return json_({ ok: true, msg: 'filego 서버가 동작 중입니다.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '잘못된 요청입니다.' });
  }

  const publicActions = { list, check, upload };
  const adminActions = {
    adminLogin: () => ({}),
    adminList, adminDetail, create, update, remove, setStatus,
    getRoster, saveRoster, zip, removeSubmission, getFile,
  };

  try {
    if (publicActions[req.action]) return json_({ ok: true, ...publicActions[req.action](req) });
    if (adminActions[req.action]) {
      if (req.password !== ADMIN_PASSWORD) return json_({ ok: false, error: '비밀번호가 맞지 않습니다.' });
      return json_({ ok: true, ...adminActions[req.action](req) });
    }
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

// ───────────────────────── 선생님용 ─────────────────────────

/** 진행 중인 취합 목록 + 교사 명단 (제출 현황은 보내지 않음) */
function list() {
  const roster = readRoster_();
  const rosterNames = roster.map(r => r.name);
  const collections = readCollections_()
    .filter(c => c.status === STATUS_OPEN)
    .map(c => ({
      id: c.id, title: c.title, desc: c.desc, deadline: c.deadline,
      targets: c.mode === MODE_PICK ? targetsOf_(c, rosterNames) : null, // null이면 전체 명단
    }));
  return { collections, roster };
}

/** 이름을 고른 사람이 이미 냈는지만 확인 */
function check(req) {
  return { submitted: !!(submittedMap_()[req.cid] || {})[req.name] };
}

/**
 * 파일 1개 업로드. 여러 파일은 순서대로 한 번씩 호출.
 * first=true 인 요청에서 그 선생님의 이전 제출 파일을 정리(교체)합니다.
 */
function upload(req) {
  const { cid, name, filename, mime, data, first, index, total } = req;
  if (!cid || !name || !filename || !data) throw new Error('필수 정보가 빠졌습니다.');

  const c = findCollection_(cid);
  if (!c) throw new Error('존재하지 않는 취합입니다.');
  if (c.status !== STATUS_OPEN) throw new Error('이미 마감된 취합입니다.');
  if (!targetsOf_(c).includes(name)) throw new Error('이번 취합 대상이 아닙니다. 관리자에게 문의해주세요.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folder = DriveApp.getFolderById(c.folderId);
    const sh = sheet_(SHEET_SUBMISSIONS);

    if (first) removeSubmissions_(cid, name);

    const ext = (filename.match(/\.[^./\\]+$/) || [''])[0];
    const suffix = total > 1 ? '_' + index : '';
    const newName = safeName_(name) + '_' + safeName_(c.title) + suffix + ext;

    const blob = Utilities.newBlob(Utilities.base64Decode(data), mime || 'application/octet-stream', newName);
    const file = folder.createFile(blob);
    sh.appendRow([new Date(), cid, name, newName, file.getId()]);
    return { file: newName };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 관리자용 ─────────────────────────

function adminList() {
  const submitted = submittedMap_();
  const rosterNames = readRoster_().map(r => r.name);
  const collections = readCollections_().map(c => {
    const names = Object.keys(submitted[c.id] || {});
    const targets = targetsOf_(c, rosterNames);
    return {
      id: c.id, title: c.title, desc: c.desc, deadline: c.deadline, status: c.status, created: c.created,
      mode: c.mode, targets: c.mode === MODE_PICK ? targets : null, // '이전 취합 대상 불러오기'용
      submittedCount: names.filter(n => targets.includes(n)).length,
      targetCount: targets.length,
    };
  }).reverse(); // 최신순
  return { collections };
}

function adminDetail(req) {
  const c = mustCollection_(req.cid);
  const subs = submittedMap_()[c.id] || {};
  const targets = targetsOf_(c);
  const rosterDept = {};
  readRoster_().forEach(r => rosterDept[r.name] = r.dept);

  const submitted = Object.keys(subs).map(n => ({
    name: n,
    dept: rosterDept[n] || '',
    outside: !targets.includes(n),
    order: targets.includes(n) ? targets.indexOf(n) : targets.length, // 명단 순서 (PDF 합치기용)
    time: subs[n].time,
    files: subs[n].files,
  })).sort((a, b) => b.time - a.time);

  const missing = targets.filter(n => !subs[n]).map(n => ({ name: n, dept: rosterDept[n] || '' }));

  return {
    collection: {
      id: c.id, title: c.title, desc: c.desc, deadline: c.deadline, status: c.status,
      mode: c.mode, targets: c.mode === MODE_PICK ? targets : null,
      folderUrl: 'https://drive.google.com/drive/folders/' + c.folderId,
    },
    submitted, missing, targetCount: targets.length,
  };
}

function create(req) {
  const title = String(req.title || '').trim();
  if (!title) throw new Error('제목을 입력해주세요.');
  const id = 'c' + Date.now().toString(36);
  const month = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
  const t = parseTargets_(req);
  const folder = getRootFolder_().createFolder(month + ' ' + safeName_(title));
  const sh = sheet_(SHEET_COLLECTIONS);
  ensureHeader_(sh, SHEET_COLLECTIONS);
  sh.appendRow([
    id, title, String(req.desc || ''), String(req.deadline || ''), STATUS_OPEN, folder.getId(), new Date(),
    t.mode === MODE_PICK ? JSON.stringify(t.targets) : '', t.mode,
  ]);
  return { id };
}

function update(req) {
  const c = mustCollection_(req.cid);
  const sh = sheet_(SHEET_COLLECTIONS);
  const title = String(req.title || '').trim();
  if (!title) throw new Error('제목을 입력해주세요.');
  sh.getRange(c.row, 2, 1, 3).setValues([[title, String(req.desc || ''), String(req.deadline || '')]]);
  if (req.mode) {
    const t = parseTargets_(req);
    let saved = '';
    if (t.mode === MODE_PICK) saved = JSON.stringify(t.targets);
    else if (c.status === STATUS_CLOSED) saved = JSON.stringify(readRoster_().map(r => r.name)); // 마감된 전체 취합은 현재 명단으로 고정
    ensureHeader_(sh, SHEET_COLLECTIONS);
    sh.getRange(c.row, 8, 1, 2).setValues([[saved, t.mode]]);
  }
  return {};
}

/** 목록에서만 삭제합니다. 드라이브 폴더와 파일은 남겨둡니다. */
function remove(req) {
  const c = mustCollection_(req.cid);
  sheet_(SHEET_COLLECTIONS).deleteRow(c.row);
  return {};
}

/** 한 사람의 제출을 취소합니다. 파일은 휴지통으로 옮겨져 30일 안에는 복구할 수 있습니다. */
function removeSubmission(req) {
  mustCollection_(req.cid);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    removeSubmissions_(req.cid, req.name);
  } finally {
    lock.releaseLock();
  }
  return {};
}

/** 마감하면 그 시점의 명단을 저장해 두어, 이후 인사이동이 있어도 당시 대상자 기준으로 집계됩니다. */
function setStatus(req) {
  const c = mustCollection_(req.cid);
  const sh = sheet_(SHEET_COLLECTIONS);
  // 직접 선택한 취합은 대상명단을 그대로 두고, 전체 취합만 마감 시 명단을 저장/재개 시 비움
  if (req.status === STATUS_CLOSED) {
    sh.getRange(c.row, 5).setValue(STATUS_CLOSED);
    if (c.mode !== MODE_PICK) sh.getRange(c.row, 8).setValue(JSON.stringify(readRoster_().map(r => r.name)));
  } else {
    sh.getRange(c.row, 5).setValue(STATUS_OPEN);
    if (c.mode !== MODE_PICK) sh.getRange(c.row, 8).setValue('');
  }
  return {};
}

function getRoster() {
  return { roster: readRoster_() };
}

/** 명단 전체를 받은 내용으로 교체합니다. */
function saveRoster(req) {
  const rows = (req.roster || [])
    .map(r => ({ name: String(r.name || '').trim(), dept: String(r.dept || '').trim() }))
    .filter(r => r.name);
  const seen = {};
  rows.forEach(r => {
    if (seen[r.name]) throw new Error(`'${r.name}' 이름이 두 번 있습니다. 동명이인은 '김민지(3-1)'처럼 구분해주세요.`);
    seen[r.name] = true;
  });
  const sh = sheet_(SHEET_ROSTER);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows.map(r => [r.name, r.dept]));
  return { count: rows.length };
}

/** 취합 폴더의 파일을 ZIP으로 묶어 돌려줍니다 (약 40MB 이하). */
function zip(req) {
  const c = mustCollection_(req.cid);
  const ids = {};
  readSubmissions_().filter(s => s.cid === c.id).forEach(s => ids[s.fileId] = true);
  const blobs = [];
  let size = 0;
  Object.keys(ids).forEach(id => {
    try {
      const f = DriveApp.getFileById(id);
      if (f.isTrashed()) return;
      size += f.getSize();
      if (size > 40 * 1024 * 1024) throw new Error('TOO_BIG');
      blobs.push(f.getBlob());
    } catch (err) {
      if (err.message === 'TOO_BIG') throw new Error('파일이 너무 커서 ZIP으로 묶을 수 없습니다. 「드라이브 폴더 열기」에서 폴더째 다운로드해주세요.');
    }
  });
  if (!blobs.length) throw new Error('제출된 파일이 없습니다.');
  const zipName = safeName_(c.title) + '.zip';
  const z = Utilities.zip(blobs, zipName);
  return { filename: zipName, data: Utilities.base64Encode(z.getBytes()) };
}

/** 제출된 파일 1개를 돌려줍니다 (관리자 화면에서 PDF 합치기용). */
function getFile(req) {
  const c = mustCollection_(req.cid);
  if (!readSubmissions_().some(s => s.cid === c.id && s.fileId === req.fileId)) throw new Error('이 취합에 제출된 파일이 아닙니다.');
  const f = DriveApp.getFileById(req.fileId);
  const blob = f.getBlob();
  return { name: f.getName(), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

// ───────────────────────── 내부 함수 ─────────────────────────

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(`'${name}' 시트가 없습니다. setup()을 먼저 실행해주세요.`);
  return sh;
}

function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS[name].length).getValues();
}

function readCollections_() {
  return rows_(SHEET_COLLECTIONS).map((r, i) => ({
    row: i + 2,
    id: String(r[0]), title: String(r[1]), desc: String(r[2]),
    deadline: r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[3]),
    status: String(r[4]) || STATUS_OPEN, folderId: String(r[5]),
    created: r[6] instanceof Date ? r[6].getTime() : 0,
    snapshot: String(r[7] || ''),
    mode: String(r[8]) === MODE_PICK ? MODE_PICK : MODE_ALL,
  })).filter(c => c.id);
}

function findCollection_(cid) {
  return readCollections_().find(c => c.id === cid);
}

function mustCollection_(cid) {
  const c = findCollection_(cid);
  if (!c) throw new Error('존재하지 않는 취합입니다.');
  return c;
}

function readRoster_() {
  return rows_(SHEET_ROSTER)
    .map(r => ({ name: String(r[0]).trim(), dept: String(r[1]).trim() }))
    .filter(r => r.name);
}

function readSubmissions_() {
  return rows_(SHEET_SUBMISSIONS).map((r, i) => ({
    row: i + 2,
    time: r[0] instanceof Date ? r[0].getTime() : 0,
    cid: String(r[1]), name: String(r[2]), file: String(r[3]), fileId: String(r[4]),
  }));
}

/** { 취합ID: { 이름: { time, files:[{name,url}] } } } */
function submittedMap_() {
  const map = {};
  readSubmissions_().forEach(s => {
    const byName = map[s.cid] = map[s.cid] || {};
    const entry = byName[s.name] = byName[s.name] || { time: 0, files: [] };
    entry.time = Math.max(entry.time, s.time);
    entry.files.push({ id: s.fileId, name: s.file, url: 'https://drive.google.com/file/d/' + s.fileId + '/view' });
  });
  return map;
}

/**
 * 취합 대상자 이름 목록
 * - 직접 선택: 고른 사람 (명단 순서대로, 명단에서 빠진 사람은 뒤에)
 * - 전체: 진행 중이면 현재 명단, 마감됐으면 마감 당시 명단
 */
function targetsOf_(c, rosterNames) {
  const roster = rosterNames || readRoster_().map(r => r.name);
  let saved = null;
  if (c.snapshot) {
    try { saved = JSON.parse(c.snapshot); } catch (err) { /* 무시 */ }
  }
  if (c.mode === MODE_PICK && saved) {
    return roster.filter(n => saved.includes(n)).concat(saved.filter(n => !roster.includes(n)));
  }
  if (c.status === STATUS_CLOSED && saved) return saved;
  return roster;
}

/** 요청의 대상방식/대상자를 검사해서 돌려줍니다. */
function parseTargets_(req) {
  if (req.mode !== MODE_PICK) return { mode: MODE_ALL, targets: null };
  const targets = [...new Set((req.targets || []).map(n => String(n).trim()).filter(Boolean))];
  if (!targets.length) throw new Error('대상자를 한 명 이상 선택해주세요.');
  return { mode: MODE_PICK, targets };
}

/** 예전에 만든 시트에 새 열 제목이 없으면 채워 넣습니다. */
function ensureHeader_(sh, name) {
  const h = HEADERS[name];
  const cur = sh.getRange(1, 1, 1, h.length).getValues()[0];
  if (cur.some((v, i) => v !== h[i])) sh.getRange(1, 1, 1, h.length).setValues([h]);
}

/** 재제출 시 이전 파일을 휴지통으로 옮기고 기록을 지웁니다. */
function removeSubmissions_(cid, name) {
  const sh = sheet_(SHEET_SUBMISSIONS);
  readSubmissions_()
    .filter(s => s.cid === cid && s.name === name)
    .reverse() // 아래 행부터 지워야 행 번호가 밀리지 않음
    .forEach(s => {
      try { DriveApp.getFileById(s.fileId).setTrashed(true); } catch (err) { /* 이미 없는 파일 */ }
      sh.deleteRow(s.row);
    });
}

function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (err) { /* 새로 만듦 */ }
  }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function safeName_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_').trim();
}
