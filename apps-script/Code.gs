/**
 * 모바일 청첩장 백엔드 — Google Apps Script 웹앱
 * ------------------------------------------------------------------
 * 정적 사이트(GitHub Pages / EdgeOne)에는 자체 서버가 없으므로, 이 웹앱이
 * 백엔드 API 역할을 한다. 텍스트는 구글 스프레드시트에, 사진/썸네일은
 * 구글 드라이브에 저장한다.
 *
 *   GET  ?action=get     → 시트 텍스트 + 드라이브 미디어 URL(JSON)
 *   POST {action:'login'}          → 관리자 비밀번호 검증
 *   POST {action:'saveText'}       → 텍스트를 시트에 저장
 *   POST {action:'uploadMedia'}    → 사진/썸네일을 드라이브에 저장
 *   POST {action:'setCover'}       → 갤러리 사진 중 대표사진 지정
 *   POST {action:'deleteGallery'}  → 갤러리 사진 삭제(드라이브 파일도 삭제)
 *   POST {action:'reorderGallery'} → 갤러리 순서 변경
 *
 * ▶ 배포: 편집기 우측 상단 [배포] → [새 배포] → 유형 '웹 앱'
 *         - 실행: '나(소유자)'
 *         - 액세스 권한: '모든 사용자'
 *   생성된 '웹 앱 URL(/exec)'을 index.html 의 GAS_ENDPOINT 에 붙여넣는다.
 *
 * ▶ 비밀번호 변경(선택): [프로젝트 설정] → [스크립트 속성]에
 *   ADMIN_PASSWORD 키로 원하는 값을 등록(없으면 기본값 1010).
 *
 * ▶ 썸네일 교체 기능을 쓰려면(선택): [서비스] → 'Drive API'(고급 서비스) 추가.
 *   추가하지 않아도 나머지 기능은 모두 동작한다.
 */

/* ========================= 설정 ========================= */
var SHEET_ID   = '1M2BlLdvKB8FRy9xY6IURZxrwYc657UfqrpG3BRsk8ao';   // 모바일청첩장_DB
var FOLDER_ID  = '1pClChKU02rh5JKMYY6hL7GuVrUOp5KSP';              // 드라이브 폴더(사진·썸네일)
var THUMB_FILE_ID = '1GyZu3lYn0Rwsiw2_cKVHznHEOaLWrjfy';           // 공유 썸네일(og:image) 고정 파일
var DEFAULT_ADMIN_PASSWORD = '1010';

var TAB_INFO = '정보';
var TAB_ACC  = '계좌';
var TAB_GAL  = '갤러리';

/* 시트가 비어 있을 때 심는 기본 텍스트(index.html의 임베드 기본값과 동일) */
var SEED = {
  greeting: '저희 두 사람이 사랑의 결실을 맺어\n인생의 새로운 시작을 함께하려 합니다.\n\n귀한 걸음 하시어 축복해 주시면\n더없는 기쁨으로 간직하겠습니다.',
  groom_name: '박윤혁', groom_order: '차남', groom_father: '박경택', groom_mother: '우희란',
  bride_name: '임미숙', bride_order: '막내딸', bride_father: '임종철', bride_mother: '현애화',
  weddingDateTime: '2026-10-10T11:30',
  venueName: '서귀포농협 하나로마트 3층 웨딩홀',
  venueAddress: '제주 서귀포시 516로 40-1, 3층',
  transport_car: '예식장 내 주차장을 이용하실 수 있습니다. 하나로마트 이용 고객과 주차 공간을 함께 사용하므로 예식 시간대에는 다소 혼잡할 수 있으니 여유 있게 도착해 주세요.',
  transport_transit: "제주 시내버스 이용 시 '비석거리' 정류장에서 하차 후 도보로 이동하실 수 있습니다. 정확한 노선과 소요 시간은 카카오맵·네이버지도의 대중교통 길찾기를 이용해 주세요.",
  transport_airport: '제주국제공항에서 예식장까지 차량으로 약 1시간~1시간 10분 정도 소요됩니다(교통 상황에 따라 변동). 공항에서 서귀포 방면 리무진버스 또는 시외버스로 이동하신 후 택시로 환승하시는 방법을 추천드립니다.',
  coverFileId: '',
  dataVersion: '1'
};
var SEED_ACCOUNTS = [
  { side: 'groom', label: '신랑 박윤혁', bank: '농협', holder: '박윤혁', number: '302-0621-7625-31' }
];

/* ========================= 공통 헬퍼 ========================= */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.openById(SHEET_ID); }

function adminPassword() {
  var p = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return p || DEFAULT_ADMIN_PASSWORD;
}

function checkPw(pw) {
  return typeof pw === 'string' && pw === adminPassword();
}

/* 드라이브 파일 ID → 브라우저에서 안정적으로 표시되는 이미지 URL */
function driveImageUrl(fileId) {
  return 'https://lh3.googleusercontent.com/d/' + fileId + '=w1600';
}

/* 필요한 탭이 없으면 생성하고, 비어 있으면 기본값을 심는다 */
function ensureSheets() {
  var book = ss();

  var info = book.getSheetByName(TAB_INFO);
  if (!info) {
    info = book.insertSheet(TAB_INFO);
    info.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }
  if (info.getLastRow() < 2) {
    var rows = [];
    for (var k in SEED) { if (SEED.hasOwnProperty(k)) rows.push([k, SEED[k]]); }
    info.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  var acc = book.getSheetByName(TAB_ACC);
  if (!acc) {
    acc = book.insertSheet(TAB_ACC);
    acc.getRange(1, 1, 1, 5).setValues([['side', 'label', 'bank', 'holder', 'number']]);
    if (SEED_ACCOUNTS.length) {
      var arows = SEED_ACCOUNTS.map(function (a) { return [a.side, a.label, a.bank, a.holder, a.number]; });
      acc.getRange(2, 1, arows.length, 5).setValues(arows);
    }
  }

  var gal = book.getSheetByName(TAB_GAL);
  if (!gal) {
    gal = book.insertSheet(TAB_GAL);
    gal.getRange(1, 1, 1, 1).setValues([['fileId']]);
  }
  return { info: info, acc: acc, gal: gal };
}

/* 정보 탭을 {key:value} 맵으로 읽기 */
function readInfoMap() {
  var info = ensureSheets().info;
  var last = info.getLastRow();
  var map = {};
  if (last >= 2) {
    var vals = info.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var key = String(vals[i][0]).trim();
      if (key) map[key] = vals[i][1];
    }
  }
  return map;
}

/* 정보 탭에 부분 업데이트(있으면 갱신, 없으면 추가) */
function writeInfo(updates) {
  var info = ensureSheets().info;
  var last = info.getLastRow();
  var rowByKey = {};
  if (last >= 2) {
    var keys = info.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      var k = String(keys[i][0]).trim();
      if (k) rowByKey[k] = i + 2;
    }
  }
  for (var key in updates) {
    if (!updates.hasOwnProperty(key)) continue;
    var val = updates[key];
    if (rowByKey[key]) {
      info.getRange(rowByKey[key], 2).setValue(val);
    } else {
      var r = info.getLastRow() + 1;
      info.getRange(r, 1, 1, 2).setValues([[key, val]]);
      rowByKey[key] = r;
    }
  }
}

/* 갤러리 fileId 목록(순서 = 행 순서) */
function readGalleryIds() {
  var gal = ensureSheets().gal;
  var last = gal.getLastRow();
  if (last < 2) return [];
  var vals = gal.getRange(2, 1, last - 1, 1).getValues();
  return vals.map(function (r) { return String(r[0]).trim(); }).filter(function (s) { return s; });
}

function writeGalleryIds(ids) {
  var gal = ensureSheets().gal;
  var last = gal.getLastRow();
  if (last >= 2) gal.getRange(2, 1, last - 1, 1).clearContent();
  if (ids.length) {
    gal.getRange(2, 1, ids.length, 1).setValues(ids.map(function (id) { return [id]; }));
  }
}

/* ========================= 데이터 조회(GET) ========================= */

function getData() {
  var m = readInfoMap();
  var galIds = readGalleryIds();

  var acc = ensureSheets().acc;
  var accounts = [];
  var lastA = acc.getLastRow();
  if (lastA >= 2) {
    var av = acc.getRange(2, 1, lastA - 1, 5).getValues();
    for (var i = 0; i < av.length; i++) {
      var side = String(av[i][0]).trim();
      if (!side) continue;
      accounts.push({
        side: side === 'bride' ? 'bride' : 'groom',
        label: String(av[i][1] || ''), bank: String(av[i][2] || ''),
        holder: String(av[i][3] || ''), number: String(av[i][4] || '')
      });
    }
  }

  var coverFileId = String(m.coverFileId || '');
  return {
    ok: true,
    greeting: String(m.greeting || ''),
    groom: { name: String(m.groom_name || ''), order: String(m.groom_order || ''), father: String(m.groom_father || ''), mother: String(m.groom_mother || '') },
    bride: { name: String(m.bride_name || ''), order: String(m.bride_order || ''), father: String(m.bride_father || ''), mother: String(m.bride_mother || '') },
    weddingDateTime: String(m.weddingDateTime || ''),
    venueName: String(m.venueName || ''),
    venueAddress: String(m.venueAddress || ''),
    transport: { car: String(m.transport_car || ''), transit: String(m.transport_transit || ''), airport: String(m.transport_airport || '') },
    accounts: accounts,
    coverFileId: coverFileId,
    coverImage: coverFileId ? driveImageUrl(coverFileId) : '',
    gallery: galIds.map(function (id) { return { fileId: id, src: driveImageUrl(id) }; }),
    dataVersion: Number(m.dataVersion || 0)
  };
}

/* ========================= 쓰기(POST) 액션 ========================= */

function saveText(data) {
  data = data || {};
  var g = data.groom || {}, b = data.bride || {}, t = data.transport || {};
  var cur = readInfoMap();
  var nextVer = Number(cur.dataVersion || 0) + 1;

  writeInfo({
    greeting: str(data.greeting, 2000),
    groom_name: str(g.name, 40), groom_order: str(g.order, 40), groom_father: str(g.father, 40), groom_mother: str(g.mother, 40),
    bride_name: str(b.name, 40), bride_order: str(b.order, 40), bride_father: str(b.father, 40), bride_mother: str(b.mother, 40),
    weddingDateTime: str(data.weddingDateTime, 40),
    venueName: str(data.venueName, 120),
    venueAddress: str(data.venueAddress, 200),
    transport_car: str(t.car, 2000), transport_transit: str(t.transit, 2000), transport_airport: str(t.airport, 2000),
    dataVersion: String(nextVer),
    updatedAt: new Date().toISOString()
  });

  // 계좌 탭 전체 교체
  var acc = ensureSheets().acc;
  var lastA = acc.getLastRow();
  if (lastA >= 2) acc.getRange(2, 1, lastA - 1, 5).clearContent();
  var list = (data.accounts || []).slice(0, 20).map(function (a) {
    a = a || {};
    return [a.side === 'bride' ? 'bride' : 'groom', str(a.label, 60), str(a.bank, 40), str(a.holder, 40), str(a.number, 60)];
  });
  if (list.length) acc.getRange(2, 1, list.length, 5).setValues(list);

  return getData();
}

function str(v, max) {
  var s = (v == null) ? '' : String(v);
  return max ? s.slice(0, max) : s;
}

function saveMediaToFolder(filename, mimeType, dataBase64) {
  var bytes = Utilities.base64Decode(dataBase64);
  var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', filename || 'photo.jpg');
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file;
}

function uploadMedia(p) {
  var kind = p.kind;
  if (kind === 'thumbnail') {
    // 고정 파일(og:image)의 '내용'을 교체해 URL을 유지하려면 고급 Drive 서비스 필요.
    try {
      var bytes = Utilities.base64Decode(p.dataBase64);
      var blob = Utilities.newBlob(bytes, p.mimeType || 'image/png', p.filename || 'og-image.png');
      Drive.Files.update({}, THUMB_FILE_ID, blob);   // 고급 서비스(Drive API) 필요
      return { ok: true, src: driveImageUrl(THUMB_FILE_ID) };
    } catch (e) {
      // 고급 서비스 미설정 등: 새 파일로 저장하고 안내(정적 og:image는 수동 교체 필요)
      var nf = saveMediaToFolder(p.filename || 'og-image.png', p.mimeType, p.dataBase64);
      return {
        ok: true, src: driveImageUrl(nf.getId()),
        note: '고정 썸네일 교체는 Drive API 고급 서비스가 필요합니다. 새 파일로 저장했습니다(링크 미리보기는 수동 반영 필요).'
      };
    }
  }

  var file = saveMediaToFolder(p.filename, p.mimeType, p.dataBase64);
  var fileId = file.getId();

  if (kind === 'cover') {
    writeInfo({ coverFileId: fileId });
    return { ok: true, fileId: fileId, src: driveImageUrl(fileId) };
  }
  // 기본: 갤러리에 추가
  var ids = readGalleryIds();
  ids.push(fileId);
  writeGalleryIds(ids);
  return { ok: true, fileId: fileId, gallery: ids.map(function (id) { return { fileId: id, src: driveImageUrl(id) }; }) };
}

function setCover(fileId) {
  writeInfo({ coverFileId: String(fileId || '') });
  return { ok: true, coverFileId: String(fileId || ''), coverImage: fileId ? driveImageUrl(fileId) : '' };
}

function deleteGallery(fileId) {
  fileId = String(fileId || '');
  var ids = readGalleryIds().filter(function (id) { return id !== fileId; });
  writeGalleryIds(ids);
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}

  // 삭제한 사진이 대표사진이었으면 대표사진 해제
  var m = readInfoMap();
  var result = { ok: true, gallery: ids.map(function (id) { return { fileId: id, src: driveImageUrl(id) }; }) };
  if (String(m.coverFileId || '') === fileId) {
    writeInfo({ coverFileId: '' });
    result.coverFileId = '';
    result.coverImage = '';
  }
  return result;
}

function reorderGallery(order) {
  if (!Array.isArray(order)) return { ok: false, error: 'bad order' };
  var existing = {};
  readGalleryIds().forEach(function (id) { existing[id] = true; });
  var clean = order.map(String).filter(function (id) { return existing[id]; });
  writeGalleryIds(clean);
  return { ok: true };
}

/* ========================= 라우팅 ========================= */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'get';
    if (action === 'get') return jsonOut(getData());
    return jsonOut({ ok: true, service: 'mobile-wedding', action: action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var p;
  try { p = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ ok: false, error: 'bad json' }); }

  var action = p.action;

  // 로그인은 비번 검증만
  if (action === 'login') {
    return jsonOut({ ok: checkPw(p.password) });
  }

  // 그 외 쓰기 액션은 모두 비번 필요
  if (!checkPw(p.password)) {
    return jsonOut({ ok: false, error: 'unauthorized' });
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e2) { return jsonOut({ ok: false, error: 'busy' }); }
  try {
    switch (action) {
      case 'saveText':       return jsonOut(saveText(p.data));
      case 'uploadMedia':    return jsonOut(uploadMedia(p));
      case 'setCover':       return jsonOut(setCover(p.fileId));
      case 'deleteGallery':  return jsonOut(deleteGallery(p.fileId));
      case 'reorderGallery': return jsonOut(reorderGallery(p.order));
      default:               return jsonOut({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* 편집기에서 한 번 실행하면 시트 탭/기본값을 만들고 권한을 부여받는다(선택). */
function setup() {
  ensureSheets();
  Logger.log('시트 준비 완료: ' + SHEET_ID);
}
