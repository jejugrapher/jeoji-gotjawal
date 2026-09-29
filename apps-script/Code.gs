/**
 * 저지곶자왈 사진 수첩 — 사진 받는 창구 (Google Apps Script 웹 앱)
 *
 * 웹페이지의 [사진 보내기] 양식이 사진을 한 장씩 이 스크립트로 보내면,
 *  - 사진은 FOLDER_ID 폴더 안에 "날짜_이름_접수번호" 하위 폴더를 만들어 저장하고
 *  - 이름·연락처·이메일은 사진 폴더가 아니라, 스크립트 소유자 드라이브의
 *    비공개 스프레드시트 「저지곶자왈 사진 접수 목록」에 한 줄씩 기록합니다.
 *
 * 배포: 배포 → 새 배포 → 유형 '웹 앱' → 실행 사용자 '나' → 액세스 권한 '모든 사용자'
 */

const FOLDER_ID = '1xCXSjjV7Aatcdgt9AG1BbBsyiu0ZJf_2';   // 사진이 모일 폴더 (fromPage)
const SHEET_NAME = '저지곶자왈 사진 접수 목록 (비공개)';
const MAX_BYTES = 20 * 1024 * 1024;                       // 장당 20MB
const MAX_PER_BATCH = 10;                                 // 한 번에 10장
const NOTIFY_OWNER = true;                                // 접수가 끝나면 소유자에게 메일 알림

const OK_EXT = /\.(jpe?g|png|webp|heic|heif|gif)$/i;

function doGet() {
  return json_({ ok: true, service: 'jeoji-gotjawal-photo-upload' });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    // 스팸 방지용 숨은 칸이 채워져 있으면 조용히 무시
    if (d.website) return json_({ ok: true });

    const name = clean_(d.name, 40);
    const phone = clean_(d.phone, 20);
    const email = clean_(d.email, 120);
    const note = clean_(d.note, 500);
    const batch = String(d.batch || '').replace(/[^a-z0-9]/gi, '').slice(0, 20);
    const index = Number(d.index) || 0;
    const total = Math.min(Number(d.total) || 1, MAX_PER_BATCH);

    if (!d.agree) throw new Error('개인정보 수집·이용 동의가 필요합니다.');
    if (!name) throw new Error('이름이 비어 있습니다.');
    if (!/^0\d{8,10}$/.test(phone.replace(/\D/g, ''))) throw new Error('연락처 형식이 올바르지 않습니다.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('이메일 형식이 올바르지 않습니다.');
    if (!batch) throw new Error('접수 번호가 없습니다.');
    if (index >= MAX_PER_BATCH) throw new Error('한 번에 ' + MAX_PER_BATCH + '장까지 보낼 수 있습니다.');

    const filename = String(d.filename || 'photo.jpg').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
    const mime = String(d.mime || '');
    if (!(mime.indexOf('image/') === 0 || OK_EXT.test(filename))) throw new Error('사진 파일만 받을 수 있습니다.');

    const bytes = Utilities.base64Decode(String(d.data || ''));
    if (!bytes.length) throw new Error('파일이 비어 있습니다.');
    if (bytes.length > MAX_BYTES) throw new Error('20MB를 넘는 사진은 받을 수 없습니다.');

    const folder = batchFolder_(batch, name);
    const blob = Utilities.newBlob(bytes, mime || guessMime_(filename), pad_(index + 1) + '_' + filename);
    const file = folder.createFile(blob);

    logRow_([new Date(), batch, name, phone, email, note, d.promo ? '동의' : '미동의',
             file.getName(), file.getUrl(), Math.round(bytes.length / 1024) + 'KB']);

    if (NOTIFY_OWNER && index === total - 1) notify_(name, total, folder.getUrl());

    return json_({ ok: true, file: file.getName() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** 접수 한 건(batch)마다 하위 폴더 하나. 동시에 여러 장이 들어와도 폴더가 하나만 생기도록 잠금 사용 */
function batchFolder_(batch, name) {
  const props = PropertiesService.getScriptProperties();
  const key = 'batch_' + batch;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = props.getProperty(key);
    if (id) return DriveApp.getFolderById(id);
    const day = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
    const safe = name.replace(/[\\/:*?"<>|]/g, '_');
    const f = DriveApp.getFolderById(FOLDER_ID).createFolder(day + '_' + safe + '_' + batch.slice(-4));
    props.setProperty(key, f.getId());
    return f;
  } finally {
    lock.releaseLock();
  }
}

/** 연락처 기록용 시트: 사진 폴더 밖(스크립트 소유자의 내 드라이브)에 비공개로 만듦 */
function logRow_(row) {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let id = props.getProperty('sheet_id');
    let ss;
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create(SHEET_NAME);
      ss.getSheets()[0].appendRow(['접수 시각', '접수 번호', '이름', '연락처', '이메일', '사진 이야기', '홍보 활용', '파일', '파일 링크', '크기']);
      ss.getSheets()[0].setFrozenRows(1);
      props.setProperty('sheet_id', ss.getId());
    }
    ss.getSheets()[0].appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

function notify_(name, total, folderUrl) {
  try {
    const to = Session.getEffectiveUser().getEmail();
    if (to) MailApp.sendEmail(to, '[저지곶자왈 사진 수첩] ' + name + '님이 사진 ' + total + '장을 보냈습니다', '폴더: ' + folderUrl);
  } catch (e) { /* 알림 실패는 접수에 영향 없음 */ }
}

/** 처음 한 번 편집기에서 직접 실행: 권한 승인 + 폴더·시트 접근 확인 */
function setup() {
  const f = DriveApp.getFolderById(FOLDER_ID);
  Logger.log('사진 폴더: ' + f.getName() + ' — ' + f.getUrl());
  logRow_([new Date(), 'setup', '(설정 확인)', '', '', '이 줄은 지워도 됩니다', '', '', '', '']);
  Logger.log('접수 목록 시트: ' + SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('sheet_id')).getUrl());
}

function clean_(v, n) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n); }
function pad_(n) { return ('0' + n).slice(-2); }
function guessMime_(f) {
  const x = (f.split('.').pop() || '').toLowerCase();
  return { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp', heic: 'image/heic', heif: 'image/heif', gif: 'image/gif' }[x] || 'application/octet-stream';
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
