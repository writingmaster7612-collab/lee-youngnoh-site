/**
 * 강의 문의 접수 스크립트 (Google Apps Script)
 *
 * 사이트의 문의 폼이 POST로 보낸 내용을
 *   1) 구글 시트에 한 줄 추가하고
 *   2) 운영자에게 '새 문의 도착' 알림 메일을 보내고
 *   3) 신청자에게 '문의가 접수되었습니다' 자동회신을 보낸다.
 *
 * 받는 항목(name 속성): name, email, phone, company, message
 * 시트 헤더 순서: 이름 / 이메일 / 전화번호 / 회사,직함 / 강의 문의 내용
 *
 * 이 스크립트는 문의를 쌓을 구글 시트에서 [확장 프로그램 > Apps Script]로 열어 붙여 넣는다.
 */

const OWNER_EMAIL = 'blueskyhanoi@naver.com';
const OWNER_NAME = '이영노';
const SHEET_NAME = '문의';
const HEADERS = ['이름', '이메일', '전화번호', '회사,직함', '강의 문의 내용'];

// 항목별 최대 글자 수. 넘으면 접수하지 않는다.
const MAX_LENGTH = { name: 50, email: 100, phone: 30, company: 100, message: 3000 };

function doPost(e) {
  try {
    const data = readRequest_(e);

    // 사람 눈에 안 보이는 칸(website)에 값이 있으면 자동 프로그램으로 보고 조용히 넘긴다
    if (data.website) return respond_({ ok: true });

    const error = validate_(data);
    if (error) return respond_({ ok: false, error: error });

    appendRow_(data);

    // 메일이 실패해도 시트에는 이미 기록됐으므로 접수는 성공으로 처리한다
    const mailErrors = [];
    try { notifyOwner_(data); } catch (err) { mailErrors.push('알림: ' + err.message); }
    try { replyToApplicant_(data); } catch (err) { mailErrors.push('자동회신: ' + err.message); }
    if (mailErrors.length) console.error('메일 발송 실패 - ' + mailErrors.join(' / '));

    return respond_({ ok: true });
  } catch (err) {
    console.error(err);
    return respond_({ ok: false, error: '접수 중 문제가 생겼습니다. 잠시 후 다시 시도해 주세요.' });
  }
}

// 브라우저에서 웹 앱 주소를 열었을 때 동작 확인용
function doGet() {
  return respond_({ ok: true, message: '문의 접수 스크립트가 동작 중입니다.' });
}

/** 폼 전송(application/x-www-form-urlencoded, FormData)과 JSON 본문을 모두 받는다 */
function readRequest_(e) {
  let raw = (e && e.parameter) || {};
  const body = e && e.postData;
  if (body && body.contents && String(body.type).indexOf('json') !== -1) {
    raw = JSON.parse(body.contents);
  }
  const clean = function (value) { return String(value == null ? '' : value).trim(); };
  return {
    name: clean(raw.name),
    email: clean(raw.email),
    phone: clean(raw.phone),
    company: clean(raw.company),
    message: clean(raw.message),
    website: clean(raw.website)
  };
}

function validate_(data) {
  if (!data.name) return '이름을 입력해 주세요.';
  if (!data.email) return '이메일을 입력해 주세요.';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) return '이메일 형식을 확인해 주세요.';
  if (!data.message) return '강의 문의 내용을 입력해 주세요.';
  for (const key in MAX_LENGTH) {
    if (data[key].length > MAX_LENGTH[key]) return '입력한 내용이 너무 깁니다.';
  }
  return '';
}

function appendRow_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // 동시에 들어온 문의가 같은 줄에 겹쳐 쓰이지 않게 한다
  try {
    const sheet = getSheet_();
    const row = [data.name, data.email, data.phone, data.company, data.message].map(safeCell_);
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length);
    range.setNumberFormat('@'); // 텍스트 서식: 010으로 시작하는 전화번호의 앞자리 0이 지워지지 않게 한다
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** =, +, -, @ 로 시작하는 값이 시트에서 수식으로 실행되지 않게 한다 */
function safeCell_(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

/** 메일 제목에 넣을 값에서 줄바꿈을 없앤다 */
function oneLine_(value) {
  return value.replace(/[\r\n]+/g, ' ');
}

function notifyOwner_(data) {
  const receivedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const body = [
    '사이트에서 새 강의 문의가 들어왔습니다.',
    '',
    '접수 시각: ' + receivedAt,
    '이름: ' + data.name,
    '이메일: ' + data.email,
    '전화번호: ' + (data.phone || '(입력 없음)'),
    '회사,직함: ' + (data.company || '(입력 없음)'),
    '',
    '[강의 문의 내용]',
    data.message,
    '',
    '이 메일에 답장하면 신청자에게 바로 회신됩니다.',
    '전체 문의 목록: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].join('\n');

  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: '[새 문의 도착] ' + oneLine_(data.name) + (data.company ? ' (' + oneLine_(data.company) + ')' : ''),
    body: body,
    replyTo: data.email,
    name: '강의 문의 알림'
  });
}

function replyToApplicant_(data) {
  // 신청자가 쓴 문의 내용은 회신에 넣지 않는다.
  // 남의 이메일 주소를 적어 넣어 원하는 글을 대신 보내게 하는 악용을 막기 위해서다.
  const body = [
    oneLine_(data.name) + '님, 안녕하세요.',
    '',
    '강의 문의가 정상적으로 접수되었습니다.',
    '보내주신 내용을 확인한 뒤 ' + OWNER_NAME + ' 강사가 직접 회신드리겠습니다.',
    '',
    '추가로 전하실 내용이 있으면 이 메일에 답장해 주세요.',
    '',
    '감사합니다.',
    OWNER_NAME + ' 드림',
    OWNER_EMAIL,
    '',
    '※ 본인이 문의하지 않았는데 이 메일을 받으셨다면 무시하셔도 됩니다.'
  ].join('\n');

  MailApp.sendEmail({
    to: data.email,
    subject: '문의가 접수되었습니다 | ' + OWNER_NAME + ' 리테일 강의·자문',
    body: body,
    replyTo: OWNER_EMAIL,
    name: OWNER_NAME
  });
}

function respond_(result) {
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 편집기에서 직접 실행하는 시험용 함수.
 * 처음 실행하면 시트·메일 권한 승인 창이 뜬다.
 * 실행하면 시트에 시험용 한 줄이 추가되고, 알림과 자동회신이 모두 운영자 주소로 간다.
 */
function testDoPost() {
  const result = doPost({
    parameter: {
      name: '홍길동',
      email: OWNER_EMAIL,
      phone: '010-1234-5678',
      company: '시험 주식회사, 교육팀 과장',
      message: '시험용 문의입니다. 베트남 진출 관련 2시간 특강이 가능한지 문의드립니다.'
    }
  });
  console.log(result.getContent());
}
