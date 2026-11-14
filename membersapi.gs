/**
 * 홈페이지 '조합원 찾기'용 명단 API (구글 Apps Script).
 * 스크립트 속성에 SITE_API_KEY(길고 무작위인 문자열)를 등록해야 응답합니다.
 * 이름·소속·status(상태/그룹)만 전달하며, 다른 칸은 보내지 않습니다.
 *
 * [추가] 총회 참석자 확인 앱용: body.action === 'attend' 일 때
 *  - 스크립트 속성 ATTEND_API_KEY(별도의 길고 무작위인 문자열)가 필요합니다.
 *  - 이름·소속·status 에 더해 휴대폰 '뒤 4자리'(phone4)와 캠퍼스(campus)를 추가로 전달합니다.
 *
 * [추가] 개인별 인증번호(총회 참석 확인·투표용)
 *  - 스크립트 속성 CODE_SECRET(길고 무작위인 문자열)이 필요합니다.
 *  - action 'attend' + sid → 각 조합원의 인증번호(code)를 함께 전달
 *  - action 'sendcodes' → 조합원 이메일(직장메일 우선, 없으면 개인메일)로 인증번호를 1명씩 발송
 *  - 인증번호는 (CODE_SECRET + 총회ID + 이름 + 번호 뒤 4자리)로 계산되므로 따로 저장하지 않으며, 다시 계산해도 같은 값이 나옵니다.
 */
const CODE_EXCLUDE = ['퇴직', '자격정지'];   // 인증번호를 보내지 않을 그룹(상태)
const MAIL_SENDER = '한국외국어대학교 노동조합';
const MEMBERS_SHEET_ID = '1P2EmP2TFQ1o3uZDRCwzaBNYNe7wDnBlY';
const MEMBERS_SHEET_NAME = '조합원(전체)';
const NAME_KEYS = ['이름', '성명'];
const DEPT_KEYS = ['소속', '부서'];
const PHONE_KEYS = ['휴대폰', '휴대전화', '전화번호', '연락처'];
const CAMPUS_KEYS = ['캠퍼스', '근무지'];   // 캠퍼스 열(없으면 근무지 열 사용)
const STATUS_KEYS = ['status', '상태', '그룹', '구분'];   // 앞에 있는 이름을 먼저 찾습니다

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const send = body.action === 'sendcodes';
    const attend = body.action === 'attend' || send;
    const key = PropertiesService.getScriptProperties().getProperty(attend ? 'ATTEND_API_KEY' : 'SITE_API_KEY');
    if (!key || body.key !== key) return apiJson_({ error: 'unauthorized' });

    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(MEMBERS_SHEET_ID);
    const sh = ss.getSheetByName(MEMBERS_SHEET_NAME);
    if (!sh) return apiJson_({ error: 'sheet-not-found' });
    const rows = sh.getDataRange().getDisplayValues();

    // 제목 줄 찾기(위 5줄 안에서 '이름' 칸이 있는 줄)
    let h = -1, ni = -1, di = -1, si = -1, pi = -1, ci = -1, wi = -1, ei = -1;
    for (let r = 0; r < Math.min(5, rows.length) && ni < 0; r++) {
      ni = findCol_(rows[r], NAME_KEYS); di = findCol_(rows[r], DEPT_KEYS); si = findCol_(rows[r], STATUS_KEYS); pi = findCol_(rows[r], PHONE_KEYS); ci = findCol_(rows[r], CAMPUS_KEYS); wi = findCol_(rows[r], ['직장메일']); ei = findCol_(rows[r], ['개인메일', '이메일']); h = r;
    }
    if (ni < 0) return apiJson_({ error: 'name-column-not-found', headers: rows[0] });

    const members = [];
    for (let r = h + 1; r < rows.length; r++) {
      const name = String(rows[r][ni]).trim();
      if (!name) continue;
      const m = {
        name: name,
        dept: di >= 0 ? String(rows[r][di]).trim() : '',
        status: si >= 0 ? String(rows[r][si]).trim() : ''
      };
      if (attend) m.campus = ci >= 0 ? String(rows[r][ci]).trim() : '';
      if (attend) m.phone4 = pi >= 0 ? String(rows[r][pi]).replace(/\D/g, '').slice(-4) : '';
      if (attend) {
        const email = (wi >= 0 && String(rows[r][wi]).trim()) || (ei >= 0 && String(rows[r][ei]).trim()) || '';
        m.hasEmail = !!email;
        if (body.sid) m.code = code_(body.sid, name, m.phone4);
        if (send) m.email = email;
      }
      members.push(m);
    }
    if (send) return sendCodes_(body, members);
    return apiJson_({ members: members });
  } catch (err) {
    return apiJson_({ error: 'server-error' });
  }
}

function findCol_(hdr, keys) {
  const norm = function (s) { return String(s).replace(/\s/g, '').toLowerCase(); };
  for (let k = 0; k < keys.length; k++) {
    const i = hdr.findIndex(function (x) { return norm(x).indexOf(norm(keys[k])) >= 0; });
    if (i >= 0) return i;
  }
  return -1;
}

function apiJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 개인별 4자리 인증번호: HMAC(비밀키, 총회ID|이름|뒤4자리) → 0000~9999
function code_(sid, name, phone4) {
  const secret = PropertiesService.getScriptProperties().getProperty('CODE_SECRET');
  if (!secret) throw new Error('CODE_SECRET 미설정');
  const sig = Utilities.computeHmacSha256Signature(sid + '|' + String(name).replace(/\s+/g, '') + '|' + phone4, secret);
  let n = 0;
  for (let i = 0; i < 4; i++) n = n * 256 + (sig[i] & 255);
  return ('000' + (n % 10000)).slice(-4);
}

// 인증번호 메일 발송 (조합원 1명당 1통, 수신자는 본인만)
function sendCodes_(body, members) {
  if (!body.sid) return apiJson_({ error: 'sid-required' });
  const esc = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const only = String(body.only || '').replace(/\s+/g, '');
  const list = members.filter(function (m) {
    if (m.phone4.length < 4) return false;
    if (CODE_EXCLUDE.some(function (x) { return m.status.indexOf(x) >= 0; })) return false;
    return !only || m.name.replace(/\s+/g, '') === only;
  });
  // 테스트 발송: 지정한 주소로 1통만 보냄 (조합원 이름을 주면 그 조합원의 실제 번호, 없으면 샘플 번호)
  if (body.test) {
    const to = String(body.test).trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return apiJson_({ error: 'invalid-email' });
    let nm = '홍길동(샘플)', cd = '1234', real = false;
    if (only) { if (!list[0]) return apiJson_({ error: 'member-not-found' }); nm = list[0].name; cd = list[0].code; real = true; }
    MailApp.sendEmail({
      to: to, name: MAIL_SENDER,
      subject: '[테스트] [' + String(body.title || '총회').slice(0, 80) + '] 참석 확인 인증번호 안내',
      htmlBody: mailHtml_(nm, String(body.title || '총회').slice(0, 80), cd, body.link,
        real ? '※ 테스트 발송입니다. 위 번호는 ' + nm + ' 조합원의 실제 인증번호입니다.' : '※ 테스트 발송입니다. 위 번호는 샘플이며 실제 참석 확인에는 사용할 수 없습니다.')
    });
    return apiJson_({ ok: true, test: true, sample: !real, name: nm, quota: MailApp.getRemainingDailyQuota() });
  }
  const off = Math.max(0, Number(body.offset) || 0), lim = Math.min(100, Number(body.limit) || 60);
  const part = list.slice(off, off + lim);
  const title = String(body.title || '총회').slice(0, 80);
  let sent = 0, stopped = false;
  const failed = [], noEmail = [];
  for (let i = 0; i < part.length; i++) {
    const m = part[i];
    if (!m.email) { noEmail.push(m.name); continue; }
    if (MailApp.getRemainingDailyQuota() < 1) { stopped = true; break; }
    try {
      MailApp.sendEmail({
        to: m.email,
        subject: '[' + title + '] 참석 확인 인증번호 안내',
        name: MAIL_SENDER,
        htmlBody: mailHtml_(m.name, title, m.code, body.link, '')
      });
      sent++;
    } catch (err) { failed.push(m.name); }
  }
  const next = (!stopped && off + lim < list.length) ? off + lim : null;
  return apiJson_({ ok: true, total: list.length, sent: sent, noEmail: noEmail, failed: failed, next: next, stopped: stopped, quota: MailApp.getRemainingDailyQuota() });
}

function mailHtml_(name, title, code, link, note) {
  const esc = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  return (note ? '<p style="background:#fff3cd;padding:8px 12px;border-radius:6px">' + esc(note) + '</p>' : '') +
    '<p>' + esc(name) + ' 조합원님, 안녕하세요.</p>' +
    '<p><b>' + esc(title) + '</b>의 참석 확인과 투표에 사용하는 개인별 인증번호입니다.</p>' +
    '<p style="font-size:28px;letter-spacing:6px;margin:12px 0"><b>' + code + '</b></p>' +
    '<p>현장에서 안내되는 QR을 스마트폰으로 비추고, <b>이름 · 전화번호 뒤 4자리 · 위 인증번호</b>를 입력해 주세요.</p>' +
    (link ? '<p><a href="' + esc(link) + '">참석 확인 페이지 열기</a></p>' : '') +
    '<p style="color:#777;font-size:13px">이 번호는 본인 확인용이므로 다른 사람에게 알려주지 마세요.</p>';
}
