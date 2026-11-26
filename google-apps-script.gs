// Cọ Xinh Studio — nhận lưu bút từ website vào Google Sheet.
// Dán toàn bộ file này vào Tiện ích mở rộng → Apps Script của một Google Sheet, rồi Triển khai dưới dạng Ứng dụng web.
// Trang tính "LuuBut" tự được tạo ở lần gửi đầu tiên. Thiệp chỉ hiện lên website khi cô tick ô "Duyệt".
// Tab "Khách hàng" của Sheet kế hoạch = danh sách khách hàng tiềm năng: mỗi lưu bút có để SĐT và mỗi đơn đăng ký
// học thử trên web đều tự thành một dòng (trùng SĐT thì gộp vào dòng cũ). Tiêu đề ở hàng 5, dữ liệu từ hàng 6.
// Cột A–Q là phần theo dõi thủ công; cột R–U (Thời gian nhận, Loại đăng ký, Khung giờ, Cảm nhận) do web tự điền.
// CÀI ĐẶT 1 LẦN: mở Sheet kế hoạch → Tiện ích mở rộng → Apps Script → dán file này → chạy setup() → Triển khai ứng dụng web.

const SHEET = 'LuuBut';
const LEADS = 'Khách hàng';   // tên tab có sẵn trong Sheet kế hoạch
const NOTIFY_EMAIL = '';      // điền email để nhận thư mỗi khi có khách mới (để trống = không gửi)
const LEAD_HEADER_ROW = 5;    // tiêu đề ở hàng 5, khách đầu tiên ở hàng 6
const LEAD_EXTRA_HEAD = ['Thời gian nhận', 'Loại đăng ký', 'Khung giờ mong muốn', 'Cảm nhận từ lưu bút'];   // cột R–U
const HEAD = ['Thời gian', 'Duyệt', 'ID', 'Ba/mẹ', 'Bé', 'Bức tranh', 'Tem', 'Lời nhắn'];

// cắt độ dài + chặn công thức: chuỗi bắt đầu bằng = + - @ sẽ bị Sheets chạy như công thức nếu không có dấu '
function clean(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// website gửi một tấm thiệp
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (e.postData.contents.length > 4000) throw new Error('too big');
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'lead') {   // đăng ký học thử từ website
      const name = clean(d.name, 40), phone = clean(d.phone, 16);
      if (!name || !phoneKey_(phone)) throw new Error('empty');
      upsertLead_({ parent: name, phone: phone, age: clean(d.age, 20), slot: clean(d.slot, 60),
                    source: 'Website (học thử)', type: 'Đăng ký học thử', status: 'Đã hẹn Trial' });
      return json_({ ok: true });
    }
    const stamps = clean((Array.isArray(d.stamps) ? d.stamps : []).slice(0, 6).map(function (s) { return String(s).slice(0, 60); }).join(' | '), 400);
    const parent = clean(d.parent, 40), msg = clean(d.msg, 320);
    if (!parent || (!msg && !stamps)) throw new Error('empty');
    const sh = sheet_();
    sh.appendRow([new Date(), false, clean(d.id, 40), parent, clean(d.child, 40), clean(d.about, 40), stamps, msg]);
    sh.getRange(sh.getLastRow(), 2).insertCheckboxes();
    try {   // lưu bút nào cũng thành một khách tiềm năng; lỗi ở bước này không được làm hỏng việc nhận thiệp
      upsertLead_({ parent: parent, phone: clean(d.phone, 16), child: clean(d.child, 40), source: 'Website (lưu bút)', type: 'Lưu bút',
                    feedback: [stamps, msg].filter(String).join(' — '), status: 'Mới từ website' });
    } catch (err2) {}
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

// website đọc các thiệp ĐÃ DUYỆT để ghim lên Bức Tường Lưu Bút (mới nhất trước, tối đa 60)
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1)
    .filter(function (r) { return r[1] === true; }).reverse().slice(0, 60);
  return json_(rows.map(function (r) {
    return { id: String(r[2]), parent: String(r[3]), child: String(r[4]), about: String(r[5]),
             stamps: r[6] ? String(r[6]).split(' | ') : [], msg: String(r[7]) };
  }));
}


// ===================== KHÁCH HÀNG TIỀM NĂNG =====================

// 84912345678 / +84 912 345 678 / 0912345678 → 0912345678 (so khớp SĐT bất kể cách gõ)
function phoneKey_(p) {
  let d = String(p == null ? '' : p).replace(/\D/g, '');
  if (d.indexOf('84') === 0 && d.length >= 11) d = '0' + d.slice(2);
  if (d.length === 9) d = '0' + d;   // Sheets có thể đã nuốt số 0 đầu
  return d.length >= 10 && d.length <= 11 ? d : '';
}

function leadSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LEADS);
  if (!sh) throw new Error('Không thấy tab "' + LEADS + '" - hãy mở script từ chính Sheet kế hoạch.');
  // bổ sung tiêu đề cột R–U nếu chưa có
  const head = sh.getRange(LEAD_HEADER_ROW, 18, 1, 4);
  if (!String(head.getValue())) {
    head.setValues([LEAD_EXTRA_HEAD]).setFontWeight('bold').setBackground('#2E1F21').setFontColor('#FFFFFF').setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    [130, 120, 140, 320].forEach(function (w, i) { sh.setColumnWidth(18 + i, w); });
  }
  return sh;
}

// thêm khách mới, hoặc gộp vào dòng cũ nếu trùng SĐT
function upsertLead_(o) {
  const sh = leadSheet_();
  const first = LEAD_HEADER_ROW + 1;
  const n = Math.max(sh.getMaxRows() - LEAD_HEADER_ROW, 1);
  const abc = sh.getRange(first, 1, n, 3).getValues();   // Mã | Họ tên | SĐT
  const key = phoneKey_(o.phone);
  let row = 0, blank = 0;
  for (let i = 0; i < abc.length; i++) {
    if (key && phoneKey_(abc[i][2]) === key) { row = first + i; break; }
    if (!blank && !abc[i][0] && !abc[i][1] && !abc[i][2]) blank = first + i;   // dòng trống đầu tiên
  }
  const isNew = !row;
  if (isNew) {
    row = blank || sh.getMaxRows() + 1;
    if (row > sh.getMaxRows()) sh.insertRowAfter(sh.getMaxRows());
    sh.getRange(row, 3).setNumberFormat('@');
    sh.getRange(row, 7).setNumberFormat('dd/mm/yyyy'); sh.getRange(row, 10, 1, 3).setNumberFormat('dd/mm/yyyy'); sh.getRange(row, 18).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(row, 1, 1, 9).setValues([['KH' + ('000' + (row - LEAD_HEADER_ROW)).slice(-3), o.parent || '', key || '', o.child || '', o.age || '', o.source || '', '', '', '']]);
    sh.getRange(row, 13, 1, 3).setValues([[o.type === 'Đăng ký học thử' ? 'Đã hẹn' : 'Chưa hẹn', key ? (o.status || 'Mới từ website') : 'Chưa có SĐT', '']]);
    sh.getRange(row, 18, 1, 4).setValues([[new Date(), o.type || '', o.slot || '', o.feedback || '']]);
  } else {   // khách đã có: bổ sung thông tin còn thiếu, không ghi đè
    const cur = sh.getRange(row, 1, 1, 21).getValues()[0];
    const fill = function (col, v) { if (v && !cur[col - 1]) sh.getRange(row, col).setValue(v); };
    fill(4, o.child); fill(5, o.age); fill(20, o.slot);
    const types = String(cur[18] || ''); if (o.type && types.indexOf(o.type) < 0) sh.getRange(row, 19).setValue(types ? types + ' + ' + o.type : o.type);
    if (o.feedback) sh.getRange(row, 21).setValue(cur[20] ? cur[20] + ' | ' + o.feedback : o.feedback);
    if (o.type === 'Đăng ký học thử') { sh.getRange(row, 14).setValue('Đã hẹn Trial'); sh.getRange(row, 13).setValue('Đã hẹn'); }
  }
  // ngày nhắc & hạn ưu đãi: tính từ ngày workshop (G); chưa có thì dùng ngày nhận (R). setFormula luôn dùng dấu phẩy.
  sh.getRange(row, 10).setFormula('=IF(G' + row + '="",IF(R' + row + '="","",INT(R' + row + ')+3),G' + row + '+3)');
  sh.getRange(row, 11).setFormula('=IF(G' + row + '="",IF(R' + row + '="","",INT(R' + row + ')+7),G' + row + '+7)');
  sh.getRange(row, 17).setFormula('=IF(OR(K' + row + '="",N' + row + '="Đã đăng ký học",N' + row + '="Từ chối"),"",K' + row + '-TODAY())');
  if (NOTIFY_EMAIL) {
    try {
      MailApp.sendEmail(NOTIFY_EMAIL, (isNew ? 'Khách mới' : 'Khách quay lại') + ' từ website Cọ Xinh: ' + (o.parent || ''),
        'Loại: ' + (o.type || '') + '\nTên: ' + (o.parent || '') + '\nSĐT: ' + (key || '(chưa có)') + '\nBé: ' + (o.child || '') + ' ' + (o.age || '') +
        '\nKhung giờ: ' + (o.slot || '') + '\n' + (o.feedback || '') + '\n\nMở sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    } catch (err) {}
  }
}

// chạy MỘT LẦN từ trình soạn Apps Script: tạo tab LuuBut và thêm cột R–U cho tab Khách hàng
function setup() {
  sheet_(); leadSheet_();
}
