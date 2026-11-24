// Cọ Xinh Studio — nhận lưu bút từ website vào Google Sheet.
// Dán toàn bộ file này vào Tiện ích mở rộng → Apps Script của một Google Sheet, rồi Triển khai dưới dạng Ứng dụng web.
// Trang tính "LuuBut" tự được tạo ở lần gửi đầu tiên. Thiệp chỉ hiện lên website khi cô tick ô "Duyệt".
// Trang tính "KhachHang" = danh sách khách hàng tiềm năng: mỗi lưu bút có để SĐT và mỗi đơn đăng ký học thử trên web
// đều tự thành một dòng (trùng SĐT thì gộp vào dòng cũ). Cột A–Q giống sheet 'Khách hàng' trong file Excel kế hoạch.
// CÀI ĐẶT 1 LẦN: dán file này vào Apps Script → chạy hàm setup() → Triển khai ứng dụng web (xem README bên dưới).

const SHEET = 'LuuBut';
const LEADS = 'KhachHang';
const NOTIFY_EMAIL = '';   // điền email để nhận thư mỗi khi có khách mới (để trống = không gửi)
const LEAD_HEAD = ['Mã', 'Họ tên ba/mẹ', 'Zalo / SĐT', 'Tên bé', 'Tuổi', 'Nguồn biết đến', 'Buổi workshop (ngày)', 'Đã đến buổi?',
  'Đã gửi phiếu quan sát?', 'Nhắc ngày 3', 'Hạn ưu đãi (ngày 7)', 'Ngày hẹn Trial Class', 'Kết quả Trial', 'Trạng thái cuối',
  'Mức quan tâm', 'Ghi chú', 'Còn lại (ngày)', 'Thời gian nhận', 'Loại đăng ký', 'Khung giờ mong muốn', 'Cảm nhận từ lưu bút'];
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
  let sh = ss.getSheetByName(LEADS);
  if (!sh) {
    sh = ss.insertSheet(LEADS);
    sh.appendRow(LEAD_HEAD);
    const h = sh.getRange(1, 1, 1, LEAD_HEAD.length);
    h.setFontWeight('bold').setBackground('#2E1F21').setFontColor('#FFFFFF').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    sh.getRange('C:C').setNumberFormat('@');   // giữ số 0 đầu của SĐT
    sh.getRange('G:G').setNumberFormat('dd/mm/yyyy'); sh.getRange('J:L').setNumberFormat('dd/mm/yyyy'); sh.getRange('R:R').setNumberFormat('dd/mm/yyyy hh:mm');
    const list = function (col, items) {
      sh.getRange(col + '2:' + col + '1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build());
    };
    list('F', ['Website (lưu bút)', 'Website (học thử)', 'Nhóm cư dân', 'Facebook/Instagram', 'TikTok', 'Trường học', 'Bạn bè giới thiệu', 'Phụ huynh cũ', 'Quán cà phê', 'Khác']);
    list('H', ['Có', 'Không']); list('I', ['Có', 'Không']);
    list('M', ['Chưa hẹn', 'Đã hẹn', 'Đã đến', 'Vắng']);
    list('N', ['Mới từ website', 'Mới đăng ký workshop', 'Đã dự workshop', 'Đã hẹn Trial', 'Đang cân nhắc', 'Đã đăng ký học', 'Từ chối']);
    list('O', ['Cao', 'Trung bình', 'Thấp']);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Đã đăng ký học').setBackground('#DCE8D6').setRanges([sh.getRange('N2:N1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Mới từ website').setBackground('#FFF0C2').setRanges([sh.getRange('N2:N1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Từ chối').setBackground('#E6E1DE').setRanges([sh.getRange('N2:N1000')]).build()
    ]);
    [70, 150, 110, 100, 50, 130, 100, 80, 90, 90, 110, 110, 100, 130, 90, 260, 80, 130, 110, 140, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  }
  return sh;
}

// thêm khách mới, hoặc gộp vào dòng cũ nếu trùng SĐT
function upsertLead_(o) {
  const sh = leadSheet_();
  const key = phoneKey_(o.phone);
  const last = sh.getLastRow();
  let row = 0;
  if (key && last > 1) {
    const phones = sh.getRange(2, 3, last - 1, 1).getValues();
    for (let i = 0; i < phones.length; i++) if (phoneKey_(phones[i][0]) === key) { row = i + 2; break; }
  }
  const isNew = !row;
  if (isNew) {
    row = last + 1;
    sh.getRange(row, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, 21).setValues([[
      'KH' + ('000' + (row - 1)).slice(-3), o.parent || '', key || '', o.child || '', o.age || '', o.source || '', '', '', '', '', '', '', 'Chưa hẹn',
      key ? (o.status || 'Mới từ website') : 'Chưa có SĐT', '', '', '', new Date(), o.type || '', o.slot || '', o.feedback || ''
    ]]);
    sh.getRange(row, 12).setValue('');
  } else {   // khách đã có: bổ sung thông tin còn thiếu, không ghi đè
    const cur = sh.getRange(row, 1, 1, 21).getValues()[0];
    const fill = function (col, v) { if (v && !cur[col - 1]) sh.getRange(row, col).setValue(v); };
    fill(4, o.child); fill(5, o.age); fill(20, o.slot);
    const types = String(cur[18] || ''); if (o.type && types.indexOf(o.type) < 0) sh.getRange(row, 19).setValue(types ? types + ' + ' + o.type : o.type);
    if (o.feedback) sh.getRange(row, 21).setValue(cur[20] ? cur[20] + ' | ' + o.feedback : o.feedback);
    if (o.type === 'Đăng ký học thử') { sh.getRange(row, 14).setValue('Đã hẹn Trial'); sh.getRange(row, 13).setValue('Đã hẹn'); }
  }
  if (isNew && o.type === 'Đăng ký học thử') sh.getRange(row, 13).setValue('Đã hẹn');
  // ngày nhắc & hạn ưu đãi tính từ ngày nhận (G = ngày workshop, nếu để trống thì dùng ngày nhận)
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

// chạy MỘT LẦN từ trình soạn Apps Script để tạo sẵn cả hai trang tính
function setup() {
  sheet_(); leadSheet_();
  SpreadsheetApp.getActiveSpreadsheet().rename('Cọ Xinh Studio · Khách hàng & Lưu bút');
}
