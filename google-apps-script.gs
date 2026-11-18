// Cọ Xinh Studio — nhận lưu bút từ website vào Google Sheet.
// Dán toàn bộ file này vào Tiện ích mở rộng → Apps Script của một Google Sheet, rồi Triển khai dưới dạng Ứng dụng web.
// Trang tính "LuuBut" tự được tạo ở lần gửi đầu tiên. Thiệp chỉ hiện lên website khi cô tick ô "Duyệt".

const SHEET = 'LuuBut';
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
    const stamps = clean((Array.isArray(d.stamps) ? d.stamps : []).slice(0, 6).map(function (s) { return String(s).slice(0, 60); }).join(' | '), 400);
    const parent = clean(d.parent, 40), msg = clean(d.msg, 320);
    if (!parent || (!msg && !stamps)) throw new Error('empty');
    const sh = sheet_();
    sh.appendRow([new Date(), false, clean(d.id, 40), parent, clean(d.child, 40), clean(d.about, 40), stamps, msg]);
    sh.getRange(sh.getLastRow(), 2).insertCheckboxes();
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
