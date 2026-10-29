/**
 * Google Apps Script – nhận đăng ký ủng hộ sách từ trang web và ghi vào Google Sheet.
 * Cách dùng: xem HUONG-DAN.txt (dán toàn bộ tệp này vào Tiện ích mở rộng → Apps Script).
 */
const TEN_TRANG = "UngHoSach";
const COT = ["Thời gian","Mã ủng hộ","Họ tên","Đối tượng","Lớp","Điện thoại","Sách (chi tiết)","Danh sách sách","Tổng số cuốn","Ngày hẹn","Buổi","Nơi nhận","Địa chỉ","Lời trao gửi","Hiện tên trên nhãn"];

function layTrang_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TEN_TRANG);
  if (!sh) sh = ss.insertSheet(TEN_TRANG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COT);
    sh.getRange(1, 1, 1, COT.length).setFontWeight("bold").setBackground("#dbeafe");
    sh.setFrozenRows(1);
    sh.getRange("F:F").setNumberFormat("@");   // giữ số 0 đầu số điện thoại
    sh.getRange("J:J").setNumberFormat("@");   // ngày hẹn dạng chữ yyyy-mm-dd
  }
  return sh;
}

function doPost(e) {
  const khoa = LockService.getScriptLock();
  khoa.waitLock(15000);
  try {
    const d = JSON.parse((e.parameter && e.parameter.duLieu) || e.postData.contents);
    const sach = (d.sach || []).slice(0, 50).map(s => ({ ten: String(s.ten || "").slice(0, 120), tacGia: String(s.tacGia || "").slice(0, 60),
      loai: String(s.loai || "").slice(0, 40), soLuong: Math.max(1, Math.min(200, Number(s.soLuong) || 1)) }));
    const tom = sach.map(s => s.ten + (s.tacGia ? " – " + s.tacGia : "") + " (" + s.loai + ") ×" + s.soLuong).join("; ");
    layTrang_().appendRow([
      new Date(), String(d.ma || ""), String(d.hoTen || "").slice(0, 60), String(d.doiTuong || ""), String(d.lop || "").slice(0, 20),
      String(d.sdt || "").slice(0, 15), JSON.stringify(sach), tom, sach.reduce((a, s) => a + s.soLuong, 0),
      String(d.ngayHen || ""), String(d.buoi || ""), String(d.noiNhan || ""), String(d.diaChi || "").slice(0, 120),
      String(d.loiNhan || "").slice(0, 160), d.hienTen ? "Có" : "Không"
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, ma: d.ma })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, loi: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    khoa.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Trang nhận đăng ký ủng hộ sách đang hoạt động.");
}

/** Chạy thử một lần trong trình soạn Apps Script (nút ▶ Chạy) để cấp quyền và tạo trang tính. */
function chayThu() {
  layTrang_();
}
