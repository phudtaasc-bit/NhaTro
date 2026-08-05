/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 QUẢN LÝ NHÀ TRỌ')
    .addItem('Tạo sheet tháng tiếp theo', 'NT_taoThangTiepTheo')
    .addItem('Tạo sheet theo tháng lựa chọn', 'NT_taoThangTheoLuaChon')
    .addSeparator()
    .addItem('Cập nhật sheet tháng hiện tại', 'NT_capNhatThangHienTai')
    .addItem('Tính lại sheet tháng hiện tại', 'NT_tinhLaiThangHienTai')
    .addSeparator()
    .addItem('Tạo/Cập nhật sheet CẤU HÌNH', 'NT_taoHoacCapNhatCauHinh')
    .addItem('Kiểm tra dữ liệu nguồn', 'NT_kiemTraDuLieuNguon')
    .addToUi();
}
function NT_taoThangTiepTheo() {
  const ss = SpreadsheetApp.getActive();
  NT_taoHoacCapNhatCauHinh();

  const latest = NT_timSheetThangMoiNhat_(ss);
  if (!latest) {
    SpreadsheetApp.getUi().alert(
      'Không tìm thấy sheet tháng mẫu theo dạng "Tháng 8.2026".'
    );
    return;
  }

  const next = new Date(latest.year, latest.month, 1);
  NT_taoHoacCapNhatSheetThang_(next, false);
}
function NT_taoThangTheoLuaChon() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Tạo sheet tháng',
    'Nhập tháng cần tạo theo dạng MM/YYYY, ví dụ 09/2026:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const monthDate = NT_parseMonthInput_(response.getResponseText());
  if (!monthDate) {
    ui.alert('Tháng không hợp lệ. Hãy nhập theo dạng MM/YYYY.');
    return;
  }

  NT_taoHoacCapNhatCauHinh();
  NT_taoHoacCapNhatSheetThang_(monthDate, false);
}
function NT_capNhatThangHienTai() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const monthDate = NT_layThangTuTenSheet_(sheet.getName());

  if (!monthDate) {
    SpreadsheetApp.getUi().alert(
      'Hãy mở một sheet tháng theo dạng "Tháng 8.2026" rồi chạy lại.'
    );
    return;
  }

  NT_taoHoacCapNhatCauHinh();
  NT_taoHoacCapNhatSheetThang_(monthDate, true);
}
function NT_tinhLaiThangHienTai() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const monthDate = NT_layThangTuTenSheet_(sheet.getName());

  if (!monthDate) {
    SpreadsheetApp.getUi().alert(
      'Hãy mở một sheet tháng theo dạng "Tháng 8.2026" rồi chạy lại.'
    );
    return;
  }

  NT_apDungCongThucVaDinhDang_(sheet, monthDate);
  NT_taoTongCongVaDashboard_(sheet, monthDate);
  NT_dinhDangBaoCaoThang_(sheet);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã tính lại tiền điện, nước và công nợ.');
}
