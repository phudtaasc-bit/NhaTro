/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

function NT_kiemTraDuLieuNguon() {
  const issues = [];
  const current = NT_docSheetDangThue_();
  const returned = NT_docSheetTraPhong_();

  current.concat(returned).forEach(r => {
    if (!r.room) {
      issues.push(r.source + ' - dòng ' + r.sourceRow + ': thiếu số phòng.');
    }
    if (!r.name) {
      issues.push(r.source + ' - dòng ' + r.sourceRow + ': thiếu họ tên.');
    }
    if (!r.id) {
      issues.push(
        r.source + ' - dòng ' + r.sourceRow +
        ': thiếu CCCD/Hộ chiếu; không thể đối chiếu trả phòng chính xác.'
      );
    }
    if (!r.effectiveDate && !r.contractDate) {
      issues.push(
        r.source + ' - dòng ' + r.sourceRow +
        ': thiếu ngày hiệu lực và ngày ký hợp đồng.'
      );
    }
    if (!r.rent) {
      issues.push(r.source + ' - dòng ' + r.sourceRow + ': thiếu giá phòng.');
    }
    if (r.source === NT.SOURCE_RETURNED && !r.returnDate) {
      issues.push(
        r.source + ' - dòng ' + r.sourceRow + ': chưa có ngày trả phòng.'
      );
    }
  });

  const ui = SpreadsheetApp.getUi();

  if (issues.length === 0) {
    ui.alert('Không phát hiện lỗi bắt buộc trong dữ liệu nguồn.');
    return;
  }

  const shown = issues.slice(0, 50);
  ui.alert(
    'Phát hiện ' + issues.length + ' vấn đề:\n\n' +
    shown.join('\n') +
    (issues.length > 50 ? '\n\n... và các lỗi khác.' : '')
  );
}
