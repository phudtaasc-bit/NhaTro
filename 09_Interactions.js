/**
 * Tương tác trên sheet tháng.
 */

/**
 * Trả về danh sách phòng trống tại cuối tháng báo cáo.
 * Logic này đồng nhất với chỉ tiêu Số phòng trống trên dashboard.
 */
function NT_layDanhSachPhongTrongCuoiThang_(monthDate) {
  const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  const sourceRoomState = NT_docTrangThaiPhongNguon_();
  const currentRecords = NT_docSheetDangThue_();
  const returnedRecords = NT_docSheetTraPhong_();
  const mergedRecords = NT_gopVaLoaiTrungKhach_(currentRecords, returnedRecords);

  const allRooms = new Set();
  sourceRoomState.allRooms.forEach(room => allRooms.add(room));
  returnedRecords.forEach(record => {
    if (record.room) allRooms.add(NT_chuanHoaPhong_(record.room));
  });

  const occupiedRooms = new Set();
  mergedRecords.forEach(record => {
    const room = NT_chuanHoaPhong_(record.room);
    if (!room) return;

    const effective = record.effectiveDate || record.contractDate;
    const startedByMonthEnd = !effective || effective <= monthEnd;
    const notReturnedByMonthEnd = !record.returnDate || record.returnDate > monthEnd;

    if (startedByMonthEnd && notReturnedByMonthEnd) occupiedRooms.add(room);
  });

  // Phòng có số phòng nhưng không có người thuê trong TH thuê trọ luôn là phòng trống.
  sourceRoomState.vacantRooms.forEach(room => occupiedRooms.delete(room));

  return Array.from(allRooms)
    .filter(room => !occupiedRooms.has(room))
    .sort((a, b) => NT_roomSort_(a) - NT_roomSort_(b));
}

/**
 * Khi người dùng chọn ô giá trị "Số phòng trống" tại J4:L4,
 * hiển thị danh sách phòng trống bằng thông báo ở góc dưới màn hình.
 */
function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const monthDate = NT_layThangTuTenSheet_(sheet.getName());
    if (!monthDate) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Ô giá trị Số phòng trống là vùng hợp nhất J4:L4.
    if (row !== 4 || col < 10 || col > 12) return;

    const rooms = NT_layDanhSachPhongTrongCuoiThang_(monthDate);
    const message = rooms.length > 0
      ? rooms.join(', ')
      : 'Không có phòng trống.';

    SpreadsheetApp.getActive().toast(
      message,
      'Danh sách phòng trống (' + rooms.length + ' phòng)',
      10
    );
  } catch (error) {
    console.error('Lỗi onSelectionChange:', error);
  }
}
