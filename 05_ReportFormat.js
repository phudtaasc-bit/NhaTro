/**
 * Định dạng báo cáo tháng.
 */
function NT_dinhDangBaoCaoThang_(sheet) {
  const lastDataRow = NT_layDongCuoiDuLieuThang_(sheet);
  const totalRow = lastDataRow + 1;
  const numRows = Math.max(lastDataRow - NT.MONTH_DATA_ROW + 1, 1);
  const widths = {1:55,2:55,3:70,4:165,5:135,6:220,7:90,8:90,9:90,10:125,11:105,12:120,13:95,14:95,15:80,16:80,17:110,18:80,19:80,20:110,21:115,22:105,23:110,24:115,25:220};
  Object.keys(widths).forEach(col => sheet.setColumnWidth(Number(col), widths[col]));
  sheet.setRowHeights(1,4,30);
  sheet.setRowHeight(1,32);
  sheet.setRowHeight(NT.MONTH_HEADER_ROW,58);
  if (lastDataRow >= NT.MONTH_DATA_ROW) sheet.setRowHeights(NT.MONTH_DATA_ROW,numRows,24);
  sheet.setRowHeight(totalRow,28);

  sheet.getRange(NT.MONTH_HEADER_ROW,1,1,NT.MONTH_COLS)
    .setFontFamily('Times New Roman').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBackground('#d9e2f3').setFontColor('#000000');

  const body = sheet.getRange(NT.MONTH_DATA_ROW,1,numRows,NT.MONTH_COLS);
  body.setFontFamily('Times New Roman').setFontSize(11).setVerticalAlignment('middle').setWrap(false);

  // Tô xen kẽ theo từng phòng, toàn bộ người trong cùng phòng dùng cùng một màu.
  const roomCells = sheet.getRange(NT.MONTH_DATA_ROW,3,numRows,1).getDisplayValues();
  const rowBackgrounds = [];
  let currentRoom = '';
  let roomIndex = -1;
  for (let i = 0; i < numRows; i++) {
    const roomValue = NT_chuanHoaPhong_(roomCells[i][0]);
    if (roomValue) {
      currentRoom = roomValue;
      roomIndex++;
    }
    const color = roomIndex % 2 === 1 ? NT.ALT_ROOM_COLOR : NT.AUTO_COLOR;
    rowBackgrounds.push(new Array(NT.MONTH_COLS).fill(color));
  }
  body.setBackgrounds(rowBackgrounds);

  sheet.getRange(NT.MONTH_DATA_ROW,1,numRows,3).setHorizontalAlignment('center');
  sheet.getRange(NT.MONTH_DATA_ROW,7,numRows,5).setHorizontalAlignment('center');
  sheet.getRange(NT.MONTH_DATA_ROW,15,numRows,6).setHorizontalAlignment('right');
  sheet.getRange(NT.MONTH_DATA_ROW,23,numRows,1).setHorizontalAlignment('center');
  sheet.getRange(NT.MONTH_DATA_ROW,4,numRows,3).setHorizontalAlignment('left');
  sheet.getRange(NT.MONTH_DATA_ROW,25,numRows,1).setHorizontalAlignment('left').setWrap(true);
  [12,13,14,17,20,21,22,24].forEach(col => sheet.getRange(NT.MONTH_DATA_ROW,col,numRows,1).setHorizontalAlignment('right'));

  sheet.getRange(NT.MONTH_HEADER_ROW,1,totalRow-NT.MONTH_HEADER_ROW+1,NT.MONTH_COLS)
    .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);

  // Các cột nhập tay luôn giữ màu vàng, không phụ thuộc màu xen kẽ phòng.
  sheet.getRange(NT.MONTH_DATA_ROW,16,numRows,1).setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW,19,numRows,1).setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW,22,numRows,2).setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW,25,numRows,1).setBackground(NT.MANUAL_COLOR);

  const statusRange = sheet.getRange(NT.MONTH_DATA_ROW,10,numRows,1);
  const cccdRange = sheet.getRange(NT.MONTH_DATA_ROW,5,numRows,1);
  const existingRules = sheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getColumn() === 10 || r.getColumn() === 5)
  );

  const statusRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Chưa đăng ký')
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('Hết hạn')
      .setBackground('#ffe599')
      .setFontColor('#7f6000')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('Còn ')
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($D' + NT.MONTH_DATA_ROW + '<>"",$E' + NT.MONTH_DATA_ROW + '="")'
      )
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setBold(true)
      .setRanges([cccdRange])
      .build()
  ];
  sheet.setConditionalFormatRules(existingRules.concat(statusRules));

  const names = sheet.getRange(NT.MONTH_DATA_ROW,4,numRows,1).getDisplayValues();
  const ids = cccdRange.getDisplayValues();
  const notes = [];

  for (let i = 0; i < numRows; i++) {
    const hasName = String(names[i][0] || '').trim() !== '';
    const hasId = String(ids[i][0] || '').trim() !== '';
    notes.push([
      hasName && !hasId
        ? 'CẢNH BÁO: Khách này chưa có CCCD/Hộ chiếu trong dữ liệu nguồn.'
        : ''
    ]);
  }

  cccdRange.setNotes(notes);
  sheet.setFrozenRows(NT.MONTH_HEADER_ROW);
  sheet.setFrozenColumns(0);
  sheet.setHiddenGridlines(true);
}
