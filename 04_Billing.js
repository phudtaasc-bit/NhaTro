/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

function NT_apDungCongThucVaDinhDang_(sheet, monthDate) {
  const lastRow = NT_layDongCuoiDuLieuThang_(sheet);
  if (lastRow < NT.MONTH_DATA_ROW) return;

  const config = NT_layCauHinhThang_(monthDate);
  const numRows = lastRow - NT.MONTH_DATA_ROW + 1;
  const sep = NT_formulaSeparator_();
  const cfgSheet = "'" + NT.CONFIG_SHEET.replace(/'/g, "''") + "'";
  const electricRef = cfgSheet + "!$B$" + config.configRow;
  const waterRef = cfgSheet + "!$C$" + config.configRow;

  const roomValues = sheet.getRange(
    NT.MONTH_DATA_ROW,
    3,
    numRows,
    1
  ).getDisplayValues();

  const noteValues = sheet.getRange(
    NT.MONTH_DATA_ROW,
    25,
    numRows,
    1
  ).getDisplayValues();

  const formulasQ = [];
  const formulasT = [];
  const formulasU = [];
  const formulasX = [];

  for (let i = 0; i < numRows; i++) {
    const row = NT.MONTH_DATA_ROW + i;
    const isLeader = NT_coGiaTri_(roomValues[i][0]);
    const debt = NT_extractDebt_(noteValues[i][0]);

    formulasQ.push([
      isLeader
        ? '=IF(OR(O' + row + '=""' + sep + 'P' + row + '="")' +
          sep + '""' +
          sep + 'MAX(0' + sep + 'P' + row + '-O' + row + ')*' + electricRef + ')'
        : ''
    ]);

    formulasT.push([
      isLeader
        ? '=IF(OR(R' + row + '=""' + sep + 'S' + row + '="")' +
          sep + '""' +
          sep + 'MAX(0' + sep + 'S' + row + '-R' + row + ')*' + waterRef + ')'
        : ''
    ]);

    formulasU.push([
      isLeader
        ? '=SUM(L' + row + sep + 'N' + row + sep + 'Q' + row + sep + 'T' + row + ')+' + debt
        : ''
    ]);

    formulasX.push([
      isLeader
        ? '=IF(U' + row + '=""' +
          sep + '""' +
          sep + 'U' + row + '-IF(V' + row + '=""' + sep + '0' + sep + 'V' + row + '))'
        : ''
    ]);
  }

  sheet.getRange(NT.MONTH_DATA_ROW, 17, numRows, 1).setFormulas(formulasQ);
  sheet.getRange(NT.MONTH_DATA_ROW, 20, numRows, 1).setFormulas(formulasT);
  sheet.getRange(NT.MONTH_DATA_ROW, 21, numRows, 1).setFormulas(formulasU);
  sheet.getRange(NT.MONTH_DATA_ROW, 24, numRows, 1).setFormulas(formulasX);

  sheet.getRange(NT.MONTH_HEADER_ROW, 12).setValue(
    'Thu phí tạm trú/rác (' +
    NT_formatNumber_(config.trashFee) +
    'đ/người)'
  );
  sheet.getRange(NT.MONTH_HEADER_ROW, 17).setValue(
    'Tiền điện (' +
    NT_formatNumber_(config.electricity) +
    'đ/kWh)'
  );
  sheet.getRange(NT.MONTH_HEADER_ROW, 20).setValue(
    'Tiền nước (' +
    NT_formatNumber_(config.water) +
    'đ/số)'
  );

  const paymentValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tiền mặt', 'Chuyển khoản'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(NT.MONTH_DATA_ROW, 23, numRows, 1)
    .setDataValidation(paymentValidation);

  sheet.getRange(NT.MONTH_DATA_ROW, 16, numRows, 1)
    .setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW, 19, numRows, 1)
    .setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW, 22, numRows, 2)
    .setBackground(NT.MANUAL_COLOR);
  sheet.getRange(NT.MONTH_DATA_ROW, 25, numRows, 1)
    .setBackground(NT.MANUAL_COLOR);

  sheet.getRange(NT.MONTH_DATA_ROW, 7, numRows, 3)
    .setNumberFormat(NT.DATE_FORMAT);
  sheet.getRange(NT.MONTH_DATA_ROW, 11, numRows, 1)
    .setNumberFormat(NT.DATE_FORMAT);

  sheet.getRange(NT.MONTH_DATA_ROW, 12, numRows, 14)
    .setNumberFormat(NT.CURRENCY_FORMAT);

  sheet.getRange(NT.MONTH_DATA_ROW, 15, numRows, 6)
    .setNumberFormat('#,##0.##');

  // Dashboard có các vùng hợp nhất đi qua nhiều cột.
  // Chỉ cố định hàng tiêu đề, không cố định cột để tránh lỗi Google Sheets.
  sheet.setFrozenRows(NT.MONTH_HEADER_ROW);
  sheet.setFrozenColumns(0);
}
