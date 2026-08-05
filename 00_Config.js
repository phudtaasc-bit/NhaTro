/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

const NT = {
  SOURCE_CURRENT: 'TH thuê trọ',
  SOURCE_RETURNED: 'Trả phòng',
  CONFIG_SHEET: 'CẤU HÌNH',
  MONTH_PREFIX: 'Tháng ',
  SOURCE_HEADER_ROW: 3,
  SOURCE_DATA_ROW: 4,
  MONTH_HEADER_ROW: 6,
  MONTH_DATA_ROW: 7,
  MONTH_COLS: 25,
  MANUAL_COLOR: '#fff2cc',
  AUTO_COLOR: '#ffffff',
  HEADER_COLOR: '#d9eaf7',
  CURRENCY_FORMAT: '#,##0',
  DATE_FORMAT: 'd/m/yyyy',
  MONTH_FORMAT: 'm/yyyy'
};

function NT_taoHoacCapNhatCauHinh() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(NT.CONFIG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(NT.CONFIG_SHEET);
  }

  const headers = [
    'Tháng áp dụng',
    'Đơn giá điện/kWh',
    'Đơn giá nước/số',
    'Phí tạm trú/rác/người/tháng',
    'Phân bổ phí rác theo ngày',
    'Ghi chú'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(NT.HEADER_COLOR)
    .setWrap(true);

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 1, 6).setValues([[
      new Date(2026, 7, 1),
      3000,
      10000,
      30000,
      'KHÔNG',
      'Mức tạm tính tháng 8/2026'
    ]]);
  }

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['KHÔNG', 'CÓ'], true)
    .setAllowInvalid(false)
    .build();

  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 5, maxRows, 1).setDataValidation(validation);
  sheet.getRange(2, 1, maxRows, 1).setNumberFormat(NT.MONTH_FORMAT);
  sheet.getRange(2, 2, maxRows, 3).setNumberFormat(NT.CURRENCY_FORMAT);

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 190);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 240);
}
function NT_layCauHinhThang_(monthDate) {
  NT_taoHoacCapNhatCauHinh();

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(NT.CONFIG_SHEET);
  const lastRow = sheet.getLastRow();

  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    : [];

  const targetKey = NT_monthKey_(monthDate);
  let exact = null;
  let nearest = null;

  values.forEach((r, index) => {
    const d = NT_asDate_(r[0]);
    if (!d) return;

    const item = {
      configRow: index + 2,
      month: new Date(d.getFullYear(), d.getMonth(), 1),
      electricity: NT_asNumber_(r[1]) || 3000,
      water: NT_asNumber_(r[2]) || 10000,
      trashFee: NT_asNumber_(r[3]) || 30000,
      prorateFee: NT_text_(r[4]) || 'KHÔNG'
    };

    if (NT_monthKey_(item.month) === targetKey) exact = item;

    if (item.month <= monthDate) {
      if (!nearest || item.month > nearest.month) nearest = item;
    }
  });

  if (exact) return exact;

  const base = nearest || {
    month: monthDate,
    electricity: 3000,
    water: 10000,
    trashFee: 30000,
    prorateFee: 'KHÔNG'
  };

  sheet.appendRow([
    new Date(monthDate.getFullYear(), monthDate.getMonth(), 1),
    base.electricity,
    base.water,
    base.trashFee,
    base.prorateFee,
    'Tự động kế thừa mức gần nhất'
  ]);

  const newRow = sheet.getLastRow();

  return {
    configRow: newRow,
    month: monthDate,
    electricity: base.electricity,
    water: base.water,
    trashFee: base.trashFee,
    prorateFee: base.prorateFee
  };
}
