/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.3
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

/**
 * Đối soát bảng tháng với sheet TH thuê trọ trước khi tính công thức.
 * Mỗi dòng có tên hoặc CCCD trong sheet nguồn phải có đúng một dòng tương ứng
 * trong bảng tháng, trừ khách đã trả phòng trước tháng báo cáo.
 */
function NT_boSungKhachNguonConThieu_(sheet, monthDate) {
  const monthStart = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  const currentRecords = NT_docSheetDangThue_();
  const returnedRecords = NT_docSheetTraPhong_();

  const normalize = value => NT_text_(value).toUpperCase();
  const personKey = record => [
    NT_chuanHoaPhong_(record.room),
    normalize(record.name),
    normalize(record.id)
  ].join('|');

  // Ngày trả phòng theo đúng phòng + họ tên + CCCD.
  const returnedByKey = {};
  returnedRecords.forEach(record => {
    const key = personKey(record);
    if (!returnedByKey[key] || record.returnDate > returnedByKey[key]) {
      returnedByKey[key] = record.returnDate;
    }
  });

  const lastDataRow = NT_layDongCuoiDuLieuThang_(sheet);
  const existingKeys = new Set();
  const existingRooms = new Set();
  let maxStt = 0;

  if (lastDataRow >= NT.MONTH_DATA_ROW) {
    const values = sheet.getRange(
      NT.MONTH_DATA_ROW,
      1,
      lastDataRow - NT.MONTH_DATA_ROW + 1,
      5
    ).getDisplayValues();

    let currentRoom = '';
    values.forEach(row => {
      const stt = Number(row[0]);
      if (Number.isFinite(stt)) maxStt = Math.max(maxStt, stt);

      if (NT_coGiaTri_(row[2])) {
        currentRoom = NT_chuanHoaPhong_(row[2]);
        existingRooms.add(currentRoom);
      }

      const name = NT_text_(row[3]);
      const id = NT_text_(row[4]);
      if (currentRoom && (name || id)) {
        existingKeys.add([
          currentRoom,
          normalize(name),
          normalize(id)
        ].join('|'));
      }
    });
  }

  const missingRows = [];

  currentRecords.forEach(person => {
    const effective = person.effectiveDate || person.contractDate;
    if (effective && effective > monthEnd) return;

    const key = personKey(person);
    const returnDate = returnedByKey[key] || null;
    if (returnDate && returnDate < monthStart) return;
    if (existingKeys.has(key)) return;

    const room = NT_chuanHoaPhong_(person.room);
    const roomAlreadyShown = existingRooms.has(room);
    const row = new Array(NT.MONTH_COLS).fill('');

    row[0] = ++maxStt;
    row[1] = NT_tangTuPhong_(room);
    row[2] = roomAlreadyShown ? '' : room;
    row[3] = NT_text_(person.name);
    row[4] = NT_text_(person.id);
    row[5] = NT_ghepThuongTru_(person);
    row[6] = person.contractDate || '';
    row[7] = person.effectiveDate || '';
    row[8] = person.contractEnd || '';

    if (!person.residenceDeadline) {
      row[9] = 'Chưa đăng ký';
    } else {
      const deadline = new Date(person.residenceDeadline);
      deadline.setHours(0, 0, 0, 0);
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      row[9] = deadline < today ? 'Hết hạn' : 'Đã có';
    }
    row[10] = person.residenceDeadline || '';

    // Dòng bổ sung trong phòng đã có chỉ dùng để theo dõi khách, không cộng lại
    // tiền phòng, tiền cọc, điện, nước hoặc phí rác của phòng.
    if (roomAlreadyShown) {
      row[24] = 'Bổ sung tự động do đối soát thiếu khách từ sheet TH thuê trọ';
    } else {
      row[24] = 'CẢNH BÁO: Phòng bị thiếu toàn bộ trong bước tổng hợp; cần kiểm tra dữ liệu nguồn';
    }

    missingRows.push(row);
    existingKeys.add(key);
    existingRooms.add(room);
  });

  if (missingRows.length > 0) {
    const startRow = Math.max(lastDataRow + 1, NT.MONTH_DATA_ROW);
    sheet.getRange(startRow, 1, missingRows.length, NT.MONTH_COLS)
      .setValues(missingRows);
  }

  return missingRows.length;
}

function NT_apDungCongThucVaDinhDang_(sheet, monthDate) {
  // Chặn mất khách: đối soát trực tiếp với sheet TH thuê trọ trước khi tính tiền.
  NT_boSungKhachNguonConThieu_(sheet, monthDate);

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

  sheet.setFrozenRows(NT.MONTH_HEADER_ROW);
  sheet.setFrozenColumns(0);
}
