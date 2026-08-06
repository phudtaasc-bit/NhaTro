/**
 * Đọc chỉ số điện, nước theo tháng và nhận diện phòng trống từ dữ liệu nguồn.
 */
function NT_docChiSoDienNuocThang_(monthDate) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(NT.METER_SHEET);
  if (!sheet || sheet.getLastRow() < 3 || sheet.getLastColumn() < 2) return {};

  const lastCol = sheet.getLastColumn();
  const headerYear = sheet.getRange(1, 1, 2, lastCol).getDisplayValues();
  const targetYear = monthDate.getFullYear();
  const targetMonth = monthDate.getMonth() + 1;
  const prevDate = new Date(targetYear, monthDate.getMonth() - 1, 1);

  let section = '';
  let sectionYear = null;
  let electricOldCol = 0;
  let electricNewCol = 0;
  let waterOldCol = 0;
  let waterNewCol = 0;

  for (let col = 2; col <= lastCol; col++) {
    const yearText = NT_text_(headerYear[0][col - 1]);
    if (yearText) {
      const electricMatch = /Số\s*điện\s*năm\s*(\d{4})/i.exec(yearText);
      const waterMatch = /Số\s*nước\s*năm\s*(\d{4})/i.exec(yearText);
      if (electricMatch) {
        section = 'ELECTRIC';
        sectionYear = Number(electricMatch[1]);
      } else if (waterMatch) {
        section = 'WATER';
        sectionYear = Number(waterMatch[1]);
      }
    }

    const monthMatch = /Tháng\s*(\d{1,2})/i.exec(NT_text_(headerYear[1][col - 1]));
    if (!monthMatch) continue;
    const monthNumber = Number(monthMatch[1]);

    if (section === 'ELECTRIC') {
      if (sectionYear === targetYear && monthNumber === targetMonth) electricNewCol = col;
      if (sectionYear === prevDate.getFullYear() && monthNumber === prevDate.getMonth() + 1) electricOldCol = col;
    }

    if (section === 'WATER') {
      if (sectionYear === targetYear && monthNumber === targetMonth) waterNewCol = col;
      if (sectionYear === prevDate.getFullYear() && monthNumber === prevDate.getMonth() + 1) waterOldCol = col;
    }
  }

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getDisplayValues();
  const result = {};

  values.forEach(row => {
    const room = NT_chuanHoaPhong_(row[0]);
    if (!room) return;

    result[room] = {
      oldElectric: electricOldCol ? NT_meterValue_(row[electricOldCol - 1]) : '',
      newElectric: electricNewCol ? NT_meterValue_(row[electricNewCol - 1]) : '',
      oldWater: waterOldCol ? NT_meterValue_(row[waterOldCol - 1]) : '',
      newWater: waterNewCol ? NT_meterValue_(row[waterNewCol - 1]) : ''
    };
  });

  return result;
}

function NT_meterValue_(value) {
  const text = NT_text_(value);
  return text === '' ? '' : NT_asNumber_(text);
}

/**
 * Đọc toàn bộ số phòng đã khai báo tại sheet TH thuê trọ, kể cả phòng chưa có khách.
 */
function NT_docTrangThaiPhongNguon_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(NT.SOURCE_CURRENT);
  if (!sheet || sheet.getLastRow() < NT.SOURCE_DATA_ROW) {
    return { allRooms: new Set(), occupiedRooms: new Set(), vacantRooms: new Set() };
  }

  const range = sheet.getRange(
    NT.SOURCE_DATA_ROW,
    1,
    sheet.getLastRow() - NT.SOURCE_DATA_ROW + 1,
    11
  );
  const display = range.getDisplayValues();
  const roomHasPerson = {};
  let currentRoom = '';

  display.forEach(row => {
    if (NT_coGiaTri_(row[0])) {
      currentRoom = NT_chuanHoaPhong_(row[0]);
      if (currentRoom && roomHasPerson[currentRoom] === undefined) roomHasPerson[currentRoom] = false;
    }
    if (!currentRoom) return;

    const name = NT_text_(row[7]);
    const id = NT_text_(row[10]);
    if (name || id) roomHasPerson[currentRoom] = true;
  });

  const allRooms = new Set(Object.keys(roomHasPerson));
  const occupiedRooms = new Set();
  const vacantRooms = new Set();

  Object.keys(roomHasPerson).forEach(room => {
    if (roomHasPerson[room]) occupiedRooms.add(room);
    else vacantRooms.add(room);
  });

  return { allRooms: allRooms, occupiedRooms: occupiedRooms, vacantRooms: vacantRooms };
}
