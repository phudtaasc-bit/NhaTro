/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

function NT_docSheetDangThue_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(NT.SOURCE_CURRENT);
  if (!sheet) throw new Error('Không tìm thấy sheet "' + NT.SOURCE_CURRENT + '".');

  const lastRow = sheet.getLastRow();
  if (lastRow < NT.SOURCE_DATA_ROW) return [];

  const values = sheet.getRange(
    NT.SOURCE_DATA_ROW,
    1,
    lastRow - NT.SOURCE_DATA_ROW + 1,
    15
  ).getValues();

  const result = [];
  let carry = {};

  values.forEach((r, index) => {
    if (NT_coGiaTri_(r[0])) {
      carry = {
        room: NT_chuanHoaPhong_(r[0]),
        contractDate: NT_asDate_(r[1]),
        effectiveDate: NT_asDate_(r[2]),
        contractEnd: NT_asDate_(r[3]),
        quantity: NT_asNumber_(r[4]),
        rent: NT_asNumber_(r[5]),
        deposit: NT_asNumber_(r[6])
      };
    }

    const name = NT_text_(r[7]);
    const id = NT_text_(r[10]);
    if (!name && !id) return;
    if (!carry.room) return;

    result.push({
      source: NT.SOURCE_CURRENT,
      sourceRow: NT.SOURCE_DATA_ROW + index,
      room: carry.room,
      contractDate: carry.contractDate,
      effectiveDate: carry.effectiveDate,
      contractEnd: carry.contractEnd,
      returnDate: null,
      quantity: carry.quantity,
      rent: carry.rent,
      deposit: carry.deposit,
      name: name,
      phone: NT_text_(r[8]),
      dob: NT_asDate_(r[9]),
      id: id,
      commune: NT_text_(r[11]),
      district: NT_text_(r[12]),
      province: NT_text_(r[13]),
      residenceDeadline: NT_asDate_(r[14])
    });
  });

  return result;
}

function NT_docSheetTraPhong_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(NT.SOURCE_RETURNED);
  if (!sheet) throw new Error('Không tìm thấy sheet "' + NT.SOURCE_RETURNED + '".');

  const lastRow = sheet.getLastRow();
  if (lastRow < NT.SOURCE_DATA_ROW) return [];

  const values = sheet.getRange(
    NT.SOURCE_DATA_ROW,
    1,
    lastRow - NT.SOURCE_DATA_ROW + 1,
    16
  ).getValues();

  const result = [];
  let carry = {};

  values.forEach((r, index) => {
    if (NT_coGiaTri_(r[0])) {
      carry = {
        room: NT_chuanHoaPhong_(r[0]),
        contractDate: NT_asDate_(r[1]),
        effectiveDate: NT_asDate_(r[2]),
        contractEnd: NT_asDate_(r[3]),
        returnDate: NT_asDate_(r[4]),
        quantity: NT_asNumber_(r[5]),
        rent: NT_asNumber_(r[6]),
        deposit: NT_asNumber_(r[7])
      };
    }

    const name = NT_text_(r[8]);
    const id = NT_text_(r[11]);
    if (!name && !id) return;
    if (!carry.room) return;

    // Chỉ lấy bản ghi trả phòng khi đã có ngày trả phòng cụ thể.
    if (!carry.returnDate) return;

    result.push({
      source: NT.SOURCE_RETURNED,
      sourceRow: NT.SOURCE_DATA_ROW + index,
      room: carry.room,
      contractDate: carry.contractDate,
      effectiveDate: carry.effectiveDate,
      contractEnd: carry.contractEnd,
      returnDate: carry.returnDate,
      quantity: carry.quantity,
      rent: carry.rent,
      deposit: carry.deposit,
      name: name,
      phone: NT_text_(r[9]),
      dob: NT_asDate_(r[10]),
      id: id,
      commune: NT_text_(r[12]),
      district: NT_text_(r[13]),
      province: NT_text_(r[14]),
      residenceDeadline: NT_asDate_(r[15])
    });
  });

  return result;
}

/**
 * Gộp dữ liệu đang thuê và trả phòng mà không làm mất người khi dữ liệu nguồn
 * có hai người cùng phòng vô tình trùng CCCD/Hộ chiếu.
 *
 * Ưu tiên đối chiếu bản ghi trả phòng theo:
 * 1) Số phòng + CCCD/Hộ chiếu + họ tên;
 * 2) Số phòng + CCCD/Hộ chiếu, nhưng chỉ khi tìm thấy đúng một người;
 * 3) Nếu không xác định được thì giữ thành bản ghi riêng.
 */
function NT_gopVaLoaiTrungKhach_(currentRecords, returnedRecords) {
  const result = currentRecords.map(r => Object.assign({}, r));
  const usedReturned = new Set();

  returnedRecords.forEach((returned, returnedIndex) => {
    const returnedRoom = NT_chuanHoaPhong_(returned.room);
    const returnedId = NT_text_(returned.id).toUpperCase();
    const returnedName = NT_text_(returned.name).toUpperCase();

    let matches = result
      .map((record, index) => ({ record: record, index: index }))
      .filter(item => {
        return NT_chuanHoaPhong_(item.record.room) === returnedRoom &&
          NT_text_(item.record.id).toUpperCase() === returnedId &&
          NT_text_(item.record.name).toUpperCase() === returnedName;
      });

    if (matches.length === 0 && returnedId) {
      matches = result
        .map((record, index) => ({ record: record, index: index }))
        .filter(item => {
          return NT_chuanHoaPhong_(item.record.room) === returnedRoom &&
            NT_text_(item.record.id).toUpperCase() === returnedId;
        });
    }

    if (matches.length === 1) {
      result[matches[0].index] = Object.assign({}, returned);
      usedReturned.add(returnedIndex);
    }
  });

  returnedRecords.forEach((returned, index) => {
    if (!usedReturned.has(index)) {
      result.push(Object.assign({}, returned));
    }
  });

  return result;
}

function NT_lapNhomThueTrongThang_(records, monthDate) {
  const monthStart = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  const map = new Map();

  records.forEach(r => {
    const effective = r.effectiveDate || r.contractDate || monthStart;
    const start = NT_maxDate_(effective, monthStart);
    const end = r.returnDate
      ? NT_minDate_(r.returnDate, monthEnd)
      : monthEnd;

    // Khách trả phòng trong tháng vẫn được đưa vào tổng hợp và tính đến ngày trả.
    if (effective > monthEnd) return;
    if (r.returnDate && r.returnDate < monthStart) return;
    if (end < start) return;

    const groupKey = [
      r.room,
      NT_dateKey_(r.contractDate),
      NT_dateKey_(r.effectiveDate),
      NT_dateKey_(r.returnDate),
      r.rent,
      r.deposit
    ].join('|');

    if (!map.has(groupKey)) {
      map.set(groupKey, {
        key: groupKey,
        room: r.room,
        contractDate: r.contractDate,
        effectiveDate: r.effectiveDate,
        contractEnd: r.contractEnd,
        returnDate: r.returnDate,
        rent: r.rent,
        deposit: r.deposit,
        start: start,
        end: end,
        people: []
      });
    }

    map.get(groupKey).people.push(r);
  });

  const groups = Array.from(map.values());

  groups.sort((a, b) => {
    const roomDiff = NT_roomSort_(a.room) - NT_roomSort_(b.room);
    if (roomDiff !== 0) return roomDiff;
    return a.start - b.start;
  });

  const byRoom = {};
  groups.forEach(g => {
    if (!byRoom[g.room]) byRoom[g.room] = [];
    byRoom[g.room].push(g);
  });

  Object.keys(byRoom).forEach(room => {
    const roomGroups = byRoom[room];
    roomGroups.sort((a, b) => a.start - b.start || a.end - b.end);
    roomGroups.forEach((g, i) => {
      g.isLastGroupOfRoom = i === roomGroups.length - 1;
      g.groupIndexInRoom = i + 1;
      g.groupCountInRoom = roomGroups.length;
    });
  });

  return groups;
}

/**
 * Đếm số phòng trống tại ngày cuối tháng báo cáo.
 * Phòng trống khi không còn bất kỳ khách nào có thời gian thuê bao phủ ngày cuối tháng.
 */
function NT_demPhongTrongCuoiThang_(monthDate) {
  const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  const currentRecords = NT_docSheetDangThue_();
  const returnedRecords = NT_docSheetTraPhong_();
  const mergedRecords = NT_gopVaLoaiTrungKhach_(currentRecords, returnedRecords);

  const allRooms = new Set();
  currentRecords.forEach(r => r.room && allRooms.add(r.room));
  returnedRecords.forEach(r => r.room && allRooms.add(r.room));

  const occupiedRooms = new Set();

  mergedRecords.forEach(r => {
    if (!r.room) return;

    const effective = r.effectiveDate || r.contractDate;
    const startedByMonthEnd = !effective || effective <= monthEnd;
    const notReturnedByMonthEnd = !r.returnDate || r.returnDate > monthEnd;

    if (startedByMonthEnd && notReturnedByMonthEnd) {
      occupiedRooms.add(r.room);
    }
  });

  let vacant = 0;
  allRooms.forEach(room => {
    if (!occupiedRooms.has(room)) vacant++;
  });

  return vacant;
}
