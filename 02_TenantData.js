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
function NT_gopVaLoaiTrungKhach_(currentRecords, returnedRecords) {
  const map = new Map();

  currentRecords.forEach(r => {
    map.set(NT_khoaKhach_(r), r);
  });

  returnedRecords.forEach(r => {
    const key = NT_khoaKhach_(r);
    const old = map.get(key);

    if (!old || r.returnDate) {
      map.set(key, r);
    }
  });

  return Array.from(map.values());
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
