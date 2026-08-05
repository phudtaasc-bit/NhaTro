/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản: 1.0
 *
 * Cấu trúc sheet nguồn:
 * 1) "TH thuê trọ": hàng tiêu đề 3, dữ liệu từ hàng 4
 *    A Số phòng | B Ngày ký HĐ | C Ngày hiệu lực | D Ngày hết hạn
 *    E Số lượng | F Giá phòng | G Đặt cọc | H Họ và tên | I SĐT
 *    J BOD | K CCCD/Hộ chiếu | L Xã | M Huyện | N Tỉnh | O Hạn đăng ký tạm trú
 *
 * 2) "Trả phòng": hàng tiêu đề 3, dữ liệu từ hàng 4
 *    A Số phòng | B Ngày ký HĐ | C Ngày hiệu lực | D Ngày hết hạn
 *    E Ngày trả phòng | F Số lượng | G Giá phòng | H Đặt cọc
 *    I Họ và tên | J SĐT | K BOD | L CCCD/Hộ chiếu
 *    M Xã | N Huyện | O Tỉnh | P Hạn đăng ký tạm trú
 *
 * 3) Sheet tháng: "Tháng 8.2026", "Tháng 9.2026", ...
 *    A:Y theo đúng kết cấu hiện tại.
 */

const NT = {
  SOURCE_CURRENT: 'TH thuê trọ',
  SOURCE_RETURNED: 'Trả phòng',
  CONFIG_SHEET: 'CẤU HÌNH',
  MONTH_PREFIX: 'Tháng ',
  SOURCE_HEADER_ROW: 3,
  SOURCE_DATA_ROW: 4,
  MONTH_HEADER_ROW: 1,
  MONTH_DATA_ROW: 2,
  MONTH_COLS: 25,
  MANUAL_COLOR: '#fff2cc',
  AUTO_COLOR: '#ffffff',
  HEADER_COLOR: '#d9eaf7',
  CURRENCY_FORMAT: '#,##0',
  DATE_FORMAT: 'd/m/yyyy',
  MONTH_FORMAT: 'm/yyyy'
};

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
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã tính lại tiền điện, nước và công nợ.');
}

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

function NT_taoHoacCapNhatSheetThang_(monthDate, isUpdate) {
  const ss = SpreadsheetApp.getActive();
  const targetName = NT_tenSheetThang_(monthDate);
  let targetSheet = ss.getSheetByName(targetName);

  const manualSnapshot = targetSheet
    ? NT_luuDuLieuNhapTay_(targetSheet)
    : { byPerson: {}, byRoom: {} };

  if (!targetSheet) {
    const template = NT_timSheetThangMau_(ss, monthDate);
    if (template) {
      targetSheet = template.copyTo(ss).setName(targetName);
    } else {
      targetSheet = ss.insertSheet(targetName);
      NT_taoKhungSheetThang_(targetSheet);
    }
  } else if (!isUpdate) {
    SpreadsheetApp.getUi().alert('Sheet "' + targetName + '" đã tồn tại.');
    return;
  }

  NT_chuanBiSheetThang_(targetSheet);

  const currentRecords = NT_docSheetDangThue_();
  const returnedRecords = NT_docSheetTraPhong_();
  const allRecords = NT_gopVaLoaiTrungKhach_(currentRecords, returnedRecords);
  const groups = NT_lapNhomThueTrongThang_(allRecords, monthDate);

  const prevSheet = NT_timSheetThangTruoc_(ss, monthDate);
  const prevState = prevSheet
    ? NT_docTrangThaiThangTruoc_(prevSheet)
    : {};

  const config = NT_layCauHinhThang_(monthDate);
  const rows = NT_taoDuLieuThang_(
    groups,
    monthDate,
    config,
    prevState,
    manualSnapshot
  );

  if (rows.length > 0) {
    targetSheet.getRange(
      NT.MONTH_DATA_ROW,
      1,
      rows.length,
      NT.MONTH_COLS
    ).setValues(rows);
  }

  NT_apDungCongThucVaDinhDang_(targetSheet, monthDate);
  targetSheet.activate();
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    (isUpdate ? 'Đã cập nhật ' : 'Đã tạo ') +
    targetName +
    '.\nSố dòng người thuê: ' +
    rows.length
  );
}

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

function NT_taoDuLieuThang_(groups, monthDate, config, prevState, manualSnapshot) {
  const daysInMonth = new Date(
    monthDate.getFullYear(),
    monthDate.getMonth() + 1,
    0
  ).getDate();

  const rows = [];
  let stt = 1;

  groups.forEach(group => {
    const occupiedDays = NT_daysInclusive_(group.start, group.end);
    const proratedRent = Math.round(
      (group.rent || 0) * occupiedDays / daysInMonth
    );

    const peopleCount = group.people.length;
    const prorateFee = String(config.prorateFee).toUpperCase() === 'CÓ';
    const trashFee = prorateFee
      ? Math.round(config.trashFee * peopleCount * occupiedDays / daysInMonth)
      : config.trashFee * peopleCount;

    const roomPrev = prevState[group.room] || {};
    const roomManual = manualSnapshot.byRoom[group.room] || {};

    const oldElectric = group.isLastGroupOfRoom
      ? NT_firstNumber_(
          roomPrev.nextElectric,
          roomPrev.oldElectric,
          roomManual.oldElectric
        )
      : '';

    const oldWater = group.isLastGroupOfRoom
      ? NT_firstNumber_(
          roomPrev.nextWater,
          roomPrev.oldWater,
          roomManual.oldWater
        )
      : '';

    const carryDebt = group.isLastGroupOfRoom
      ? NT_asNumber_(roomPrev.debt)
      : 0;

    group.people.forEach((person, personIndex) => {
      const leader = personIndex === 0;
      const personKey = NT_khoaDongThang_(group.room, person.id, person.name);
      const saved = manualSnapshot.byPerson[personKey] || {};

      const row = new Array(NT.MONTH_COLS).fill('');

      row[0] = stt++;
      row[1] = NT_tangTuPhong_(group.room);
      row[2] = leader ? group.room : '';
      row[3] = person.name;
      row[4] = person.id;
      row[5] = NT_ghepThuongTru_(person);
      row[6] = person.contractDate || '';
      row[7] = person.effectiveDate || '';
      row[8] = person.contractEnd || '';
      row[9] = person.residenceDeadline ? 'Đã có hạn đăng ký' : 'Chưa rõ';
      row[10] = person.residenceDeadline || '';

      if (leader) {
        row[11] = trashFee;
        row[12] = group.deposit || '';
        row[13] = proratedRent;

        if (group.isLastGroupOfRoom) {
          row[14] = oldElectric;
          row[15] = NT_coGiaTri_(saved.newElectric)
            ? saved.newElectric
            : '';
          row[17] = oldWater;
          row[18] = NT_coGiaTri_(saved.newWater)
            ? saved.newWater
            : '';
        }

        row[21] = NT_coGiaTri_(saved.paid) ? saved.paid : '';
        row[22] = saved.paymentMethod || '';
      }

      row[24] = NT_ghepGhiChu_(
        saved.note,
        group,
        occupiedDays,
        daysInMonth,
        carryDebt,
        leader
      );

      rows.push(row);
    });
  });

  return rows;
}

function NT_apDungCongThucVaDinhDang_(sheet, monthDate) {
  const lastRow = Math.max(sheet.getLastRow(), NT.MONTH_DATA_ROW);
  if (lastRow < NT.MONTH_DATA_ROW) return;

  const config = NT_layCauHinhThang_(monthDate);
  const numRows = lastRow - NT.MONTH_DATA_ROW + 1;

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
        ? '=IF(OR(O' + row + '="",P' + row + '=""),"",MAX(0,P' + row + '-O' + row + ')*' + config.electricity + ')'
        : ''
    ]);

    formulasT.push([
      isLeader
        ? '=IF(OR(R' + row + '="",S' + row + '=""),"",MAX(0,S' + row + '-R' + row + ')*' + config.water + ')'
        : ''
    ]);

    formulasU.push([
      isLeader
        ? '=SUM(L' + row + ',N' + row + ',Q' + row + ',T' + row + ')+' + debt
        : ''
    ]);

    formulasX.push([
      isLeader
        ? '=IF(U' + row + '="","",U' + row + '-IF(V' + row + '="",0,V' + row + '))'
        : ''
    ]);
  }

  sheet.getRange(NT.MONTH_DATA_ROW, 17, numRows, 1).setFormulas(formulasQ);
  sheet.getRange(NT.MONTH_DATA_ROW, 20, numRows, 1).setFormulas(formulasT);
  sheet.getRange(NT.MONTH_DATA_ROW, 21, numRows, 1).setFormulas(formulasU);
  sheet.getRange(NT.MONTH_DATA_ROW, 24, numRows, 1).setFormulas(formulasX);

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

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
}

function NT_luuDuLieuNhapTay_(sheet) {
  const lastRow = sheet.getLastRow();
  const result = { byPerson: {}, byRoom: {} };
  if (lastRow < NT.MONTH_DATA_ROW) return result;

  const values = sheet.getRange(
    NT.MONTH_DATA_ROW,
    1,
    lastRow - NT.MONTH_DATA_ROW + 1,
    NT.MONTH_COLS
  ).getValues();

  let currentRoom = '';

  values.forEach(r => {
    if (NT_coGiaTri_(r[2])) currentRoom = NT_chuanHoaPhong_(r[2]);
    if (!currentRoom) return;

    const personKey = NT_khoaDongThang_(
      currentRoom,
      NT_text_(r[4]),
      NT_text_(r[3])
    );

    result.byPerson[personKey] = {
      newElectric: r[15],
      newWater: r[18],
      paid: r[21],
      paymentMethod: NT_text_(r[22]),
      note: NT_text_(r[24])
    };

    if (!result.byRoom[currentRoom]) result.byRoom[currentRoom] = {};

    if (NT_coGiaTri_(r[14])) result.byRoom[currentRoom].oldElectric = r[14];
    if (NT_coGiaTri_(r[15])) result.byRoom[currentRoom].newElectric = r[15];
    if (NT_coGiaTri_(r[17])) result.byRoom[currentRoom].oldWater = r[17];
    if (NT_coGiaTri_(r[18])) result.byRoom[currentRoom].newWater = r[18];
  });

  return result;
}

function NT_docTrangThaiThangTruoc_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < NT.MONTH_DATA_ROW) return map;

  const values = sheet.getRange(
    NT.MONTH_DATA_ROW,
    1,
    lastRow - NT.MONTH_DATA_ROW + 1,
    NT.MONTH_COLS
  ).getValues();

  let currentRoom = '';

  values.forEach(r => {
    if (NT_coGiaTri_(r[2])) currentRoom = NT_chuanHoaPhong_(r[2]);
    if (!currentRoom) return;

    if (!map[currentRoom]) {
      map[currentRoom] = {
        oldElectric: '',
        nextElectric: '',
        oldWater: '',
        nextWater: '',
        debt: 0
      };
    }

    if (NT_coGiaTri_(r[14])) map[currentRoom].oldElectric = r[14];
    if (NT_coGiaTri_(r[15])) map[currentRoom].nextElectric = r[15];
    if (NT_coGiaTri_(r[17])) map[currentRoom].oldWater = r[17];
    if (NT_coGiaTri_(r[18])) map[currentRoom].nextWater = r[18];

    const debt = NT_asNumber_(r[23]);
    if (debt) map[currentRoom].debt += debt;
  });

  return map;
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

  values.forEach(r => {
    const d = NT_asDate_(r[0]);
    if (!d) return;

    const item = {
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

  return {
    month: monthDate,
    electricity: base.electricity,
    water: base.water,
    trashFee: base.trashFee,
    prorateFee: base.prorateFee
  };
}

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

function NT_taoKhungSheetThang_(sheet) {
  const headers = [
    'STT',
    'Tầng',
    'Phòng',
    'Tên người thuê',
    'CCCD',
    'Thường trú',
    'Ngày hợp đồng',
    'Ngày hiệu lực',
    'Ngày hết hạn',
    'Đăng ký tạm trú',
    'Hạn đăng ký tạm trú',
    'Thu phí tạm trú/rác',
    'Tiền cọc',
    'Tiền phòng',
    'Số điện cũ',
    'Số điện mới',
    'Tiền điện (3k)',
    'Số nước cũ',
    'Số nước mới',
    'Tiền nước (10K)',
    'Tiền phải thu',
    'Đã nộp',
    'Hình thức nạp',
    'Tiền còn nợ',
    'Ghi chú'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground(NT.HEADER_COLOR);

  sheet.setRowHeight(1, 55);
}

function NT_chuanBiSheetThang_(sheet) {
  if (sheet.getMaxColumns() < NT.MONTH_COLS) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      NT.MONTH_COLS - sheet.getMaxColumns()
    );
  }

  if (sheet.getLastColumn() < NT.MONTH_COLS ||
      !NT_coGiaTri_(sheet.getRange(1, 1).getValue())) {
    NT_taoKhungSheetThang_(sheet);
  }

  const rowsToClear = Math.max(sheet.getMaxRows() - 1, 1);
  const range = sheet.getRange(
    NT.MONTH_DATA_ROW,
    1,
    rowsToClear,
    NT.MONTH_COLS
  );

  range.clearContent();
  range.clearDataValidations();
  range.setBackground(NT.AUTO_COLOR);
}

function NT_timSheetThangMau_(ss, targetMonth) {
  const previous = NT_timSheetThangTruoc_(ss, targetMonth);
  if (previous) return previous;

  const latest = NT_timSheetThangMoiNhat_(ss);
  return latest ? latest.sheet : null;
}

function NT_timSheetThangMoiNhat_(ss) {
  const found = [];

  ss.getSheets().forEach(sheet => {
    const d = NT_layThangTuTenSheet_(sheet.getName());
    if (d) {
      found.push({
        sheet: sheet,
        date: d,
        year: d.getFullYear(),
        month: d.getMonth()
      });
    }
  });

  if (found.length === 0) return null;
  found.sort((a, b) => b.date - a.date);
  return found[0];
}

function NT_timSheetThangTruoc_(ss, monthDate) {
  const prev = new Date(monthDate.getFullYear(), monthDate.getMonth() - 1, 1);
  return ss.getSheetByName(NT_tenSheetThang_(prev));
}

function NT_layThangTuTenSheet_(name) {
  const match = /^Tháng\s+(\d{1,2})\.(\d{4})$/i.exec(String(name).trim());
  if (!match) return null;

  const month = Number(match[1]);
  const year = Number(match[2]);

  if (month < 1 || month > 12 || year < 2000) return null;
  return new Date(year, month - 1, 1);
}

function NT_tenSheetThang_(date) {
  return NT.MONTH_PREFIX + (date.getMonth() + 1) + '.' + date.getFullYear();
}

function NT_parseMonthInput_(text) {
  const match = /^(\d{1,2})[\/\.\-](\d{4})$/.exec(String(text).trim());
  if (!match) return null;

  const month = Number(match[1]);
  const year = Number(match[2]);

  if (month < 1 || month > 12 || year < 2000) return null;
  return new Date(year, month - 1, 1);
}

function NT_tangTuPhong_(room) {
  const n = parseInt(String(room).replace(/[^\d]/g, ''), 10);
  if (!Number.isFinite(n)) return '';
  return Math.floor(n / 100);
}

function NT_ghepThuongTru_(person) {
  return [person.commune, person.district, person.province]
    .filter(Boolean)
    .join(', ');
}

function NT_khoaKhach_(r) {
  const id = NT_text_(r.id).toUpperCase();
  if (id) return r.room + '|ID|' + id;

  return [
    r.room,
    'FALLBACK',
    NT_text_(r.name).toUpperCase(),
    NT_dateKey_(r.dob)
  ].join('|');
}

function NT_khoaDongThang_(room, id, name) {
  const normalizedId = NT_text_(id).toUpperCase();
  return normalizedId
    ? room + '|ID|' + normalizedId
    : room + '|NAME|' + NT_text_(name).toUpperCase();
}

function NT_chuanHoaPhong_(value) {
  if (value === null || value === undefined || value === '') return '';
  return String(value).trim().replace(/\.0$/, '');
}

function NT_roomSort_(room) {
  const n = parseInt(String(room).replace(/[^\d]/g, ''), 10);
  return Number.isFinite(n) ? n : 999999;
}

function NT_ghepGhiChu_(savedNote, group, occupiedDays, daysInMonth, debt, leader) {
  const parts = [];

  if (leader) {
    parts.push(
      'Tính tiền từ ' +
      NT_formatDate_(group.start) +
      ' đến ' +
      NT_formatDate_(group.end) +
      ' (' +
      occupiedDays +
      '/' +
      daysInMonth +
      ' ngày)'
    );

    if (group.groupCountInRoom > 1) {
      parts.push(
        'Đợt thuê ' +
        group.groupIndexInRoom +
        '/' +
        group.groupCountInRoom
      );
    }

    if (debt) {
      parts.push('Nợ chuyển tháng trước: ' + NT_formatNumber_(debt));
    }
  }

  if (savedNote) {
    const cleaned = String(savedNote)
      .replace(/Nợ chuyển tháng trước:\s*[\d\.,]+/gi, '')
      .replace(/Tính tiền từ.*?\(\d+\/\d+\s+ngày\)/gi, '')
      .replace(/Đợt thuê\s+\d+\/\d+/gi, '')
      .replace(/\s*;\s*;\s*/g, '; ')
      .replace(/^\s*;\s*|\s*;\s*$/g, '')
      .trim();

    if (cleaned) parts.push(cleaned);
  }

  return parts.join('; ');
}

function NT_extractDebt_(note) {
  const match = /Nợ chuyển tháng trước:\s*([\d\.,]+)/i.exec(String(note || ''));
  if (!match) return 0;

  return Number(String(match[1]).replace(/[^\d\-]/g, '')) || 0;
}

function NT_monthKey_(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

function NT_dateKey_(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function NT_formatDate_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}

function NT_formatNumber_(n) {
  return Number(n || 0).toLocaleString('vi-VN');
}

function NT_asDate_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!NT_coGiaTri_(value)) return null;

  const text = String(value).trim();
  const match = /^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/.exec(text);

  if (match) {
    const d = new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1])
    );
    return isNaN(d) ? null : d;
  }

  const parsed = new Date(value);
  return isNaN(parsed)
    ? null
    : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function NT_asNumber_(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (!NT_coGiaTri_(value)) return 0;

  const normalized = String(value)
    .trim()
    .replace(/\s/g, '')
    .replace(/[^\d,\.\-]/g, '');

  if (!normalized) return 0;

  const cleaned = normalized
    .replace(/\.(?=\d{3}(\D|$))/g, '')
    .replace(',', '.');

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function NT_firstNumber_() {
  for (let i = 0; i < arguments.length; i++) {
    if (NT_coGiaTri_(arguments[i])) return NT_asNumber_(arguments[i]);
  }
  return '';
}

function NT_text_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function NT_coGiaTri_(value) {
  return value !== null && value !== undefined && value !== '';
}

function NT_maxDate_(a, b) {
  return a > b ? a : b;
}

function NT_minDate_(a, b) {
  return a < b ? a : b;
}

function NT_daysInclusive_(start, end) {
  const ms = 24 * 60 * 60 * 1000;
  return Math.floor((end - start) / ms) + 1;
}