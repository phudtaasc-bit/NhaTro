/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

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
  NT_taoTongCongVaDashboard_(targetSheet, monthDate);
  NT_dinhDangBaoCaoThang_(targetSheet);
  targetSheet.activate();
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    (isUpdate ? 'Đã cập nhật ' : 'Đã tạo ') +
    targetName +
    '.\nSố dòng người thuê: ' +
    rows.length
  );
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
      if (!person.residenceDeadline) {
        row[9] = 'Chưa đăng ký';
      } else {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const deadline = new Date(person.residenceDeadline);
        deadline.setHours(0, 0, 0, 0);

        row[9] = deadline < today ? 'Hết hạn' : 'Đã có';
      }
      row[10] = person.residenceDeadline || '';

      if (leader) {
        const cfgRow = config.configRow;
        const cfgSheet = "'" + NT.CONFIG_SHEET.replace(/'/g, "''") + "'";

        if (prorateFee) {
          row[11] =
            '=ROUND(' +
            peopleCount +
            '*' +
            cfgSheet +
            '!$D$' +
            cfgRow +
            '*' +
            occupiedDays +
            '/' +
            daysInMonth +
            NT_formulaSeparator_() +
            '0)';
        } else {
          row[11] =
            '=' +
            peopleCount +
            '*' +
            cfgSheet +
            '!$D$' +
            cfgRow;
        }

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
function NT_luuDuLieuNhapTay_(sheet) {
  const lastRow = NT_layDongCuoiDuLieuThang_(sheet);
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
  const lastRow = NT_layDongCuoiDuLieuThang_(sheet);
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

  sheet.getRange(NT.MONTH_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(NT.MONTH_HEADER_ROW, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground(NT.HEADER_COLOR);

  sheet.setRowHeight(NT.MONTH_HEADER_ROW, 55);
}
function NT_chuanBiSheetThang_(sheet) {
  if (sheet.getMaxColumns() < NT.MONTH_COLS) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      NT.MONTH_COLS - sheet.getMaxColumns()
    );
  }

  const oldHeader = NT_text_(sheet.getRange(1, 1).getDisplayValue()).toUpperCase();
  const newHeader = NT_text_(
    sheet.getRange(NT.MONTH_HEADER_ROW, 1).getDisplayValue()
  ).toUpperCase();

  if (oldHeader === 'STT' && newHeader !== 'STT') {
    sheet.insertRowsBefore(1, NT.MONTH_HEADER_ROW - 1);
  }

  if (
    sheet.getLastColumn() < NT.MONTH_COLS ||
    NT_text_(sheet.getRange(NT.MONTH_HEADER_ROW, 1).getDisplayValue()).toUpperCase() !== 'STT'
  ) {
    NT_taoKhungSheetThang_(sheet);
  }

  const rowsToClear = Math.max(
    sheet.getMaxRows() - NT.MONTH_DATA_ROW + 1,
    1
  );

  const range = sheet.getRange(
    NT.MONTH_DATA_ROW,
    1,
    rowsToClear,
    NT.MONTH_COLS
  );

  range.clearContent();
  range.clearDataValidations();
  range.setBackground(NT.AUTO_COLOR);

  sheet.getRange(1, 1, NT.MONTH_HEADER_ROW - 1, NT.MONTH_COLS)
    .breakApart()
    .clearContent()
    .clearFormat();
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
