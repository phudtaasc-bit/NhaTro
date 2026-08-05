/**
 * QUẢN LÝ NHÀ TRỌ - GOOGLE APPS SCRIPT
 * Phiên bản module hóa: 1.2
 * Sheet nguồn: TH thuê trọ, Trả phòng
 * Sheet tháng: Tháng M.YYYY
 */

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
function NT_formulaSeparator_() {
  const locale = String(SpreadsheetApp.getActive().getSpreadsheetLocale() || '')
    .toLowerCase();

  return (
    locale.indexOf('vi') === 0 ||
    locale.indexOf('de') === 0 ||
    locale.indexOf('fr') === 0 ||
    locale.indexOf('es') === 0 ||
    locale.indexOf('it') === 0 ||
    locale.indexOf('pt') === 0 ||
    locale.indexOf('ru') === 0
  ) ? ';' : ',';
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

function NT_layDongCuoiDuLieuThang_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < NT.MONTH_DATA_ROW) return NT.MONTH_DATA_ROW - 1;

  const values = sheet.getRange(
    NT.MONTH_DATA_ROW,
    1,
    lastRow - NT.MONTH_DATA_ROW + 1,
    1
  ).getValues();

  let lastDataRow = NT.MONTH_DATA_ROW - 1;

  values.forEach((row, index) => {
    const value = row[0];
    if (typeof value === 'number' && Number.isFinite(value)) {
      lastDataRow = NT.MONTH_DATA_ROW + index;
    }
  });

  return lastDataRow;
}

function NT_colToLetter_(column) {
  let result = '';
  let n = column;

  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
}
