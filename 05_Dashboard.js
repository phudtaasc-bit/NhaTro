/**
 * Dashboard và dòng tổng cộng cuối bảng.
 */
function NT_taoTongCongVaDashboard_(sheet, monthDate) {
  const lastDataRow = NT_layDongCuoiDuLieuThang_(sheet);
  if (lastDataRow < NT.MONTH_DATA_ROW) return;

  const sep = NT_formulaSeparator_();
  const totalRow = lastDataRow + 1;

  sheet.getRange(totalRow,1,1,NT.MONTH_COLS).breakApart().clearContent().clearFormat();
  sheet.getRange(totalRow,1,1,11).merge();
  sheet.getRange(totalRow,1).setValue('TỔNG CỘNG');

  [12,13,14,17,20,21,22,24].forEach(col => {
    const letter = NT_colToLetter_(col);
    sheet.getRange(totalRow,col).setFormula('=SUM(' + letter + NT.MONTH_DATA_ROW + ':' + letter + lastDataRow + ')');
  });

  sheet.getRange(totalRow,1,1,NT.MONTH_COLS)
    .setFontFamily('Times New Roman').setFontSize(11).setFontWeight('bold')
    .setVerticalAlignment('middle').setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.getRange(totalRow,1,1,11).setHorizontalAlignment('center');
  [12,13,14,17,20,21,22,24].forEach(col => {
    sheet.getRange(totalRow,col).setHorizontalAlignment('right').setNumberFormat(NT.CURRENCY_FORMAT);
  });
  sheet.getRange(totalRow,1,1,NT.MONTH_COLS)
    .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.DOUBLE);

  sheet.getRange(1,1,4,NT.MONTH_COLS).breakApart().clearContent().clearFormat();
  sheet.getRange(1,1,1,NT.MONTH_COLS).merge();
  sheet.getRange(1,1).setValue('BÁO CÁO TỔNG HỢP THUÊ TRỌ - ' + NT_tenSheetThang_(monthDate).toUpperCase())
    .setFontFamily('Times New Roman').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const dashboard = [
    ['Số phòng phát sinh','=COUNTUNIQUE(FILTER(C'+NT.MONTH_DATA_ROW+':C'+lastDataRow+sep+'C'+NT.MONTH_DATA_ROW+':C'+lastDataRow+'<>""))','Số khách','=COUNTA(D'+NT.MONTH_DATA_ROW+':D'+lastDataRow+')','Tiền phòng','=N'+totalRow,'Tiền điện','=Q'+totalRow],
    ['Phí tạm trú/rác','=L'+totalRow,'Tiền nước','=T'+totalRow,'Tổng phải thu','=U'+totalRow,'Đã thu','=V'+totalRow],
    ['Còn phải thu','=X'+totalRow,'Tiền cọc theo dõi','=M'+totalRow,'Khách chưa đăng ký','=COUNTIF(J'+NT.MONTH_DATA_ROW+':J'+lastDataRow+sep+'"Chưa đăng ký")','Khách hết hạn','=COUNTIF(J'+NT.MONTH_DATA_ROW+':J'+lastDataRow+sep+'"Hết hạn")']
  ];
  const starts = [1,7,13,19];
  for (let r=0;r<dashboard.length;r++) {
    for (let block=0;block<4;block++) {
      const startCol = starts[block];
      sheet.getRange(r+2,startCol,1,3).merge().setValue(dashboard[r][block*2]);
      sheet.getRange(r+2,startCol+3,1,3).merge().setFormula(dashboard[r][block*2+1]);
    }
  }
  sheet.getRange(2,1,3,24)
    .setFontFamily('Times New Roman').setFontSize(11).setVerticalAlignment('middle')
    .setBorder(true,true,true,true,true,true,'#7f8c8d',SpreadsheetApp.BorderStyle.SOLID);
  starts.forEach(startCol => {
    sheet.getRange(2,startCol,3,3).setBackground('#d9e2f3').setFontWeight('bold').setHorizontalAlignment('left');
    sheet.getRange(2,startCol+3,3,3).setBackground('#ffffff').setFontWeight('bold').setHorizontalAlignment('right');
  });
  [[2,16],[2,22],[3,4],[3,10],[3,16],[3,22],[4,4],[4,10]].forEach(pos => {
    sheet.getRange(pos[0],pos[1],1,3).setNumberFormat(NT.CURRENCY_FORMAT);
  });
}
