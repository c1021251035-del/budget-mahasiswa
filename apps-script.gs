/**
 * STEEP BUDGET — Google Sheets Sync
 * ===================================
 * Cara deploy:
 * 1. Buka https://sheets.new → bikin spreadsheet baru
 * 2. Extensions > Apps Script
 * 3. Hapus code default, paste isi file ini
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Access: Anyone with link
 * 5. Klik Deploy, copy URL web app
 * 6. Tempel URL di budget app → "Google Sheets Settings"
 *
 * Setiap sync: data lama dihapus, ditulis ulang.
 * 3 sheet: Ringkasan, Detail, Riwayat
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const monthLabel = data.month || '';
    const income = data.income || 0;
    const groups = data.groups || [];
    const history = data.history || [];

    // === SHEET 1: RINGKASAN ===
    let s1 = ss.getSheetByName('Ringkasan');
    if (s1) s1.clear(); else s1 = ss.insertSheet('Ringkasan');
    s1.getRange('A1:D1').merge()
      .setValue('STEEP BUDGET — LAPORAN BULANAN')
      .setFontSize(16).setFontColor('#5d2a1a').setFontFamily('Georgia')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    s1.getRange('A2:D2').merge()
      .setValue(monthLabel)
      .setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center');

    s1.getRange('A4:B4').setValues([['Penghasilan', income]]);
    s1.getRange('A4').setFontWeight('bold');
    s1.getRange('B4').setNumberFormat('#,##0');

    const hdrRow = 6;
    s1.getRange(hdrRow, 1, 1, 4).setValues([['Kategori', 'Anggaran', 'Realisasi', 'Selisih']])
      .setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#17191c').setHorizontalAlignment('center');

    groups.forEach((g, i) => {
      const r = hdrRow + 1 + i;
      const diff = g.budget - g.actual;
      s1.getRange(r, 1).setValue(g.label).setFontWeight('bold');
      s1.getRange(r, 2).setValue(g.budget).setNumberFormat('#,##0').setHorizontalAlignment('right');
      s1.getRange(r, 3).setValue(g.actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
      s1.getRange(r, 4).setValue(diff).setNumberFormat('#,##0').setHorizontalAlignment('right');
      if (diff > 0) s1.getRange(r, 4).setFontColor('#2d8a4e');
      else if (diff < 0) s1.getRange(r, 4).setFontColor('#c0392b');
      if (i % 2 === 0) s1.getRange(r, 1, 1, 4).setBackground('#f7f7f8');
    });

    // Total row
    const totalRow = hdrRow + 1 + groups.length;
    const totalBudget = groups.reduce((s, g) => s + g.budget, 0);
    const totalActual = groups.reduce((s, g) => s + g.actual, 0);
    const totalDiff = totalBudget - totalActual;
    s1.getRange(totalRow, 1, 1, 4).setValues([['TOTAL', totalBudget, totalActual, totalDiff]])
      .setFontWeight('bold').setBackground('#fbe1d1');
    s1.getRange(totalRow, 2, 1, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');

    // Summary section
    const sumStart = totalRow + 2;
    const savingsActual = groups.find(g => g.id === 'savings')?.actual || 0;
    const savingsRate = income > 0 ? Math.round((savingsActual / income) * 100) : 0;
    const summary = [
      ['RINGKASAN', ''],
      ['Penghasilan', income],
      ['Total Pengeluaran', totalActual],
      ['Sisa', totalDiff],
      ['Rate Tabungan', savingsRate + '%'],
    ];
    s1.getRange(sumStart, 1, summary.length, 2).setValues(summary);
    s1.getRange(sumStart, 1).setFontSize(12).setFontColor('#5d2a1a').setFontWeight('bold');
    for (let i = 1; i < summary.length; i++) {
      s1.getRange(sumStart + i, 1).setFontWeight('bold');
      s1.getRange(sumStart + i, 2).setNumberFormat(i < 4 ? '#,##0' : '@').setHorizontalAlignment('right');
    }

    s1.getRange('A:D').setBorder(true, true, true, true, false, false);
    s1.setColumnWidths(1, 4, 160);

    // === SHEET 2: DETAIL ===
    let s2 = ss.getSheetByName('Detail');
    if (s2) s2.clear(); else s2 = ss.insertSheet('Detail');
    s2.getRange('A1:D1').merge()
      .setValue('STEEP BUDGET — DETAIL PENGELUARAN')
      .setFontSize(16).setFontColor('#5d2a1a').setFontFamily('Georgia')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    s2.getRange('A2:D2').merge()
      .setValue(monthLabel)
      .setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center');

    let row = 4;
    groups.forEach(g => {
      s2.getRange(row, 1, 1, 4).setValues([[g.label, 'Anggaran', 'Realisasi', 'Selisih']])
        .setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#17191c').setHorizontalAlignment('center');
      row++;
      g.items.forEach((item, ii) => {
        const diff = item.budget - item.actual;
        s2.getRange(row, 1).setValue(item.name);
        s2.getRange(row, 2).setValue(item.budget).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s2.getRange(row, 3).setValue(item.actual).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s2.getRange(row, 4).setValue(diff).setNumberFormat('#,##0').setHorizontalAlignment('right');
        if (diff > 0) s2.getRange(row, 4).setFontColor('#2d8a4e');
        else if (diff < 0) s2.getRange(row, 4).setFontColor('#c0392b');
        if (ii % 2 === 0) s2.getRange(row, 1, 1, 4).setBackground('#f7f7f8');
        row++;
      });
      // subtotal
      const subBudget = g.items.reduce((s, it) => s + it.budget, 0);
      const subActual = g.items.reduce((s, it) => s + it.actual, 0);
      const subDiff = subBudget - subActual;
      s2.getRange(row, 1, 1, 4).setValues([['Subtotal ' + g.label, subBudget, subActual, subDiff]])
        .setFontWeight('bold').setBackground('#fbe1d1');
      s2.getRange(row, 2, 1, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
      row += 2;
    });

    s2.getRange('A:D').setBorder(true, true, true, true, false, false);
    s2.setColumnWidths(1, 4, 180);

    // === SHEET 3: RIWAYAT ===
    let s3 = ss.getSheetByName('Riwayat');
    if (s3) s3.clear(); else s3 = ss.insertSheet('Riwayat');
    s3.getRange('A1:G1').merge()
      .setValue('STEEP BUDGET — RIWAYAT BULANAN')
      .setFontSize(16).setFontColor('#5d2a1a').setFontFamily('Georgia')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    if (history.length > 0) {
      const headers = ['Bulan', 'Penghasilan', 'Kebutuhan', 'Keinginan', 'Tabungan', 'Total', 'Sisa'];
      s3.getRange('A3:G3').setValues([headers])
        .setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#17191c').setHorizontalAlignment('center');

      history.forEach((h, i) => {
        const r = 4 + i;
        s3.getRange(r, 1).setValue(h.month);
        s3.getRange(r, 2).setValue(h.income).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s3.getRange(r, 3).setValue(h.needs).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s3.getRange(r, 4).setValue(h.wants).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s3.getRange(r, 5).setValue(h.savings).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s3.getRange(r, 6).setValue(h.total).setNumberFormat('#,##0').setHorizontalAlignment('right');
        s3.getRange(r, 7).setValue(h.sisa).setNumberFormat('#,##0').setHorizontalAlignment('right');
        if (i % 2 === 0) s3.getRange(r, 1, 1, 7).setBackground('#f7f7f8');
      });

      s3.getRange('A:G').setBorder(true, true, true, true, false, false);
      s3.setColumnWidths(1, 7, 140);
    }

    // Move Ringkasan to first position
    ss.setActiveSheet(ss.getSheetByName('Ringkasan'));
    ss.moveActiveSheet(1);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, url: ss.getUrl() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Steep Budget Sync API siap.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
