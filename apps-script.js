// ============================================================
// NBM Dashboard — Google Apps Script
// ============================================================
// INSTRUCCIONES:
// 1. Crea un Google Sheet nuevo (o usa uno existente)
// 2. Ve a Extensiones → Apps Script
// 3. Borra el código que viene por defecto y pega TODO este archivo
// 4. Click en "Implementar" → "Nueva implementación"
//    - Tipo: Aplicación web
//    - Ejecutar como: Yo
//    - Quién tiene acceso: Cualquier persona
// 5. Click en "Implementar" → copia la URL que te da
// 6. Pega esa URL en el Dashboard → Configuración → URL del Web App
// ============================================================

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const meses = payload.meses || [];

    // === Hoja _raw: backup completo del JSON ===
    let rawSheet = ss.getSheetByName('_raw');
    if (!rawSheet) { rawSheet = ss.insertSheet('_raw'); rawSheet.hideSheet(); }
    rawSheet.getRange('A1').setValue(JSON.stringify(meses));
    rawSheet.getRange('A2').setValue(new Date().toISOString());

    // === Hoja Resumen: vista mensual formateada ===
    let res = ss.getSheetByName('Resumen');
    if (!res) res = ss.insertSheet('Resumen');
    res.clear();

    const h = ['Mes','Año','Comisiones Netas','Gastos Deducibles','Sueldo','Retiro','IVA%','PPM%','IVA Débito','IVA Crédito','IVA Neto','PPM','Total F29','Caja Neta','Utilidad'];
    res.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#f2f4f6');

    if (meses.length) {
      const rows = meses.map(m => {
        const iva = (m.tasaIva || 19) / 100;
        const ppm = (m.tasaPpm || 0.2) / 100;
        const gN = (m.gastos || []).reduce((a, g) => a + g.monto, 0);
        const ivaD = m.comisiones * iva;
        const ivaC = gN * iva;
        const ivaP = Math.max(0, ivaD - ivaC);
        const ppmM = m.comisiones * ppm;
        const f29 = ivaP + ppmM;
        const caja = m.comisiones - gN - (m.sueldo||0) - (m.retiro||0) - ppmM;
        const util = m.comisiones - gN - (m.sueldo||0);
        return [m.mes, m.anio, m.comisiones, gN, m.sueldo||0, m.retiro||0,
                m.tasaIva||19, m.tasaPpm||0.2,
                Math.round(ivaD), Math.round(ivaC), Math.round(ivaP),
                Math.round(ppmM), Math.round(f29), Math.round(caja), Math.round(util)];
      });
      res.getRange(2, 1, rows.length, h.length).setValues(rows);

      // Formato moneda
      const moneyFmt = '$#,##0';
      res.getRange(2, 3, rows.length, 4).setNumberFormat(moneyFmt);
      res.getRange(2, 9, rows.length, 7).setNumberFormat(moneyFmt);
      res.autoResizeColumns(1, h.length);

      // Fila de totales
      const tr = rows.length + 2;
      res.getRange(tr, 1).setValue('TOTAL').setFontWeight('bold');
      [3,4,5,6,9,10,11,12,13,14,15].forEach(c => {
        const colLetter = String.fromCharCode(64 + c);
        res.getRange(tr, c).setFormula('=SUM(' + colLetter + '2:' + colLetter + (rows.length+1) + ')');
      });
      res.getRange(tr, 1, 1, h.length).setFontWeight('bold').setBackground('#f2f4f6').setNumberFormat(moneyFmt);
    }

    // === Hoja Gastos: detalle de cada gasto ===
    let gas = ss.getSheetByName('Gastos');
    if (!gas) gas = ss.insertSheet('Gastos');
    gas.clear();
    gas.getRange(1, 1, 1, 4).setValues([['Mes','Año','Descripción','Monto Neto']]).setFontWeight('bold').setBackground('#f2f4f6');

    const gRows = [];
    meses.forEach(m => {
      (m.gastos || []).forEach(g => {
        gRows.push([m.mes, m.anio, g.desc, g.monto]);
      });
    });
    if (gRows.length) {
      gas.getRange(2, 1, gRows.length, 4).setValues(gRows);
      gas.getRange(2, 4, gRows.length, 1).setNumberFormat('$#,##0');
      gas.autoResizeColumns(1, 4);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      meses: meses.length,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('_raw');
    if (!rawSheet) {
      return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    }
    const raw = rawSheet.getRange('A1').getValue();
    return ContentService.createTextOutput(raw || '[]').setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
}
