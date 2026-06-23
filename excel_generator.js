import ExcelJS from 'exceljs';

export class ExcelGenerator {
  constructor() {
    this.workbook = new ExcelJS.Workbook();
    this.workbook.creator = 'OBHS Enterprise Monitoring System';
    this.workbook.lastModifiedBy = 'System';
    this.workbook.created = new Date();
    this.workbook.modified = new Date();
  }

  async getBuffer() {
    return await this.workbook.xlsx.writeBuffer();
  }

  createSummarySheet(data, title) {
    const sheet = this.workbook.addWorksheet('Summary', {
      views: [{ showGridLines: false }]
    });

    // Title
    sheet.mergeCells('A1:F2');
    const titleCell = sheet.getCell('A1');
    titleCell.value = title.toUpperCase();
    titleCell.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0E3A75' } // Dark Railway Blue
    };

    let y = 4;

    const addSection = (sectionTitle, keyValueData) => {
      sheet.mergeCells(`A${y}:F${y}`);
      const secHeader = sheet.getCell(`A${y}`);
      secHeader.value = sectionTitle;
      secHeader.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      secHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1C6ABF' } };
      y++;

      Object.entries(keyValueData).forEach(([key, value]) => {
        sheet.getCell(`B${y}`).value = key;
        sheet.getCell(`B${y}`).font = { bold: true };
        sheet.getCell(`C${y}`).value = value;
        y++;
      });
      y++; // Spacing
    };

    if (data.meta) addSection('DOCUMENT CONTROL INFORMATION', data.meta);
    if (data.trainInfo) addSection('TRAIN INFORMATION', data.trainInfo);
    if (data.kpi && data.kpi.metrics) {
      sheet.mergeCells(`A${y}:F${y}`);
      const secHeader = sheet.getCell(`A${y}`);
      secHeader.value = 'KPI SUMMARY';
      secHeader.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      secHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1C6ABF' } };
      y++;

      sheet.getCell(`B${y}`).value = 'Metric';
      sheet.getCell(`C${y}`).value = 'Result';
      sheet.getCell(`D${y}`).value = 'Status';
      sheet.getRow(y).font = { bold: true };
      sheet.getRow(y).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
      y++;

      data.kpi.metrics.forEach(k => {
        sheet.getCell(`B${y}`).value = k.metric;
        sheet.getCell(`C${y}`).value = k.value;
        const statusCell = sheet.getCell(`D${y}`);
        statusCell.value = k.status;
        if (['Pass', 'Verified', 'Completed'].includes(k.status)) {
           statusCell.font = { color: { argb: 'FF1E8E3E' }, bold: true };
        } else if (['Fail', 'Missing', 'Late'].includes(k.status)) {
           statusCell.font = { color: { argb: 'FFD93025' }, bold: true };
        }
        y++;
      });
    }

    sheet.getColumn('A').width = 5;
    sheet.getColumn('B').width = 30;
    sheet.getColumn('C').width = 30;
    sheet.getColumn('D').width = 20;
    sheet.getColumn('E').width = 10;
    sheet.getColumn('F').width = 10;
  }

  createDataTableSheet(sheetName, headers, rows) {
    const sheet = this.workbook.addWorksheet(sheetName);
    
    // Setup columns
    sheet.columns = headers.map(h => ({ header: h, key: h, width: 20 }));
    
    // Style header row
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0E3A75' } };
    
    // Add rows
    rows.forEach(r => sheet.addRow(r));
  }
}
