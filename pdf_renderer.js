import PDFDocument from 'pdfkit';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Constants for layout
const COLORS = {
  primary: '#0e3a75', // Dark Railway Blue
  secondary: '#1c6abf', // Lighter Blue
  success: '#1e8e3e', // Green Status
  danger: '#d93025',  // Red Status
  warning: '#f29900', // Orange Status
  text: '#333333',
  lightGray: '#f5f5f5',
  border: '#dddddd',
  white: '#ffffff'
};

const LOGO_IR = path.join(__dirname, 'assets', 'indian_railway.png');
const LOGO_PARTNER = path.join(__dirname, 'assets', 'mirtha.jpg');

export class PDFRenderer {
  constructor(options = {}) {
    this.doc = new PDFDocument({ margin: 30, size: 'A4' });
    this.buffers = [];
    this.doc.on('data', chunk => this.buffers.push(chunk));
  }

  async getBuffer() {
    return new Promise((resolve, reject) => {
      this.doc.on('end', () => resolve(Buffer.concat(this.buffers)));
      this.doc.on('error', reject);
      this.doc.end();
    });
  }

  // --- Helper Methods ---
  drawHeader(title, subtitle, statusText, isApproved = true) {
    // Top banner
    this.doc.rect(30, 30, 535, 60).fill(COLORS.primary);
    
    // Logos
    if (fs.existsSync(LOGO_IR)) {
      this.doc.image(LOGO_IR, 40, 35, { height: 50 });
    }
    
    if (fs.existsSync(LOGO_PARTNER)) {
      this.doc.image(LOGO_PARTNER, 505, 35, { height: 50, width: 50 });
    }

    // Title text
    this.doc.fillColor(COLORS.white)
      .font('Helvetica-Bold')
      .fontSize(16)
      .text(title, 95, 45, { width: 285, align: 'center' });
    
    this.doc.fontSize(10)
      .font('Helvetica')
      .text(subtitle, 95, 70, { width: 285, align: 'center' });

    // Status Badge background
    this.doc.rect(390, 35, 105, 50).fill(COLORS.white);
    
    this.doc.fillColor(isApproved ? COLORS.success : COLORS.danger)
      .font('Helvetica-Bold')
      .fontSize(8)
      .text('FINAL COMPLIANCE STATUS', 395, 45, { width: 95, align: 'center' });

    this.doc.fontSize(12)
      .text(statusText, 395, 60, { width: 95, align: 'center' });
      
    this.doc.moveDown(2);
  }

  drawSectionHeader(title, yOffset) {
    if (this.doc.y + 50 > 800) this.doc.addPage();
    const currentY = yOffset || this.doc.y;
    
    this.doc.rect(30, currentY, 535, 20).fill(COLORS.primary);
    this.doc.fillColor(COLORS.white)
      .font('Helvetica-Bold')
      .fontSize(10)
      .text(title, 40, currentY + 5);
      
    return currentY + 20;
  }

  drawKeyValueGrid(data, startY, columns = 2) {
    const colWidth = 535 / columns;
    const rowHeight = 20;
    let y = startY;
    
    this.doc.fontSize(9);
    
    Object.entries(data).forEach(([key, value], index) => {
      if (this.doc.y + rowHeight > 800) {
        this.doc.addPage();
        y = this.doc.y;
      }
      
      const isRightCol = (index % columns) !== 0;
      const x = 30 + (isRightCol ? colWidth : 0);
      
      // Draw row borders
      if (!isRightCol) {
        this.doc.rect(30, y, 535, rowHeight).stroke(COLORS.border);
        // Alternate background
        if ((index / columns) % 2 === 0) {
           this.doc.rect(30, y, 535, rowHeight).fillAndStroke(COLORS.lightGray, COLORS.border);
        }
      }

      this.doc.fillColor(COLORS.text).font('Helvetica-Bold').text(key, x + 10, y + 5, { width: colWidth * 0.4 });
      this.doc.font('Helvetica').text(': ' + (value || 'N/A'), x + 10 + (colWidth * 0.4), y + 5, { width: colWidth * 0.6 - 10 });

      if (isRightCol) {
        y += rowHeight;
        this.doc.y = y;
      }
    });
    
    if (Object.keys(data).length % columns !== 0) {
       this.doc.y += rowHeight;
    }
    
    this.doc.moveDown(1);
  }

  drawTable(headers, rows, colWidths) {
    if (this.doc.y + 60 > 800) this.doc.addPage();
    
    let y = this.doc.y;
    
    // Header
    this.doc.rect(30, y, 535, 20).fill(COLORS.lightGray);
    this.doc.fillColor(COLORS.primary).font('Helvetica-Bold').fontSize(8);
    
    let x = 35;
    headers.forEach((h, i) => {
      this.doc.text(h, x, y + 5, { width: colWidths[i] - 5 });
      x += colWidths[i];
    });
    
    y += 20;
    
    // Rows
    this.doc.font('Helvetica').fontSize(8).fillColor(COLORS.text);
    
    rows.forEach((row, rowIndex) => {
      const rowHeight = 25; // can be dynamic if text wraps
      if (y + rowHeight > 800) {
        this.doc.addPage();
        y = this.doc.y;
      }
      
      this.doc.rect(30, y, 535, rowHeight).stroke(COLORS.border);
      
      let curX = 35;
      row.forEach((cell, i) => {
        // Simple coloring for status
        if (cell === 'Verified' || cell === 'Completed' || cell === 'Pass') {
          this.doc.fillColor(COLORS.success).font('Helvetica-Bold');
        } else if (cell === 'Missing' || cell === 'Failed' || cell === 'Fail' || cell === 'Late') {
          this.doc.fillColor(COLORS.danger).font('Helvetica-Bold');
        } else {
          this.doc.fillColor(COLORS.text).font('Helvetica');
        }
        
        this.doc.text(String(cell), curX, y + 8, { width: colWidths[i] - 5 });
        curX += colWidths[i];
      });
      
      y += rowHeight;
    });
    
    this.doc.y = y + 10;
  }
  
  drawFooter(reportId) {
    const pages = this.doc.bufferedPageRange();
    for (let i = 0; i < pages.count; i++) {
      this.doc.switchToPage(i);
      const y = 810;
      this.doc.rect(30, y, 535, 20).fill(COLORS.primary);
      this.doc.fillColor(COLORS.white).fontSize(7).font('Helvetica');
      this.doc.text(`Report ID: ${reportId}`, 40, y + 5);
      this.doc.text(`Page ${i + 1} of ${pages.count}`, 0, y + 5, { align: 'center' });
      this.doc.text(`Generated On: ${new Date().toLocaleString('en-IN')}`, 0, y + 5, { align: 'right', margins: { right: 40 } });
    }
  }

  // --- Specialized Report Generators ---
  
  async generateAttendanceAudit(data) {
    const { meta, trainInfo, employeeInfo, attendanceList, kpi } = data;
    
    this.drawHeader(
      'OBHS ATTENDANCE & EVIDENCE COMPLIANCE AUDIT REPORT',
      'Attendance Verification | GPS Validation | Evidence Compliance | Operational Audit',
      kpi.overallStatus,
      kpi.isApproved
    );
    
    let y = this.drawSectionHeader('1. DOCUMENT CONTROL INFORMATION');
    this.drawKeyValueGrid({
      'Report ID': meta.reportId,
      'Generated On': meta.generatedOn,
      'Generated By': meta.generatedBy,
      'Audit Type': meta.auditType,
      'Report Classification': meta.classification,
      'Division': meta.division,
      'Audit Status': meta.auditStatus
    }, y);
    
    y = this.doc.y;
    // We can draw train and employee info side-by-side or stacked. Stacking is easier.
    y = this.drawSectionHeader('2. TRAIN & OPERATIONAL RUN INFORMATION', y);
    this.drawKeyValueGrid(trainInfo, y, 2);
    
    y = this.drawSectionHeader('3. EMPLOYEE & DEPLOYMENT INFORMATION', this.doc.y);
    this.drawKeyValueGrid(employeeInfo, y, 2);
    
    y = this.drawSectionHeader('4. ATTENDANCE COMPLIANCE VERIFICATION', this.doc.y);
    
    // Table
    const headers = ['Attendance Type', 'Time', 'Device TS', 'Server TS', 'GPS', 'Sync', 'Compliance'];
    const colWidths = [75, 60, 80, 80, 90, 60, 90];
    
    const rows = attendanceList.map(a => [
      a.type, a.time, a.deviceTs, a.serverTs, a.gps, a.syncStatus, a.complianceResult
    ]);
    
    this.doc.y = y;
    this.drawTable(headers, rows, colWidths);
    
    y = this.drawSectionHeader('5. KPI & OPERATIONAL PERFORMANCE SUMMARY', this.doc.y);
    const kpiHeaders = ['KPI Metric', 'Result', 'Status'];
    const kpiColWidths = [300, 100, 135];
    const kpiRows = kpi.metrics.map(k => [k.metric, k.value, k.status]);
    
    this.doc.y = y;
    this.drawTable(kpiHeaders, kpiRows, kpiColWidths);
    
    // Final Observation
    y = this.drawSectionHeader('6. FINAL AUDIT OBSERVATION', this.doc.y);
    this.doc.rect(30, y, 535, 40).stroke(COLORS.border);
    this.doc.fillColor(COLORS.text).font('Helvetica').fontSize(9)
      .text(kpi.observation, 40, y + 10, { width: 515 });
    
    this.doc.y = y + 60;
    
    this.drawFooter(meta.reportId);
    return this.getBuffer();
  }

  async generateWorkerActivityAudit(data) {
    const { meta, workerInfo, trainInfo, tasksList, kpi } = data;
    
    this.drawHeader(
      'OBHS WORKER ACTIVITY & EVIDENCE AUDIT REPORT',
      'Operational Audit | Task Execution | Evidence Verification | Compliance Review',
      kpi.overallStatus,
      kpi.isApproved
    );
    
    let y = this.drawSectionHeader('1. WORKER INFORMATION');
    this.drawKeyValueGrid(workerInfo, y, 2);
    
    y = this.drawSectionHeader('2. TRAIN & RUN INFORMATION', this.doc.y);
    this.drawKeyValueGrid(trainInfo, y, 2);
    
    y = this.drawSectionHeader('3. TASK EXECUTION & EVIDENCE DETAILS', this.doc.y);
    
    const headers = ['Task ID', 'Task Category', 'Coach', 'Completion Time', 'Worker Comment', 'Task Status'];
    const colWidths = [80, 100, 60, 100, 115, 80];
    
    const rows = tasksList.map(t => [
      t.id, t.category, t.coach, t.completionTime, t.comment, t.status
    ]);
    
    this.doc.y = y;
    this.drawTable(headers, rows, colWidths);
    
    y = this.drawSectionHeader('4. COMPLIANCE KPI SUMMARY', this.doc.y);
    const kpiHeaders = ['Metric', 'Value'];
    const kpiColWidths = [300, 235];
    const kpiRows = kpi.metrics.map(k => [k.metric, k.value]);
    
    this.doc.y = y;
    this.drawTable(kpiHeaders, kpiRows, kpiColWidths);
    
    this.drawFooter(meta.reportId);
    return this.getBuffer();
  }

  async generateComplaintAudit(data) {
    const { meta, trainInfo, complaintInfo, resolutionInfo, kpi } = data;
    
    this.drawHeader(
      'OBHS WORKER COMPLAINT & ISSUE TRACKING REPORT',
      'Worker Complaint Registration | Issue Tracking | Resolution Monitoring',
      kpi.overallStatus,
      kpi.overallStatus === 'RESOLVED' || kpi.overallStatus === 'CLOSED'
    );
    
    let y = this.drawSectionHeader('1. TRAIN & RUN INFORMATION');
    this.drawKeyValueGrid(trainInfo, y, 2);
    
    y = this.drawSectionHeader('2. WORKER COMPLAINT INFORMATION', this.doc.y);
    this.drawKeyValueGrid(complaintInfo, y, 2);
    
    y = this.drawSectionHeader('3. RESOLUTION TRACKING', this.doc.y);
    this.drawKeyValueGrid(resolutionInfo, y, 2);
    
    y = this.drawSectionHeader('4. COMPLAINT KPI SUMMARY', this.doc.y);
    const kpiHeaders = ['KPI Metric', 'Result'];
    const kpiColWidths = [300, 235];
    const kpiRows = kpi.metrics.map(k => [k.metric, k.value]);
    
    this.doc.y = y;
    this.drawTable(kpiHeaders, kpiRows, kpiColWidths);
    
    this.drawFooter(meta.reportId);
    return this.getBuffer();
  }

  async generateEnterpriseOperationalAudit(data) {
    const { meta, trainInfo, coachAssignment, timeline, kpi } = data;
    
    this.drawHeader(
      'OBHS ENTERPRISE TRAIN RUN & OPERATIONAL AUDIT REPORT',
      'Operational Audit | Journey Monitoring | Coach Management | Compliance Verified',
      kpi.overallStatus,
      kpi.isApproved
    );
    
    let y = this.drawSectionHeader('1. TRAIN & JOURNEY INFORMATION');
    this.drawKeyValueGrid(trainInfo, y, 2);
    
    y = this.drawSectionHeader('2. COACH & WORKER ASSIGNMENT DETAILS', this.doc.y);
    const headers = ['Coach', 'Type', 'Workers Assigned', 'Worker Names', 'Status'];
    const colWidths = [60, 80, 100, 215, 80];
    
    const rows = coachAssignment.map(c => [
      c.coach, c.type, c.workerCount, c.workerNames, c.status
    ]);
    
    this.doc.y = y;
    this.drawTable(headers, rows, colWidths);
    
    y = this.drawSectionHeader('3. JOURNEY TIMELINE & OPERATIONAL EVENTS', this.doc.y);
    const tHeaders = ['Journey Event', 'Time', 'Location', 'Operational Status'];
    const tColWidths = [150, 100, 150, 135];
    const tRows = timeline.map(t => [t.event, t.time, t.location, t.status]);
    
    this.doc.y = y;
    this.drawTable(tHeaders, tRows, tColWidths);
    
    y = this.drawSectionHeader('4. RUN INSTANCE & OPERATIONAL KPI SUMMARY', this.doc.y);
    const kpiHeaders = ['Operational Metric', 'Result'];
    const kpiColWidths = [300, 235];
    const kpiRows = kpi.metrics.map(k => [k.metric, k.value]);
    
    this.doc.y = y;
    this.drawTable(kpiHeaders, kpiRows, kpiColWidths);
    
    this.drawFooter(meta.reportId);
    return this.getBuffer();
  }
}
