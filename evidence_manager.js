import sharp from 'sharp';
import crypto from 'crypto';
import path from 'path';
import fs from 'fs';
import admin from 'firebase-admin';

const EVIDENCE_TYPES = [
  'Before', 'After', 'Exception', 'Complaint', 'Resolution',
  'Attendance', 'FaceVerification', 'GPSVerification',
  'SupervisorVerification', 'LinenEvidence', 'PassengerComplaintEvidence'
];

const STORAGE_TIER = {
  ACTIVE: 'active',
  ARCHIVE: 'archive',
  LONG_TERM: 'long_term'
};

// File path builder: Division/Train/Date/Coach/Task/EvidenceType/
export function buildStoragePath(division, trainNumber, date, coach, taskId, evidenceType, filename) {
  const sanitize = s => String(s).replace(/[^a-zA-Z0-9_-]/g, '_');
  return [
    sanitize(division),
    sanitize(trainNumber),
    date.replace(/[/\\]/g, '-'),
    sanitize(coach),
    sanitize(taskId),
    sanitize(evidenceType),
    filename
  ].join('/');
}

// Compute MD5 hash for duplicate detection
export function computeFileHash(buffer) {
  return crypto.createHash('md5').update(buffer).digest('hex');
}

// Check if file hash already exists in evidence_metadata
export async function findDuplicateByHash(db, hash) {
  const snap = await db.collection('evidence_metadata')
    .where('fileHash', '==', hash)
    .where('deleted', '==', false)
    .limit(1)
    .get();
  return snap.empty ? null : { id: snap.docs[0].id, ...snap.docs[0].data() };
}

// Image processing pipeline: original → WEBP + thumbnail + compressed
export async function processImage(buffer, options = {}) {
  const {
    thumbnailWidth = 200,
    thumbnailHeight = 200,
    maxWidth = 1920,
    maxHeight = 1080,
    quality = 80,
    thumbnailQuality = 60
  } = options;

  const image = sharp(buffer);
  const metadata = await image.metadata();

  // Generate WEBP compressed version
  const webpBuffer = await image
    .resize({
      width: Math.min(metadata.width || maxWidth, maxWidth),
      height: Math.min(metadata.height || maxHeight, maxHeight),
      fit: 'inside',
      withoutEnlargement: true
    })
    .webp({ quality })
    .toBuffer();

  // Generate thumbnail
  const thumbnailBuffer = await image
    .resize(thumbnailWidth, thumbnailHeight, { fit: 'cover', position: 'centre' })
    .webp({ quality: thumbnailQuality })
    .toBuffer();

  return {
    originalBuffer: buffer,
    webpBuffer,
    thumbnailBuffer,
    originalSize: buffer.length,
    compressedSize: webpBuffer.length,
    thumbnailSize: thumbnailBuffer.length,
    originalFormat: metadata.format,
    width: metadata.width,
    height: metadata.height,
    compressionRatio: ((1 - webpBuffer.length / buffer.length) * 100).toFixed(1)
  };
}

// Upload processed images to Firebase Storage with hierarchy
export async function uploadToStorage(bucket, storagePath, processed, contentType = 'image/webp') {
  const uniqueToken = crypto.randomUUID();

  // Upload WEBP
  const webpPath = `evidence/${storagePath}.webp`;
  const webpFile = bucket.file(webpPath);
  await webpFile.save(processed.webpBuffer, {
    metadata: {
      contentType: 'image/webp',
      metadata: { firebaseStorageDownloadTokens: uniqueToken }
    }
  });

  // Upload thumbnail
  const thumbPath = `evidence/thumbnails/${storagePath}_thumb.webp`;
  const thumbFile = bucket.file(thumbPath);
  await thumbFile.save(processed.thumbnailBuffer, {
    metadata: {
      contentType: 'image/webp',
      metadata: { firebaseStorageDownloadTokens: crypto.randomUUID() }
    }
  });

  const bucketName = bucket.name;
  const baseUrl = `https://firebasestorage.googleapis.com/v0/b/${bucketName}/o/`;

  return {
    webpUrl: `${baseUrl}${encodeURIComponent(webpPath)}?alt=media&token=${uniqueToken}`,
    thumbUrl: `${baseUrl}${encodeURIComponent(thumbPath)}?alt=media`,
    webpPath,
    thumbPath,
    webpSize: processed.compressedSize,
    thumbSize: processed.thumbnailSize
  };
}

// Create evidence metadata document
export function buildEvidenceMetadata(params) {
  const {
    storageRef, trainNumber, coach, taskId, taskType,
    evidenceType, uploadedBy, uploadedByName, uploadedByRole,
    division, contractor, runInstanceId, fileHash,
    originalSize, compressedSize, thumbnailSize,
    compressionRatio, width, height, originalFormat,
    webpUrl, thumbUrl, webpPath, thumbPath,
    gpsLat, gpsLng, remarks, complaintId, attendanceId,
    thumbnailWidth = 200, thumbnailHeight = 200
  } = params;

  return {
    fileHash,
    storageRef,
    trainNumber: trainNumber || null,
    coach: coach || null,
    taskId: taskId || null,
    taskType: taskType || null,
    evidenceType: evidenceType || 'Before',
    uploadedBy,
    uploadedByName: uploadedByName || 'Unknown',
    uploadedByRole: uploadedByRole || null,
    division: division || null,
    contractor: contractor || null,
    runInstanceId: runInstanceId || null,
    complaintId: complaintId || null,
    attendanceId: attendanceId || null,
    gpsLat: gpsLat || null,
    gpsLng: gpsLng || null,
    remarks: remarks || null,
    // File info
    webpUrl,
    thumbUrl,
    webpPath,
    thumbPath,
    originalFormat: originalFormat || 'unknown',
    originalSize,
    compressedSize,
    thumbnailSize,
    compressionRatio: parseFloat(compressionRatio) || 0,
    width: width || null,
    height: height || null,
    thumbnailWidth,
    thumbnailHeight,
    // Storage tier
    storageTier: STORAGE_TIER.ACTIVE,
    // Timestamps
    uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
    archivedAt: null,
    longTermAt: null,
    lastAccessedAt: admin.firestore.FieldValue.serverTimestamp(),
    // Flags
    deleted: false,
    deletedAt: null,
    viewCount: 0,
    downloadCount: 0,
    // Metadata
    contentType: 'image/webp'
  };
}

// Log audit event
export async function logAudit(db, action, params) {
  const {
    evidenceId, userId, userName, userRole, trainNumber, coach,
    taskId, complaintId, details, ipAddress
  } = params;

  const auditEntry = {
    action,
    evidenceId: evidenceId || null,
    userId: userId || null,
    userName: userName || 'System',
    userRole: userRole || null,
    trainNumber: trainNumber || null,
    coach: coach || null,
    taskId: taskId || null,
    complaintId: complaintId || null,
    details: details || null,
    ipAddress: ipAddress || null,
    timestamp: admin.firestore.FieldValue.serverTimestamp()
  };

  await db.collection('audit_evidence').add(auditEntry);
  return auditEntry;
}

// Archive records older than given days
export async function archiveOldRecords(db, bucket, daysOld = 7, batchSize = 100) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysOld);

  let processed = 0;
  let lastDoc = null;

  while (true) {
    let query = db.collection('evidence_metadata')
      .where('storageTier', '==', STORAGE_TIER.ACTIVE)
      .where('uploadedAt', '<', cutoff)
      .where('deleted', '==', false)
      .orderBy('uploadedAt')
      .limit(batchSize);

    if (lastDoc) {
      query = query.startAfter(lastDoc);
    }

    const snap = await query.get();
    if (snap.empty) break;

    const batch = db.batch();
    for (const doc of snap.docs) {
      const data = { ...doc.data(), archivedAt: admin.firestore.FieldValue.serverTimestamp() };
      const archiveRef = db.collection('archive_evidence').doc(doc.id);
      batch.set(archiveRef, { ...data, storageTier: STORAGE_TIER.ARCHIVE });
      batch.update(doc.ref, {
        storageTier: STORAGE_TIER.ARCHIVE,
        archivedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      await logAudit(db, 'EVIDENCE_ARCHIVED', {
        evidenceId: doc.id,
        trainNumber: data.trainNumber,
        coach: data.coach,
        taskId: data.taskId,
        details: `Auto-archived after ${daysOld} days`
      });
    }
    await batch.commit();
    processed += snap.docs.length;
    lastDoc = snap.docs[snap.docs.length - 1];
  }

  return { archived: processed, daysOld };
}

// Move to long-term storage after 30 days
export async function moveToLongTerm(db, daysOld = 30, batchSize = 100) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysOld);

  let processed = 0;
  let lastDoc = null;

  while (true) {
    let query = db.collection('archive_evidence')
      .where('storageTier', '==', STORAGE_TIER.ARCHIVE)
      .where('archivedAt', '<', cutoff)
      .where('deleted', '==', false)
      .orderBy('archivedAt')
      .limit(batchSize);

    if (lastDoc) {
      query = query.startAfter(lastDoc);
    }

    const snap = await query.get();
    if (snap.empty) break;

    const batch = db.batch();
    for (const doc of snap.docs) {
      const data = { ...doc.data(), longTermAt: admin.firestore.FieldValue.serverTimestamp() };
      const longRef = db.collection('long_term_evidence').doc(doc.id);
      batch.set(longRef, { ...data, storageTier: STORAGE_TIER.LONG_TERM });
      batch.update(doc.ref, {
        storageTier: STORAGE_TIER.LONG_TERM,
        longTermAt: admin.firestore.FieldValue.serverTimestamp()
      });

      await logAudit(db, 'EVIDENCE_LONG_TERM', {
        evidenceId: doc.id,
        trainNumber: data.trainNumber,
        details: `Moved to long-term storage after ${daysOld} days`
      });
    }
    await batch.commit();
    processed += snap.docs.length;
    lastDoc = snap.docs[snap.docs.length - 1];
  }

  return { moved: processed, daysOld };
}

// Restore from archive to active
export async function restoreFromArchive(db, evidenceId) {
  const archiveDoc = await db.collection('archive_evidence').doc(evidenceId).get();
  if (!archiveDoc.exists) {
    const longDoc = await db.collection('long_term_evidence').doc(evidenceId).get();
    if (!longDoc.exists) throw new Error('Evidence not found in archive or long-term storage');
    const data = longDoc.data();
    await db.collection('evidence_metadata').doc(evidenceId).set({
      ...data,
      storageTier: STORAGE_TIER.ACTIVE,
      archivedAt: null,
      longTermAt: null,
      lastAccessedAt: admin.firestore.FieldValue.serverTimestamp()
    });
    await logAudit(db, 'EVIDENCE_RESTORED', {
      evidenceId, details: 'Restored from long-term to active'
    });
    return { restored: true, tier: 'long_term' };
  }

  const data = archiveDoc.data();
  await db.collection('evidence_metadata').doc(evidenceId).set({
    ...data,
    storageTier: STORAGE_TIER.ACTIVE,
    archivedAt: null,
    lastAccessedAt: admin.firestore.FieldValue.serverTimestamp()
  });
  await logAudit(db, 'EVIDENCE_RESTORED', {
    evidenceId, details: 'Restored from archive to active'
  });
  return { restored: true, tier: 'archive' };
}

// Get evidence by ID (from active, archive, or long-term)
export async function getEvidenceById(db, evidenceId) {
  const collections = ['evidence_metadata', 'archive_evidence', 'long_term_evidence'];
  for (const coll of collections) {
    const doc = await db.collection(coll).doc(evidenceId).get();
    if (doc.exists) {
      const data = doc.data();
      // Increment view count
      doc.ref.update({ viewCount: admin.firestore.FieldValue.increment(1), lastAccessedAt: admin.firestore.FieldValue.serverTimestamp() }).catch(() => {});
      return { id: doc.id, ...data };
    }
  }
  return null;
}

// Search evidence across active and archived storage
export async function searchEvidence(db, params) {
  const {
    trainNumber, coach, dateFrom, dateTo, contractor, worker,
    supervisor, taskType, evidenceType, complaintId, runInstanceId,
    storageTier, uploadedBy, limit = 50, offset = 0
  } = params;

  const collections = storageTier
    ? [storageTier === STORAGE_TIER.ACTIVE ? 'evidence_metadata' :
       storageTier === STORAGE_TIER.ARCHIVE ? 'archive_evidence' : 'long_term_evidence']
    : ['evidence_metadata', 'archive_evidence', 'long_term_evidence'];

  const results = [];
  let totalCount = 0;

  for (const coll of collections) {
    let query = db.collection(coll);
    const constraints = [];

    if (trainNumber) constraints.push({ field: 'trainNumber', op: '==', val: trainNumber });
    if (coach) constraints.push({ field: 'coach', op: '==', val: coach });
    if (contractor) constraints.push({ field: 'contractor', op: '==', val: contractor });
    if (uploadedBy) constraints.push({ field: 'uploadedBy', op: '==', val: uploadedBy });
    if (evidenceType) constraints.push({ field: 'evidenceType', op: '==', val: evidenceType });
    if (taskType) constraints.push({ field: 'taskType', op: '==', val: taskType });
    if (complaintId) constraints.push({ field: 'complaintId', op: '==', val: complaintId });
    if (runInstanceId) constraints.push({ field: 'runInstanceId', op: '==', val: runInstanceId });
    constraints.push({ field: 'deleted', op: '==', val: false });

    // Firestore requires composite indexes for multiple equality filters;
    // we apply constraints in memory for flexible search
    let snap;
    try {
      // Use first constraint for indexed query if available
      if (constraints.length > 0) {
        const first = constraints[0];
        query = query.where(first.field, first.op, first.val);
      }
      snap = await query.orderBy('uploadedAt', 'desc').get();
    } catch {
      snap = await query.get();
    }

    let docs = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    // Apply remaining constraints in memory
    for (let i = 1; i < constraints.length; i++) {
      const c = constraints[i];
      docs = docs.filter(d => d[c.field] === c.val);
    }

    if (worker) docs = docs.filter(d => d.uploadedByName?.toLowerCase().includes(worker.toLowerCase()));
    if (supervisor) docs = docs.filter(d => d.uploadedByName?.toLowerCase().includes(supervisor.toLowerCase()) || d.uploadedByRole?.toLowerCase() === 'supervisor');
    if (dateFrom) docs = docs.filter(d => d.uploadedAt?.toDate() >= new Date(dateFrom));
    if (dateTo) docs = docs.filter(d => d.uploadedAt?.toDate() <= new Date(dateTo + 'T23:59:59Z'));

    totalCount += docs.length;
    results.push(...docs);
  }

  // Sort by upload time descending
  results.sort((a, b) => {
    const tA = a.uploadedAt?.toDate?.() || new Date(0);
    const tB = b.uploadedAt?.toDate?.() || new Date(0);
    return tB - tA;
  });

  return {
    total: results.length,
    page: Math.floor(offset / limit) + 1,
    limit,
    results: results.slice(offset, offset + limit)
  };
}

// === PDF REPORT GENERATION ===

export async function generatePDFReport(reportData) {
  const PDFDocument = (await import('pdfkit')).default;
  const doc = new PDFDocument({ margin: 50, size: 'A4' });
  const buffers = [];

  return new Promise((resolve, reject) => {
    doc.on('data', chunk => buffers.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.on('error', reject);

    const {
      title, generatedBy, trainNumber, coach, contractor,
      periodStart, periodEnd, sections = [], images = []
    } = reportData;

    // Header
    doc.fontSize(18).font('Helvetica-Bold').text('INDIAN RAILWAYS', { align: 'center' });
    doc.fontSize(14).text('OBHS - Coach & Washroom Cleaning System', { align: 'center' });
    doc.moveDown(0.5);
    doc.fontSize(12).font('Helvetica-Bold').text(title || 'EVIDENCE REPORT', { align: 'center' });
    doc.moveDown();

    // Meta info
    doc.fontSize(9).font('Helvetica');
    const metaLines = [
      `Generated By: ${generatedBy || 'System'}`,
      `Period: ${periodStart || 'N/A'} to ${periodEnd || 'N/A'}`
    ];
    if (trainNumber) metaLines.push(`Train: ${trainNumber}`);
    if (coach) metaLines.push(`Coach: ${coach}`);
    if (contractor) metaLines.push(`Contractor: ${contractor}`);

    metaLines.forEach(line => {
      doc.text(line, { align: 'left' });
    });
    doc.moveDown();

    // Divider
    doc.moveTo(50, doc.y).lineTo(545, doc.y).stroke('#cccccc');
    doc.moveDown();

    // Sections
    for (const section of sections) {
      if (doc.y > 700) doc.addPage();

      doc.fontSize(11).font('Helvetica-Bold').text(section.heading, { underline: true });
      doc.moveDown(0.3);

      if (section.items && section.items.length > 0) {
        // Table header
        const colWidths = section.colWidths || [80, 120, 120, 80, 80];
        const headers = section.headers || ['Field', 'Value'];
        const startX = 50;
        let y = doc.y;

        doc.fontSize(8).font('Helvetica-Bold');
        let x = startX;
        headers.forEach((h, i) => {
          doc.text(h, x, y, { width: colWidths[i], align: 'left' });
          x += colWidths[i];
        });
        doc.moveDown(0.5);
        y = doc.y;
        doc.moveTo(50, y - 2).lineTo(545, y - 2).stroke('#dddddd');

        // Table rows
        doc.fontSize(8).font('Helvetica');
        for (const item of section.items) {
          if (doc.y > 720) doc.addPage();
          y = doc.y;
          x = startX;
          const values = section.fields ? section.fields.map(f => item[f] || '') : [item.label || '', item.value || ''];
          values.forEach((v, i) => {
            doc.text(String(v), x, y, { width: colWidths[i], align: 'left' });
            x += colWidths[i];
          });
          doc.moveDown(0.4);
        }
      } else {
        doc.fontSize(9).font('Helvetica').text('No data available', { indent: 20 });
      }
      doc.moveDown();
    }

    // Images section
    if (images.length > 0) {
      if (doc.y > 650) doc.addPage();
      doc.fontSize(11).font('Helvetica-Bold').text('EVIDENCE IMAGES', { underline: true });
      doc.moveDown();

      for (const img of images) {
        if (doc.y > 700) doc.addPage();
        try {
          if (img.path && fs.existsSync(img.path)) {
            const imgWidth = Math.min(250, img.width || 200);
            doc.image(img.path, 50, doc.y, { width: imgWidth });
            doc.moveDown(4);
            doc.fontSize(8).font('Helvetica').text(`${img.caption || ''}`, { indent: 50 });
            doc.moveDown(0.5);
          } else if (img.url) {
            doc.fontSize(8).font('Helvetica').text(`[Image: ${img.caption || img.url}]`, { indent: 20 });
            doc.moveDown(0.3);
          }
        } catch {
          doc.fontSize(8).font('Helvetica').text(`[Image unavailable: ${img.caption || 'unknown'}]`, { indent: 20 });
          doc.moveDown(0.3);
        }
      }
    }

    // Footer
    const footerY = doc.page.height - 50;
    doc.fontSize(7).font('Helvetica').text(
      `Generated on ${new Date().toLocaleString()} | Railway OBHS Evidence Management System`,
      50, footerY, { align: 'center', width: 495 }
    );

    doc.end();
  });
}

// === STORAGE ANALYTICS ===

export async function getStorageAnalytics(db) {
  const collections = ['evidence_metadata', 'archive_evidence', 'long_term_evidence'];
  const tierNames = ['Active', 'Archive', 'Long Term'];
  const result = [];

  for (let i = 0; i < collections.length; i++) {
    const snap = await db.collection(collections[i])
      .where('deleted', '==', false)
      .get();

    let totalSize = 0;
    let totalThumbSize = 0;
    let count = 0;
    const trainMap = {};
    const coachMap = {};
    const contractorMap = {};
    const dateCounts = {};

    snap.docs.forEach(doc => {
      const d = doc.data();
      count++;
      totalSize += d.compressedSize || d.originalSize || 0;
      totalThumbSize += d.thumbnailSize || 0;
      if (d.trainNumber) trainMap[d.trainNumber] = (trainMap[d.trainNumber] || 0) + 1;
      if (d.coach) coachMap[d.coach] = (coachMap[d.coach] || 0) + 1;
      if (d.contractor) contractorMap[d.contractor] = (contractorMap[d.contractor] || 0) + 1;
      if (d.uploadedAt?.toDate) {
        const dStr = d.uploadedAt.toDate().toISOString().split('T')[0];
        dateCounts[dStr] = (dateCounts[dStr] || 0) + 1;
      }
    });

    result.push({
      tier: tierNames[i],
      collection: collections[i],
      totalDocuments: count,
      totalStorageBytes: totalSize,
      totalThumbnailBytes: totalThumbSize,
      totalBytes: totalSize + totalThumbSize,
      totalMB: ((totalSize + totalThumbSize) / (1024 * 1024)).toFixed(2),
      uniqueTrains: Object.keys(trainMap).length,
      uniqueCoaches: Object.keys(coachMap).length,
      uniqueContractors: Object.keys(contractorMap).length,
      dailyCounts: dateCounts
    });
  }

  const totalBytes = result.reduce((s, r) => s + (r.totalBytes || 0), 0);

  return {
    tiers: result,
    totals: {
      totalDocuments: result.reduce((s, r) => s + r.totalDocuments, 0),
      totalMB: (totalBytes / (1024 * 1024)).toFixed(2),
      totalGB: (totalBytes / (1024 * 1024 * 1024)).toFixed(4)
    },
    forecast: await forecastStorageGrowth(db)
  };
}

// Forecast storage growth based on 30-day trend
async function forecastStorageGrowth(db) {
  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  const snap = await db.collection('evidence_metadata')
    .where('uploadedAt', '>=', thirtyDaysAgo)
    .where('deleted', '==', false)
    .get();

  const dailyBytes = {};
  snap.docs.forEach(doc => {
    const d = doc.data();
    if (d.uploadedAt?.toDate) {
      const dStr = d.uploadedAt.toDate().toISOString().split('T')[0];
      dailyBytes[dStr] = (dailyBytes[dStr] || 0) + (d.compressedSize || d.originalSize || 0);
    }
  });

  const days = Object.keys(dailyBytes).sort();
  if (days.length < 2) {
    return { avgDailyMB: 0, projected30DayMB: 0, projected90DayMB: 0, note: 'Insufficient data for forecast' };
  }

  const totalBytesInPeriod = Object.values(dailyBytes).reduce((s, v) => s + v, 0);
  const avgDailyBytes = totalBytesInPeriod / days.length;

  return {
    avgDailyMB: (avgDailyBytes / (1024 * 1024)).toFixed(2),
    projected30DayMB: ((avgDailyBytes * 30) / (1024 * 1024)).toFixed(2),
    projected90DayMB: ((avgDailyBytes * 90) / (1024 * 1024)).toFixed(2),
    basedOnDays: days.length
  };
}

// === BACKUP OPERATIONS ===

export async function performBackup(db, type = 'daily') {
  const timestamp = new Date().toISOString();
  const collections = ['evidence_metadata', 'archive_evidence', 'long_term_evidence'];
  const backup = { timestamp, type, collections: {} };

  for (const coll of collections) {
    const snap = await db.collection(coll).get();
    backup.collections[coll] = snap.docs.map(d => ({ id: d.id, data: d.data() }));
  }

  const backupRef = db.collection('backup_logs').doc();
  await backupRef.set({
    backupId: backupRef.id,
    type,
    timestamp: admin.firestore.FieldValue.serverTimestamp(),
    recordCounts: Object.fromEntries(
      Object.entries(backup.collections).map(([k, v]) => [k, v.length])
    ),
    status: 'completed'
  });

  return {
    backupId: backupRef.id,
    type,
    timestamp,
    recordCounts: Object.fromEntries(
      Object.entries(backup.collections).map(([k, v]) => [k, v.length])
    ),
    status: 'completed'
  };
}

// === EMAIL REPORTING ===

export function buildEmailReport(role, reportType, params = {}) {
  const roleEmails = {
    rsa: 'rsa@railway.gov.in',
    cm: params.cmEmail || 'cm@contractor.com',
    ca: params.caEmail || 'ca@contractor.com',
    cts: params.ctsEmail || 'cts@train.com',
    cs: params.csEmail || 'cs@supervisor.com'
  };

  const routing = {
    rsa: {
      reports: ['division_weekly', 'contractor_comparison', 'compliance_overview'],
      to: [roleEmails.rsa]
    },
    cm: {
      reports: ['contractor_performance', 'multi_train_summary'],
      to: [roleEmails.cm]
    },
    ca: {
      reports: ['contractor_wide', 'manpower', 'compliance'],
      to: [roleEmails.ca]
    },
    cts: {
      reports: ['train_daily', 'train_exceptions', 'train_complaints'],
      to: [roleEmails.cts]
    },
    cs: {
      reports: ['supervisor_review', 'pending_validations', 'rejected_evidence'],
      to: [roleEmails.cs]
    }
  };

  return routing[role] || routing.rsa;
}

export function getRecipientsForRole(role, userDoc = null) {
  const map = {
    rsa: userDoc?.email || 'rsa@railway.gov.in',
    cm: userDoc?.email || 'cm@contractor.com',
    ca: userDoc?.email || 'ca@contractor.com',
    cts: userDoc?.email || 'cts@train.com',
    cs: userDoc?.email || 'cs@supervisor.com',
    janitor: userDoc?.email || 'janitor@worker.com',
    coach_attendant: userDoc?.email || 'attendant@worker.com'
  };
  return [map[role] || userDoc?.email || 'user@unknown.com'];
}

export const evidenceUtils = {
  EVIDENCE_TYPES,
  STORAGE_TIER,
  buildStoragePath,
  computeFileHash,
  findDuplicateByHash,
  processImage,
  uploadToStorage,
  buildEvidenceMetadata,
  logAudit,
  archiveOldRecords,
  moveToLongTerm,
  restoreFromArchive,
  getEvidenceById,
  searchEvidence,
  generatePDFReport,
  getStorageAnalytics,
  performBackup,
  buildEmailReport,
  getRecipientsForRole
};

export default evidenceUtils;
