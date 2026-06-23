import sgMail from '@sendgrid/mail';
import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import admin from 'firebase-admin';
import jwt from 'jsonwebtoken';
import cron from 'node-cron';
import { Resend } from 'resend';
import twilio from 'twilio';
import ExcelJS from 'exceljs';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import PDFDocument from 'pdfkit';
import fs from 'fs';
import path from 'path';
dotenv.config();
import axios from 'axios';
import { RekognitionClient, CompareFacesCommand, DetectLabelsCommand, DetectFacesCommand } from "@aws-sdk/client-rekognition";
import sharp from 'sharp';
import * as evidence from './evidence_manager.js';
import { PDFRenderer } from './pdf_renderer.js';
import { ExcelGenerator } from './excel_generator.js';
import { ReportService } from './report_service.js';

const TWO_FACTOR_API_KEY = process.env.TWOF_API_KEY;
let serviceAccount;
if (process.env.FIREBASE_SERVICE_ACCOUNT) {
  serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
} else {
  try {
    serviceAccount = await import('./serviceAccountKey.json', { with: { type: 'json' } }).then(m => m.default);
  } catch {
    console.error('❌ CRITICAL: Firebase credentials not found.');
    console.error('   Set FIREBASE_SERVICE_ACCOUNT env var or place serviceAccountKey.json in the project root.');
    process.exit(1);
  }
}

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const validateEnvironment = () => {
  const required = {
    'JWT_SECRET': 'JWT secret for authentication',
    'RESEND_API_KEY': 'Email service API key',
    'TWOF_API_KEY': '2Factor.in API key for SMS OTP',
    'AWS_ACCESS_KEY_ID': 'AWS access key for Rekognition',
    'AWS_SECRET_ACCESS_KEY': 'AWS secret key for Rekognition'
  };

  const missing = [];
  const warnings = [];

  for (const [key, description] of Object.entries(required)) {
    if (!process.env[key]) {
      missing.push(`${key} (${description})`);
    }
  }

  if (!process.env.TWILIO_ACCOUNT_SID) {
    warnings.push('TWILIO_ACCOUNT_SID - SMS features will not work');
  }
  if (!process.env.SENDGRID_API_KEY) {
    warnings.push('SENDGRID_API_KEY - Some email features may not work');
  }

  if (missing.length > 0) {
    console.error('❌ CRITICAL: Missing required environment variables:');
    missing.forEach(item => console.error(`   - ${item}`));
    console.error('\nPlease add these to your .env file and restart the server.');
    process.exit(1);
  }

  if (warnings.length > 0) {
    console.warn('⚠️ WARNING: Optional environment variables missing:');
    warnings.forEach(item => console.warn(`   - ${item}`));
  }

  console.log('✅ Environment validation passed');
};

validateEnvironment();

console.log("--- DEBUGGING .env VARIABLES ---");
const accountSid = process.env.TWILIO_ACCOUNT_SID;
const authToken = process.env.TWILIO_AUTH_TOKEN;
const client = accountSid && authToken ? twilio(accountSid, authToken) : null;
if (process.env.SENDGRID_API_KEY) {
  sgMail.setApiKey(process.env.SENDGRID_API_KEY);
}


const app = express();
const port = process.env.PORT || 5000;
app.use(cors());
app.use(express.json());

const db = admin.firestore();

// Firestore-backed OTP store (survives restarts, auto-expires via TTL)
const otpStore = {
  _col: () => db.collection('_otpStore'),
  async set(key, value) {
    await this._col().doc(key).set({
      value,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      expireAt: new Date(Date.now() + 300000).toISOString()
    });
  },
  async get(key) {
    const doc = await this._col().doc(key).get();
    if (!doc.exists) return undefined;
    const data = doc.data();
    if (data.expireAt && new Date(data.expireAt) < new Date()) {
      await this._col().doc(key).delete();
      return undefined;
    }
    return data.value;
  },
  async delete(key) {
    await this._col().doc(key).delete();
  }
};
const resend = new Resend(process.env.RESEND_API_KEY);
// =======================================================
// == MIDDLEWARE: Verify Token & Fetch User Details from DB
// =======================================================
const verifyToken = async (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).send({ error: 'No token provided.' });
  }

  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    const userDoc = await db.collection('users').doc(decoded.uid).get();

    if (!userDoc.exists) {
      return res.status(401).send({ error: 'User no longer exists.' });
    }

    const userData = userDoc.data();

    let entityDetails = null;
    if (userData.entityId) {
      const entityDoc = await db.collection('entities').doc(userData.entityId).get();
      if (entityDoc.exists) {
        entityDetails = entityDoc.data();
      }
    }

    req.user = {
      uid: decoded.uid,
      email: decoded.email,
      role: decoded.role,
      fullName: userData.fullName,
      name: userData.fullName,
      zone: userData.zone,
      division: userData.division,
      depot: userData.depot,
      userType: userData.userType,
      entityId: userData.entityId,
      entityName: entityDetails?.companyName || null,
      entityDetails: entityDetails
    };

    next();

  } catch (err) {
    console.error('Middleware Error:', err);
    return res.status(403).send({ error: 'Invalid or expired token.' });
  }
};

app.get('/', (req, res) => {
  res.send('Hello, Your  API server is running.');
});

async function getDailyReportData() {
  const today = new Date();
  const startOfDay = new Date(today.setHours(0, 0, 0, 0));
  const endOfDay = new Date(today.setHours(23, 59, 59, 999));

  try {
    const adminSnapshot = await admin.firestore().collection('users')
      .where('role', '==', 'admin')
      .get();

    if (adminSnapshot.empty) {
      console.log("No admins found in the database.");
      return;
    }

    for (const doc of adminSnapshot.docs) {
      const adminInfo = doc.data();
      const adminDivision = adminInfo.division;
      const adminEmail = adminInfo.email;

      if (!adminDivision) {
        console.log(`Skipping admin ${adminEmail} because division is not set.`);
        continue;
      }

      const coachSnapshot = await admin.firestore().collection('coachForms')
        .where('submittedByDivision', '==', adminDivision)
        .where('createdAt', '>=', startOfDay)
        .where('createdAt', '<=', endOfDay)
        .get();

      const coachData = coachSnapshot.docs.map(d => d.data());

      const premisesSnapshot = await admin.firestore().collection('premisesForms')
        .where('submittedByDivision', '==', adminDivision)
        .where('createdAt', '>=', startOfDay)
        .where('createdAt', '<=', endOfDay)
        .get();

      const premisesData = premisesSnapshot.docs.map(d => d.data());

      // Always send email, even if no data
      const coachBuffer = coachData.length > 0 ? await generateCoachExcelBuffer(coachData) : null;
      const premisesBuffer = premisesData.length > 0 ? await generatePremisesExcelBuffer(premisesData) : null;

      await sendEmailWithAttachments(adminEmail, adminDivision, coachBuffer, premisesBuffer, {
        coachCount: coachData.length,
        premisesCount: premisesData.length
      });
    }
  } catch (error) {
    console.error("Error in Automated Reporting Loop:", error);
  }
}

// --- COACH EXCEL GENERATOR (Exact mapping from your Firestore) ---
async function generateCoachExcelBuffer(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Coach Cleaning Report');

  // Header Styling
  const headerStyle = {
    font: { bold: true, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } },
    alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
    border: { outline: { style: 'thin' } }
  };

  // Define Columns
  sheet.columns = [
    { header: 'Date', key: 'date', width: 15 },
    { header: 'Train Details', key: 'train', width: 30 },
    { header: 'Work Type', key: 'workType', width: 20 },
    { header: 'ACWP Status', key: 'acwp', width: 15 },
    { header: 'Total Penalty', key: 'penalty', width: 15 },
    { header: 'Internal (A)', key: 'intA', width: 12 },
    { header: 'Intensive (NA)', key: 'intNA', width: 12 },
    { header: 'Status', key: 'status', width: 15 }
  ];

  // Apply Style to Header Row
  sheet.getRow(1).eachCell((cell) => { cell.style = headerStyle; });

  // Add Data Rows
  data.forEach(item => {
    sheet.addRow({
      date: item.formDateTime ? item.formDateTime.split('T')[0] : 'N/A',
      train: `${item.submittedTo?.trainNumber || ''} - ${item.submittedTo?.trainName || ''}`,
      workType: item.ratingDetails?.workType || 'N/A',
      acwp: item.ratingDetails?.acwpStatus || 'N/A',
      penalty: item.summary?.totalPenalty || 0,
      intA: item.summary?.internal?.A || 0,
      intNA: item.summary?.intensive?.NA || 0,
      status: item.status || 'LOCKED'
    });
  });

  return await workbook.xlsx.writeBuffer();
}

// --- PREMISES EXCEL GENERATOR (Exact mapping from your Firestore) ---
async function generatePremisesExcelBuffer(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Premises Report');

  const headerStyle = {
    font: { bold: true, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '4472C4' } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };

  sheet.columns = [
    { header: 'Date', key: 'date', width: 15 },
    { header: 'Location', key: 'location', width: 20 },
    { header: 'Area (Sq Mtrs)', key: 'area', width: 15 },
    { header: 'Housekeeping Score', key: 'hkScore', width: 20 },
    { header: 'Pit-Line Score', key: 'pitScore', width: 15 },
    { header: 'Garbage Score', key: 'gbScore', width: 15 },
    { header: 'Overall Score', key: 'overall', width: 15 },
    { header: 'Status', key: 'status', width: 15 }
  ];

  sheet.getRow(1).eachCell((cell) => { cell.style = headerStyle; });

  data.forEach(item => {
    sheet.addRow({
      date: item.formDateTime ? item.formDateTime.split('T')[0] : 'N/A',
      location: item.location || 'N/A',
      area: item.area || 0,
      hkScore: item.summary?.housekeepingScore || '0',
      pitScore: item.summary?.pitLineScore || '0',
      gbScore: item.summary?.garbageDisposalScore || '0',
      overall: item.summary?.overallScore || '0',
      status: item.status || 'N/A'
    });
  });

  return await workbook.xlsx.writeBuffer();
}

// --- STEP 4.1: FINAL EMAIL LOGIC WITH VERIFIED DOMAIN ---
async function sendEmailWithAttachments(email, division, coachBuffer, premisesBuffer, stats = { coachCount: 0, premisesCount: 0 }) {
  const todayDate = new Date().toISOString().split('T')[0];

  try {
    const hasData = stats.coachCount > 0 || stats.premisesCount > 0;

    await resend.emails.send({
      // Using your verified domain
      from: 'Swachh Railways <reports@swachhrailways.com>',
      to: [email],
      subject: `Daily Cleaning Report | ${division} | ${todayDate}`,
      html: `
                <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 8px;">
                    <h2 style="color: #1F4E78; border-bottom: 2px solid #1F4E78; padding-bottom: 10px;">Daily Summary Report</h2>
                    
                    <p>Dear Admin,</p>
                    
                    <p>Please find the automated cleaning performance reports for the <b>${division}</b> division.</p>
                    
                    <div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #1F4E78; margin: 20px 0;">
                        <p style="margin: 0; font-weight: bold;">Summary for ${todayDate}:</p>
                        <ul style="margin: 10px 0 0 0;">
                            <li>Coach Forms: ${stats.coachCount}</li>
                            <li>Premises Forms: ${stats.premisesCount}</li>
                        </ul>
                    </div>

                    ${hasData ? `
                        <div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #1F4E78; margin: 20px 0;">
                            <p style="margin: 0; font-weight: bold;">Attached Files:</p>
                            <ul style="margin: 10px 0 0 0;">
                                ${stats.coachCount > 0 ? '<li>Coach Cleaning Detailed Report (.xlsx)</li>' : ''}
                                ${stats.premisesCount > 0 ? '<li>Premises Cleaning Detailed Report (.xlsx)</li>' : ''}
                            </ul>
                        </div>
                    ` : `
                        <div style="background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0;">
                            <p style="margin: 0; color: #856404;">
                                ⚠️ No forms were submitted today. Please check if forms are being created properly.
                            </p>
                        </div>
                    `}
                    
                    <p style="margin-top: 30px; font-size: 0.85em; color: #888; border-top: 1px solid #eee; padding-top: 15px;">
                        <i>This is an automated message generated by the Swachh Railways System. Please do not reply to this email.</i>
                    </p>
                    
                    <p style="font-weight: bold; color: #1F4E78; margin-bottom: 0;">Swachh Railways Management System</p>
                    <p style="font-size: 0.8em; color: #666; margin-top: 4px;">Powered by Backend Automation</p>
                </div>
            `,
      attachments: [
        ...(coachBuffer ? [{ filename: `Coach_Report_${division}_${todayDate}.xlsx`, content: coachBuffer.toString('base64') }] : []),
        ...(premisesBuffer ? [{ filename: `Premises_Report_${division}_${todayDate}.xlsx`, content: premisesBuffer.toString('base64') }] : [])
      ],
    });
    console.log(`Report successfully emailed to Admin: ${email} (${division})`);
  } catch (error) {
    console.error(`Error sending email to ${email}:`, error);
  }
}


// =======================================================
// == MULTER & MEDIA UPLOAD CONFIGURATION (ES MODULES)
// =======================================================

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 5 * 1024 * 1024 // Limit: 5MB max size
  },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) {
      cb(null, true);
    } else {
      cb(new Error('Only image files are allowed!'), false);
    }
  }
});

// Firebase admin pehle se upar initialized hai, direct use karein
const bucket = admin.storage().bucket('swachh-railways.firebasestorage.app');

// =======================================================
// == API 4.10: Upload Camera Image to Firebase Storage (Fixed & Multi-Key Compatible)
// =======================================================
app.post('/api/media/upload', verifyToken, (req, res, next) => {
  upload.fields([{ name: 'file', maxCount: 1 }, { name: 'image', maxCount: 1 }])(req, res, (err) => {
    if (err instanceof multer.MulterError) {
      console.error('(Multer Error):', err);
      return res.status(400).send({ error: 'Multipart form error', details: err.message });
    } else if (err) {
      console.error('(Upload Error):', err);
      return res.status(400).send({ error: err.message });
    }

    if (req.files) {
      req.file = (req.files['file'] && req.files['file'][0]) || (req.files['image'] && req.files['image'][0]);
    }
    next();
  });
}, async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send({ error: "No file uploaded. Please attach an image." });
    }

    const originalExt = req.file.originalname.split('.').pop() || 'jpg';
    const uniqueToken = uuidv4();
    const fileName = `obhs_tasks/${uuidv4()}_${Date.now()}.${originalExt}`;
    const fileUpload = bucket.file(fileName);

    const blobStream = fileUpload.createWriteStream({
      metadata: {
        contentType: req.file.mimetype,
        metadata: {
          firebaseStorageDownloadTokens: uniqueToken
        }
      }
    });

    blobStream.on('error', (error) => {
      console.error('(Media Storage) Stream Upload Error:', error);
      if (!res.headersSent) {
        return res.status(500).send({ error: 'Upload streaming failed', details: error.message });
      }
    });

    blobStream.on('finish', async () => {
      try {
        const publicUrl = `https://firebasestorage.googleapis.com/v0/b/${bucket.name}/o/${encodeURIComponent(fileName)}?alt=media&token=${uniqueToken}`;

        console.log(`(Media Storage) File uploaded successfully and activated: ${fileName}`);

        return res.status(200).json({
          success: true,
          message: "Image uploaded successfully.",
          imageUrl: publicUrl
        });

      } catch (innerErr) {
        console.error('(Media Storage) Fallback generation error:', innerErr);
        const fallbackUrl = `https://storage.googleapis.com/${bucket.name}/${fileName}`;
        return res.status(200).json({
          success: true,
          message: "Image uploaded (via fallback asset channel).",
          imageUrl: fallbackUrl
        });
      }
    });

    blobStream.end(req.file.buffer);

  } catch (error) {
    console.error('(Media Upload) High-Level Catch Triggered:', error);
    if (!res.headersSent) {
      res.status(500).send({ error: 'Failed to process media upload', details: error.message });
    }
  }
});

// =======================================================
// == 1. AUTHENTICATION APIs (2FACTOR.IN VERSION)
// =======================================================

// 1. Send Mobile OTP
app.post('/api/auth/send-otp', async (req, res) => {
  try {
    const { phone } = req.body;
    if (!phone) return res.status(400).json({ error: 'Phone number is required' });

    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    otpStore.set(phone, otp);

    const url = `https://2factor.in/API/V1/${TWO_FACTOR_API_KEY}/SMS/91${phone}/${otp}`;

    const response = await axios.get(url);

    if (response.data.Status === "Success") {
      console.log(`(2Factor) OTP ${otp} sent to ${phone}`);
      res.status(200).json({ success: true, message: "OTP has been sent to your mobile number." });
    } else {
      throw new Error(response.data.Details || "Failed to send SMS via 2Factor");
    }

  } catch (error) {
    console.error('(2Factor) ERROR SENDING OTP:', error.message);
    res.status(500).json({ error: 'Failed to send OTP', details: error.message });
  }
});

// ==========================================
// 2. VERIFY MOBILE OTP & LOGIN (2FACTOR LOGIC)
// ==========================================
app.post('/api/auth/verify-otp', async (req, res) => {
  try {
    const { phone, otp } = req.body;

    // 1. Basic Validation
    if (!phone || !otp) {
      return res.status(400).json({ error: "Phone number and OTP are required." });
    }

    const storedOtp = await otpStore.get(phone);
    if (!storedOtp) {
      return res.status(400).json({ error: "OTP expired or not requested. Please try again." });
    }

    if (storedOtp !== otp) {
      return res.status(400).json({ error: "Invalid OTP. Please check and try again." });
    }

    await otpStore.delete(phone);

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('mobile', '==', phone).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).json({ error: "This mobile number is not registered in our database." });
    }

    const userData = snapshot.docs[0].data();

    // 4. Status Check: Only 'APPROVED' users can login
    if (userData.status !== 'APPROVED') {
      return res.status(403).json({
        error: `Your account status is ${userData.status}. Please contact Admin for access.`
      });
    }

    // 5. Contractor/Entity Logic (Optional Enrichment)
    let entityDetails = null;
    if (userData.userType === 'contractor' && userData.entityId) {
      const entityDoc = await db.collection('entities').doc(userData.entityId).get();
      if (entityDoc.exists) {
        entityDetails = entityDoc.data();
        userData.entityDetails = entityDetails;
      }
    }

    const token = jwt.sign(
      {
        uid: userData.uid,
        role: userData.role,
        userType: userData.userType,
        fullName: userData.fullName,
        email: userData.email,
        mobile: userData.mobile,
        zone: userData.zone,
        division: userData.division,
        depot: userData.depot,
        entityId: userData.entityId,
        entityDetails: entityDetails
      },
      process.env.JWT_SECRET,
      { expiresIn: '15d' }
    );

    delete userData.password;

    console.log(`(Login) Success via 2Factor OTP for ${phone}`);

    res.status(200).json({
      success: true,
      message: "Login Successful",
      token: token,
      user: userData
    });

  } catch (error) {
    console.error('(Login) Failed to verify Mobile OTP:', error);
    res.status(500).json({
      error: 'Internal Server Error',
      details: error.message
    });
  }
});


// =======================================================
// == PASSENGER VERIFICATION APIs FOR FEEDBACK PAGE (NO LOGIN)
// =======================================================

// 1. Send OTP to Passenger
app.post('/api/passenger/send-otp', async (req, res) => {
  try {
    const { phone } = req.body;
    if (!phone) {
      return res.status(400).json({ error: 'Phone number is required' });
    }

    // Generate 6-digit random OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    otpStore.set(phone, otp);

    // Hit 2Factor.in API to deliver SMS
    const url = `https://2factor.in/API/V1/${TWO_FACTOR_API_KEY}/SMS/91${phone}/${otp}`;
    const response = await axios.get(url);

    if (response.data.Status === "Success") {
      console.log(`(Passenger Verification) OTP ${otp} sent to ${phone}`);
      return res.status(200).json({
        success: true,
        message: "OTP has been sent to passenger mobile number."
      });
    } else {
      throw new Error(response.data.Details || "Failed to send SMS via 2Factor");
    }

  } catch (error) {
    console.error('(Passenger OTP) ERROR SENDING OTP:', error.message);
    return res.status(500).json({ error: 'Failed to send OTP', details: error.message });
  }
});

// 2. Verify Passenger OTP
app.post('/api/passenger/verify-otp', async (req, res) => {
  try {
    const { phone, otp } = req.body;

    // Basic Validation
    if (!phone || !otp) {
      return res.status(400).json({ error: "Phone number and OTP are required." });
    }

    // Check if OTP exists in store
    const storedOtp = await otpStore.get(phone);
    if (!storedOtp) {
      return res.status(400).json({ error: "OTP expired or not requested. Please try again." });
    }

    // Brute force protection: track attempts
    const attemptKey = `ATTEMPT_${phone}`;
    const attempts = (await otpStore.get(attemptKey)) || 0;
    if (attempts >= 5) {
      return res.status(429).json({ error: "Too many attempts. Please request a new OTP." });
    }

    // Validate OTP match
    if (storedOtp !== otp) {
      await otpStore.set(attemptKey, attempts + 1);
      return res.status(400).json({ error: "Invalid OTP. Please check and try again." });
    }

    // Clear OTP and attempt tracking after successful verification
    await Promise.all([
      otpStore.delete(phone),
      otpStore.delete(attemptKey)
    ]);

    console.log(`(Passenger Verification) Success for mobile number: ${phone}`);

    // Return simple success payload without login token or DB object
    return res.status(200).json({
      success: true,
      message: "Passenger mobile number verified successfully."
    });

  } catch (error) {
    console.error('(Passenger OTP Verification) Failed:', error);
    return res.status(500).json({
      error: 'Internal Server Error',
      details: error.message
    });
  }
});

// 3. Create Passenger Task
app.post('/api/passenger/create-task', async (req, res) => {
  try {
    const { trainNo, coachNo, seatNo, taskType, description } = req.body;

    if (!trainNo || !coachNo || !taskType) {
      return res.status(400).json({ error: "Train No, Coach No, and Task Type are required." });
    }

    const taskId = `pass_${Date.now()}`;
    const taskData = {
      taskId,
      trainNo,
      coachNo,
      seatNo: seatNo || 'N/A',
      taskType,
      description: description || '',
      status: 'OPEN',
      source: 'PASSENGER',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await db.collection('passenger_tasks').doc(taskId).set(taskData);

    console.log(`(Passenger) Task Created: ${taskType} for Train ${trainNo}, Coach ${coachNo}`);

    return res.status(201).json({
      success: true,
      message: "Task created successfully.",
      taskId
    });

  } catch (error) {
    console.error('(Passenger Task Creation) Failed:', error);
    return res.status(500).json({
      error: 'Internal Server Error',
      details: error.message
    });
  }
});

// 5. Create Emergency Task (CTS Generated)
app.post('/api/cts/create-emergency-task', verifyToken, async (req, res) => {
  try {
    const { trainNo, coachNo, taskType, description, priority } = req.body;

    // Role check: Only CTS/Supervisor can create emergency tasks
    const allowedRoles = ['cts', 'railway supervisor', 'railway admin'];
    const userRole = (req.user.role || '').toLowerCase();
    if (!allowedRoles.includes(userRole)) {
      return res.status(403).json({ error: 'Only CTS/Supervisors can create emergency tasks' });
    }

    if (!trainNo || !coachNo || !taskType) {
      return res.status(400).json({ error: "Train No, Coach No, and Task Type are required." });
    }

    const taskId = `emg_${Date.now()}`;
    const taskData = {
      taskId,
      trainNo,
      coachNo,
      taskType,
      description: description || 'Emergency Task',
      priority: priority || 'HIGH',
      status: 'OPEN',
      taskSource: 'CTS',
      createdBy: req.user.uid,
      createdByName: req.user.fullName,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await db.collection('passenger_tasks').doc(taskId).set(taskData); // Re-using passenger_tasks collection for simplified worker view or create a unified task_exceptions collection

    console.log(`(CTS Emergency) Task Created: ${taskType} for Train ${trainNo}, Coach ${coachNo}`);

    return res.status(201).json({
      success: true,
      message: "Emergency task created successfully.",
      taskId
    });

  } catch (error) {
    console.error('(CTS Emergency Task Creation) Failed:', error);
    return res.status(500).json({ error: 'Internal Server Error', details: error.message });
  }
});
app.get('/api/passenger/tasks', verifyToken, async (req, res) => {
  try {
    const { trainNo } = req.query;

    let query = db.collection('passenger_tasks');

    if (trainNo) {
      query = query.where('trainNo', '==', trainNo);
    }

    const snapshot = await query.get();

    const tasks = [];
    snapshot.forEach(doc => {
      tasks.push({ uid: doc.id, ...doc.data() });
    });

    tasks.sort((a, b) => {
      const dateA = a.createdAt ? new Date(a.createdAt) : new Date(0);
      const dateB = b.createdAt ? new Date(b.createdAt) : new Date(0);
      return dateB - dateA;
    });

    return res.status(200).json({
      success: true,
      count: tasks.length,
      tasks
    });

  } catch (error) {
    console.error('(Passenger Tasks Retrieval) Failed:', error);
    return res.status(500).json({ error: 'Internal Server Error', details: error.message });
  }
});

// API: Get Active Coaches for a Train (For Passenger App Dropdown)
app.get('/api/passenger/train/:trainNo/coaches', async (req, res) => {
  try {
    const { trainNo } = req.params;

    // Find an active run instance for this train
    const runsSnap = await db.collection('RunInstance')
      .where('trainNo', '==', trainNo)
      .where('status', 'in', ['ACTIVE', 'RUNNING', 'ALLOCATED', 'READY'])
      .get();

    if (runsSnap.empty) {
      return res.status(200).json({ success: true, coaches: [] });
    }

    // Combine coaches from all active instances (usually there's only 1 or 2 overlapping)
    const coachSet = new Set();
    runsSnap.forEach(doc => {
      const runData = doc.data();
      if (runData.coaches && Array.isArray(runData.coaches)) {
        runData.coaches.forEach(c => {
          if (c.coachPosition) coachSet.add(c.coachPosition);
          else if (c.coachNumber) coachSet.add(c.coachNumber);
        });
      }
    });

    const sortedCoaches = Array.from(coachSet).sort();
    return res.status(200).json({ success: true, coaches: sortedCoaches });
  } catch (error) {
    console.error('(Passenger Coaches Fetch) Failed:', error);
    return res.status(500).json({ error: 'Internal Server Error', details: error.message });
  }
});

// 1. Send Email OTP
app.post('/api/auth/send-email-otp', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) return res.status(400).json({ error: "Email is required." });

    const userSnapshot = await db.collection('users').where('email', '==', email).limit(1).get();
    if (userSnapshot.empty) {
      return res.status(404).json({ error: "This email is not registered." });
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    otpStore.set(email, otp);
    console.log(`(Resend) Generated OTP ${otp} for ${email}`);

    const { data, error } = await resend.emails.send({
      from: 'Swachh Railways <auth@swachhrailways.com>', // Aapka verified domain
      to: [email],
      subject: 'Login OTP - Swachh Railways',
      html: `
        <div style="font-family: Arial, sans-serif; border: 1px solid #ddd; padding: 20px; border-radius: 10px;">
          <h2 style="color: #2c3e50;">Verify Your Login</h2>
          <p>Hello,</p>
          <p>Your One-Time Password (OTP) for logging into Swachh Railways is:</p>
          <h1 style="color: #e67e22; letter-spacing: 5px;">${otp}</h1>
          <p>This code is valid for <b>5 minutes</b>. Please do not share this with anyone.</p>
          <hr style="border: 0; border-top: 1px solid #eee;" />
          <p style="font-size: 12px; color: #7f8c8d;">If you didn't request this, please ignore this email.</p>
        </div>
      `,
    });

    if (error) {
      console.error('(Resend) ERROR:', error);
      return res.status(400).json({ error: "Failed to send email via Resend", details: error });
    }

    res.status(200).json({ message: "OTP has been sent to your email." });

  } catch (error) {
    console.error('(Resend) SERVER ERROR:', error);
    res.status(500).json({ error: 'Failed to process request', details: error.message });
  }
});

app.post('/api/auth/verify-email-otp', async (req, res) => {
  try {
    const { email, otp } = req.body;

    const storedOtp = await otpStore.get(email);
    if (!storedOtp) {
      return res.status(400).json({ error: "OTP expired or not requested. Please try again." });
    }

    if (storedOtp !== otp) {
      return res.status(400).json({ error: "Invalid OTP. Please check and try again." });
    }

    await otpStore.delete(email);

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('email', '==', email).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).json({ error: "Record not found in database." });
    }

    const userData = snapshot.docs[0].data();

    if (userData.status !== 'APPROVED') {
      return res.status(403).json({ error: `Your account is ${userData.status}. Contact admin.` });
    }

    let entityDetails = null;
    if (userData.userType === 'contractor' && userData.entityId) {
      const entityDoc = await db.collection('entities').doc(userData.entityId).get();
      if (entityDoc.exists) {
        entityDetails = entityDoc.data();
        userData.entityDetails = entityDetails;
      }
    }

    const token = jwt.sign(
      {
        uid: userData.uid,
        role: userData.role,
        userType: userData.userType,
        fullName: userData.fullName,
        email: userData.email,
        zone: userData.zone,
        division: userData.division,
        depot: userData.depot,
        entityId: userData.entityId,
        entityDetails: entityDetails
      },
      process.env.JWT_SECRET,
      { expiresIn: '7d' }
    );

    delete userData.password;

    console.log(`(Login) Success via Resend OTP for ${email}`);

    res.status(200).json({
      message: "Login Successful",
      token: token,
      user: userData
    });

  } catch (error) {
    console.error('(Login) Failed to verify Email OTP:', error);
    res.status(500).json({ error: 'Internal Server Error', details: error.message });
  }
});


// --- Method C: Email + Password (Traditional) ---
app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) {
      return res.status(400).send({ error: "Email and Password are required." });
    }

    const normalizedEmail = email.trim().toLowerCase();

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('email', '==', normalizedEmail).limit(1).get();

    if (snapshot.empty) {
      console.log(`(Login) Failed login: Email not found ${normalizedEmail}`);
      return res.status(401).send({ error: "Invalid credentials." });
    }

    const userData = snapshot.docs[0].data();

    if (userData.password !== password) {
      console.log(`(Login) Failed login: Password mismatch for ${normalizedEmail}`);
      return res.status(401).send({ error: "Invalid credentials." });
    }

    if (userData.status !== 'APPROVED') {
      console.log(`(Login) Failed login for ${normalizedEmail}: Status is ${userData.status}`);
      return res.status(403).send({ error: `Your account status is: ${userData.status}. Contact admin.` });
    }

    let entityDetails = null;
    if (userData.userType === 'contractor' && userData.entityId) {
      const entityDoc = await db.collection('entities').doc(userData.entityId).get();
      if (entityDoc.exists) {
        entityDetails = entityDoc.data();
        userData.entityDetails = entityDetails;
      }
    }

    let activeRunInstanceId = null;
    try {
      if (userData.uid) {

        const runInstanceSnapshot = await db.collection('RunInstance')
          .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running'])
          .get();

        for (const doc of runInstanceSnapshot.docs) {
          const runData = doc.data();
          if (runData.coaches && Array.isArray(runData.coaches)) {
            const isWorkerAssigned = runData.coaches.some(coach => coach.workerId === userData.uid);
            if (isWorkerAssigned) {
              activeRunInstanceId = runData.runInstanceId || doc.id;
              break;
            }
          }
        }
      }
    } catch (runError) {

      console.error('(Login) Optional RunInstance fetch failed:', runError);
    }

    userData.activeRunInstanceId = activeRunInstanceId;

    const customAppToken = jwt.sign(
      {
        uid: userData.uid,
        role: userData.role,
        userType: userData.userType,
        fullName: userData.fullName,
        email: userData.email,
        zone: userData.zone,
        division: userData.division,
        depot: userData.depot,
        entityId: userData.entityId,
        entityDetails: entityDetails,
        activeRunInstanceId: activeRunInstanceId
      },
      process.env.JWT_SECRET,
      { expiresIn: '7d' }
    );

    console.log(`(Login) Successful login for ${normalizedEmail}, ActiveRunInstance: ${activeRunInstanceId}`);
    delete userData.password;

    res.status(200).json({
      message: "Login Successful",
      token: customAppToken,
      user: userData
    });

  } catch (error) {
    console.error('(Login) Error during login:', error);
    res.status(500).send({ error: 'Login failed', details: error.message });
  }
});

app.post('/api/auth/loginWithMobile', async (req, res) => {
  try {
    const { mobile, password } = req.body;
    if (!mobile || !password) {
      return res.status(400).send({ error: "Mobile number and Password are required." });
    }

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('mobile', '==', mobile).limit(1).get();

    if (snapshot.empty) {
      console.log(`(Login) Failed login: Mobile not found ${mobile}`);
      return res.status(401).send({ error: "Invalid credentials." });
    }

    const userData = snapshot.docs[0].data();

    if (userData.password !== password) {
      console.log(`(Login) Failed login: Password mismatch for mobile ${mobile}`);
      return res.status(401).send({ error: "Invalid credentials." });
    }

    if (userData.status !== 'APPROVED') {
      console.log(`(Login) Failed login for mobile ${mobile}: Status is ${userData.status}`);
      return res.status(403).send({ error: `Your account status is: ${userData.status}. Contact admin.` });
    }

    let entityDetails = null;
    if (userData.userType === 'contractor' && userData.entityId) {
      const entityDoc = await db.collection('entities').doc(userData.entityId).get();
      if (entityDoc.exists) {
        entityDetails = entityDoc.data();
        userData.entityDetails = entityDetails;
      }
    }

    const customAppToken = jwt.sign(
      {
        uid: userData.uid,
        role: userData.role,
        userType: userData.userType,
        fullName: userData.fullName,
        email: userData.email,
        mobile: userData.mobile,
        zone: userData.zone,
        division: userData.division,
        depot: userData.depot,
        entityId: userData.entityId,
        entityDetails: entityDetails
      },
      process.env.JWT_SECRET,
      { expiresIn: '7d' }
    );

    console.log(`(Login) Successful login for mobile ${mobile}`);
    delete userData.password;

    res.status(200).json({
      message: "Login Successful",
      token: customAppToken,
      user: userData
    });

  } catch (error) {
    console.error('(Login) Error during mobile login:', error);
    res.status(500).send({ error: 'Login failed', details: error.message });
  }
});


// =======================================================
// == API 1.1: Admin Create User (Fixed: Auth, Case-Insensitive Email)
// =======================================================
app.post('/api/admin/createUser', verifyToken, async (req, res) => {
  try {
    const {
      email, password, role, userType,
      fullName, designation, mobile,
      zone, division, depot,
      entityId, trainId, trainIds, worker_type
    } = req.body;
    const normalizedEmail = email ? email.trim().toLowerCase() : null;

    const { uid: creatorId, name, fullName: creatorNameAuth, role: creatorRole } = req.user;
    const creatorName = creatorNameAuth || name || creatorRole || 'Admin';

    if (!email || !password || !role || !userType) {
      return res.status(400).send({ error: "Email, Password, Role, and UserType are required." });
    }
    const emailQuery = await db.collection('users').where('email', '==', normalizedEmail).get();
    if (!emailQuery.empty) return res.status(400).send({ error: "Email already registered." });

    if (mobile) {
      const mobileQuery = await db.collection('users').where('mobile', '==', mobile).get();
      if (!mobileQuery.empty) return res.status(400).send({ error: "Mobile Number already registered." });
    }

    const roleUpper = role.toUpperCase();
    if (roleUpper === 'CTS' || roleUpper === 'CONTRACTOR SUPERVISOR') {
      if (!division || !trainId) {
        return res.status(400).send({ error: "Division and Train ID are mandatory for Contractor Supervisor." });
      }
      // Strictly one train for Contractor Supervisor
      if (trainIds && trainIds.length > 1) {
        return res.status(400).send({ error: "Contractor Supervisor can only be mapped to ONE train." });
      }
    }

    if (roleUpper === 'RAILWAY SUPERVISOR') {
      if (!division || (!trainId && (!trainIds || trainIds.length === 0))) {
        return res.status(400).send({ error: "Division and at least one Train ID are mandatory for Railway Supervisor." });
      }
    }

    // Worker Bifurcation Check (Janitor vs Attendant)
    const normalizedUserType = userType.toLowerCase();
    if (roleUpper.includes('WORKER')) {
      if (!worker_type || !['Janitor', 'Attendant'].includes(worker_type)) {
        return res.status(400).send({ error: "worker_type (Janitor or Attendant) is mandatory for workers." });
      }
    }

    let entityData = null;
    if (userType.toLowerCase() === 'contractor') {
      if (!entityId) {
        return res.status(400).send({ error: "Contractor users must have an 'entityId' (Company ID)." });
      }

      const entityDoc = await db.collection('entities').doc(entityId).get();
      if (!entityDoc.exists) {
        return res.status(404).send({ error: "Entity (Company) not found." });
      }
      entityData = entityDoc.data();

      const userRoleLower = role.toLowerCase();
      if (userRoleLower.includes('admin') || userRoleLower.includes('supervisor')) {
        if (!zone || !division) {
          return res.status(400).send({ error: "Zone and Division are mandatory to check Active Contracts." });
        }

        console.log(`(CreateUser) Checking Active Contract for Entity: ${entityId}`);

        const contractSnapshot = await db.collection('contracts')
          .where('entityId', '==', entityId)
          .where('zone', '==', zone)
          .where('division', '==', division)
          .where('status', 'in', ['Active', 'active', 'APPROVED'])
          .limit(1)
          .get();

        if (contractSnapshot.empty) {
          return res.status(403).send({
            error: `Cannot create ${role}. No Active Contract found for this Company in ${division} (${zone}).`
          });
        }
      }
    }

    const userRecord = await admin.auth().createUser({
      email: normalizedEmail,
      password: password,
      displayName: fullName,
      disabled: false
    });

    const newUid = userRecord.uid;
    await db.collection('users').doc(newUid).set({
      uid: newUid,
      email: normalizedEmail,
      password: password,
      role,
      userType: userType.toLowerCase(),

      fullName: fullName || null,
      mobile: mobile || null,
      designation: designation || null,
      zone: zone || null,
      division: division || null,
      depot: depot || null,

      entityId: entityId || null,
      entityDetails: entityData,
      trainId: trainId || null,
      trainIds: trainIds || (trainId ? [trainId] : []),
      worker_type: worker_type || null,

      createdBy: creatorId,
      createdByName: creatorName,

      status: 'PENDING',

      createdAt: new Date().toISOString(),
      submitted_at: new Date().toISOString()
    });

    console.log(`(Admin) User Created: ${fullName} by ${creatorName}`);

    res.status(201).send({
      message: 'User created successfully.',
      uid: newUid
    });

  } catch (error) {
    if (error.code === 'auth/email-already-exists') {
      return res.status(400).send({ error: "Email is already in use (Auth)." });
    }
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).send({ error: 'Database Index Missing for Contract Check.' });
    }
    console.error('(Admin) Error creating user:', error);
    res.status(500).send({ error: 'Failed to create user', details: error.message });
  }
});

// =======================================================
// == API 1.2: Admin Update User (Fixed: Name in Response)
// =======================================================
app.put('/api/admin/updateUser/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const {
      fullName, designation, mobile, zone, division, depot,
      role, userType, password, entityId, trainId, trainIds, worker_type
    } = req.body;
    const { uid: editorId, name, fullName: editorAuthName, role: editorRole } = req.user;
    const editorName = editorAuthName || name || editorRole || 'Admin';

    if (!uid) return res.status(400).send({ error: "User ID is required." });

    const userDocRef = db.collection('users').doc(uid);
    const doc = await userDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: "User not found." });

    const currentData = doc.data();
    const targetUserName = fullName || currentData.fullName || 'Unknown User';
    const finalRole = role || currentData.role;
    const finalRoleUpper = finalRole ? finalRole.toUpperCase() : '';
    const finalDivision = division || currentData.division;
    const finalTrainId = trainId || currentData.trainId;
    const finalTrainIds = trainIds || currentData.trainIds;

    if (finalRoleUpper === 'CTS' || finalRoleUpper === 'CONTRACTOR SUPERVISOR') {
      if (!finalDivision || !finalTrainId) {
        return res.status(400).send({ error: "Division and Train ID are mandatory for Contractor Supervisor." });
      }
      if (Array.isArray(finalTrainId) || (finalTrainIds && finalTrainIds.length > 1)) {
        return res.status(400).send({ error: "Contractor Supervisor can only be mapped to ONE train." });
      }
    }

    if (finalRoleUpper === 'RAILWAY SUPERVISOR') {
      if (!finalDivision || (!finalTrainId && (!finalTrainIds || finalTrainIds.length === 0))) {
        return res.status(400).send({ error: "Division and at least one Train ID are mandatory for Railway Supervisor." });
      }
    }

    const updateData = {};
    if (fullName !== undefined) updateData.fullName = fullName;
    if (designation !== undefined) updateData.designation = designation;
    if (mobile !== undefined) updateData.mobile = mobile;
    if (zone !== undefined) updateData.zone = zone;
    if (division !== undefined) updateData.division = division;
    if (depot !== undefined) updateData.depot = depot;
    if (role !== undefined) updateData.role = role;
    if (userType !== undefined) updateData.userType = userType.toLowerCase();
    if (trainId !== undefined) updateData.trainId = trainId;
    if (trainIds !== undefined) updateData.trainIds = trainIds;
    if (worker_type !== undefined) updateData.worker_type = worker_type;

    // RE-APPROVAL LOGIC: Reset to PENDING on edit
    updateData.status = 'PENDING';
    updateData.updatedAt = new Date().toISOString();
    updateData.updatedBy = editorId;
    updateData.updatedByName = editorName;

    if (password) {
      try {
        await admin.auth().updateUser(uid, { password: password });
        console.log(`(Admin) Password updated for user: ${uid}`);
      } catch (authError) {
        console.error('Error updating password in Auth:', authError);
        return res.status(400).send({ error: "Failed to update password. Password must be strong." });
      }
    }

    if (entityId !== undefined) {
      updateData.entityId = entityId;
      if (entityId) {
        const entityDoc = await db.collection('entities').doc(entityId).get();
        if (!entityDoc.exists) {
          return res.status(404).send({ error: "Entity (Company) not found with this ID." });
        }
        updateData.entityDetails = entityDoc.data();
      } else {
        updateData.entityDetails = null;
      }
    }

    updateData.updatedBy = editorId;
    updateData.updatedByName = editorName;
    updateData.updatedAt = new Date().toISOString();

    if (Object.keys(updateData).length === 0 && !password) {
      return res.status(400).send({ error: "No fields to update provided." });
    }

    await userDocRef.update(updateData);

    console.log(`(Admin) User "${targetUserName}" updated by ${editorName}.`);
    res.status(200).send({
      message: `User "${targetUserName}" has been updated successfully.`,
      updates: updateData
    });

  } catch (error) {
    console.error('(Admin) Error updating user:', error);
    res.status(500).send({ error: 'Failed to update user', details: error.message });
  }
});

// --- API 1.3: Master User Approve
app.post('/api/master/approveUser/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;

    const { uid: approverId, name, fullName, role } = req.user;
    const approverName = fullName || name || role || 'Master Admin';

    if (!uid) return res.status(400).send({ error: "User ID is required." });

    const userDocRef = db.collection('users').doc(uid);
    const doc = await userDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: "User not found." });

    const userData = doc.data();
    const userName = userData.fullName || "User";
    await userDocRef.update({
      status: 'APPROVED',

      approvedBy: approverId,
      approvedByName: approverName,
      approvedAt: new Date().toISOString(),

      approved_at: null,

      rejectedBy: null, rejectedByName: null, rejectedAt: null
    });

    console.log(`(Master) User ${userName} (${uid}) APPROVED by ${approverName}.`);

    res.status(200).send({
      message: `User ${userName} has been approved successfully.`,
      approvedBy: approverName
    });

  } catch (error) {
    console.error('(Master) Error approving user:', error);
    res.status(500).send({ error: 'Failed to approve user', details: error.message });
  }
});
// =======================================================
// == API 1.4: Master Pending Users (FIXED: Date Crash & Added Security)
// =======================================================
app.get('/api/master/pending-users', verifyToken, async (req, res) => {
  try {
    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('status', '==', 'PENDING').get();

    if (snapshot.empty) {
      return res.status(200).json({
        count: 0,
        users: [],
        message: 'No pending users found.'
      });
    }
    const formatOptions = {
      hour: 'numeric', minute: 'numeric', hour12: true,
      day: '2-digit', month: '2-digit', year: '2-digit',
      timeZone: 'Asia/Kolkata'
    };

    const safeFormat = (val) => {
      if (!val) return null;

      if (typeof val.toDate === 'function') {
        return val.toDate().toLocaleString('en-IN', formatOptions);
      }

      const dateObj = new Date(val);
      if (!isNaN(dateObj.getTime())) {
        return dateObj.toLocaleString('en-IN', formatOptions);
      }

      return val;
    };

    const pendingUsers = [];
    snapshot.forEach(doc => {
      const userData = doc.data();

      delete userData.password;
      userData.createdAt = safeFormat(userData.createdAt);
      userData.submitted_at = safeFormat(userData.submitted_at);

      pendingUsers.push({
        ...userData,
        uid: doc.id
      });
    });

    console.log(`(Master) Fetched ${pendingUsers.length} pending users.`);
    res.status(200).json({
      count: pendingUsers.length,
      users: pendingUsers
    });

  } catch (error) {
    console.error('(Master) Error fetching pending users:', error);
    res.status(500).send({ error: 'Failed to fetch pending users', details: error.message });
  }
});

// =======================================================
// == API 1.7: Master REJECT User (Fixed: Audit & Cleanup)
// =======================================================
app.post('/api/master/rejectUser/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;

    const { uid: adminId, name, fullName, role } = req.user;
    const adminName = fullName || name || role || 'Master Admin';

    if (!uid) return res.status(400).send({ error: "User ID is required." });

    const userDocRef = db.collection('users').doc(uid);
    const doc = await userDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: "User not found." });

    const userData = doc.data();
    const userName = userData.fullName || "User";

    await userDocRef.update({
      status: 'REJECTED',

      rejectedBy: adminId,
      rejectedByName: adminName,
      rejectedAt: new Date().toISOString(),

      updatedBy: adminId,
      updatedByName: adminName,
      updatedAt: new Date().toISOString(),

      approvedBy: null, approvedByName: null, approvedAt: null, approved_at: null,
      suspendedBy: null, suspendedByName: null, suspendedAt: null, suspended_at: null,
      reviewed_at: null
    });

    console.log(`(Master) User ${userName} (${uid}) REJECTED by ${adminName}.`);

    res.status(200).send({
      message: `User ${userName} has been rejected.`,
      rejectedBy: adminName
    });

  } catch (error) {
    console.error('(Master) Error rejecting user:', error);
    res.status(500).send({ error: 'Failed to reject user', details: error.message });
  }
});

// =======================================================
// == API 1.8: Admin SUSPEND User (Fixed: Audit & Cleanup)
// =======================================================
app.post('/api/admin/suspendUser/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { suspensionReason } = req.body;

    const { uid: adminId, name, fullName, role } = req.user;
    const adminName = fullName || name || role || 'Admin';

    if (!uid) return res.status(400).send({ error: "User ID is required." });

    const userDocRef = db.collection('users').doc(uid);
    const doc = await userDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: "User not found." });

    const userData = doc.data();

    if (userData.status !== 'APPROVED') {
      return res.status(400).send({ error: `Cannot suspend a user with status: ${userData.status}` });
    }

    await userDocRef.update({
      status: 'SUSPENDED',
      suspensionReason: suspensionReason || 'No reason provided',

      suspendedBy: adminId,
      suspendedByName: adminName,
      suspendedAt: new Date().toISOString(),

      updatedBy: adminId,
      updatedByName: adminName,
      updatedAt: new Date().toISOString(),

      approvedBy: null, approvedByName: null, approvedAt: null, approved_at: null,
      rejectedBy: null, rejectedByName: null, rejectedAt: null,
      suspended_at: null
    });

    console.log(`(Admin) User ${userData.fullName} (${uid}) SUSPENDED by ${adminName}.`);

    res.status(200).send({
      message: `User ${userData.fullName} has been suspended.`,
      suspendedBy: adminName
    });

  } catch (error) {
    console.error('(Admin) Error suspending user:', error);
    res.status(500).send({ error: 'Failed to suspend user', details: error.message });
  }
});

app.get('/api/admin/users', verifyToken, async (req, res) => {
  try {
    const { status: filterStatus, division, zone, depot } = req.query;
    const { uid: requesterUid, role, zone: userZone, division: userDivision } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    let query = db.collection('users');

    if (userRole === 'company master' || userRole === 'super admin' || userRole === 'admin') {
      if (zone) query = query.where('zone', '==', zone);
      if (division) query = query.where('division', '==', division);
    }
    else if (userRole === 'railway master') {
      query = query.where('zone', '==', userZone);
      if (division) query = query.where('division', '==', division);
    }
    else {
      query = query.where('division', '==', userDivision);
    }

    const snapshot = await query.get();

    let userList = [];
    let stats = { pending: 0, approved: 0, rejected: 0 };

    snapshot.forEach(doc => {
      const d = doc.data();
      const targetRole = (d.role || '').toLowerCase();

      if (doc.id === requesterUid) return;

      if (userRole.includes('supervisor') && !userRole.includes('admin')) {
        const highLevelRoles = ['admin', 'super admin', 'company master', 'railway master'];
        if (highLevelRoles.some(r => targetRole.includes(r))) return;
      }

      const s = (d.status || '').toUpperCase();
      if (s === 'PENDING') stats.pending++;
      if (s === 'APPROVED') stats.approved++;
      if (s === 'REJECTED') stats.rejected++;

      if (filterStatus) {
        if (s === filterStatus.toUpperCase()) {
          userList.push({ ...d, uid: doc.id });
        }
      } else {
        userList.push({ ...d, uid: doc.id });
      }
    });

    res.status(200).json({
      success: true,
      count: userList.length,
      stats: stats,
      users: userList
    });

  } catch (error) {
    res.status(500).send({ error: error.message });
  }
});

// =======================================================
// == API: Get All Railway Workers (Filtered by Role & Region)
// =======================================================
app.get('/api/admin/railway-workers', verifyToken, async (req, res) => {
  try {
    const { status: filterStatus, division, zone } = req.query;
    const { role: requesterRole, zone: requesterZone, division: requesterDivision } = req.user;

    const userRole = (requesterRole || '').trim().toLowerCase();
    console.log(`[GET /api/admin/railway-workers] userRole: ${userRole}, requesterDivision: ${requesterDivision}`);

    let query = db.collection('users').where('role', '==', 'Railway Worker');

    if (userRole === 'company master' || userRole === 'super admin' || userRole === 'admin') {
      if (zone) query = query.where('zone', '==', zone);
      if (division) query = query.where('division', '==', division);
    }
    else if (userRole === 'railway master' || userRole === 'contractor master') {
      if (requesterZone) query = query.where('zone', '==', requesterZone);
      if (division) query = query.where('division', '==', division);
    }
    else {
      if (requesterDivision) {
        query = query.where('division', '==', requesterDivision);
      } else if (requesterZone) {
        query = query.where('zone', '==', requesterZone);
      }
    }

    let snapshot = await query.get();
    console.log(`[GET /api/admin/railway-workers] Firestore query returned ${snapshot.size} users`);

    // Fallback: If division/local query returned 0 results, fall back to checking the entire zone
    if (snapshot.empty && userRole !== 'company master' && userRole !== 'super admin' && userRole !== 'admin') {
      if (requesterZone) {
        console.log(`[GET /api/admin/railway-workers] Initial query was empty. Trying zone fallback: ${requesterZone}`);
        const fallbackQuery = db.collection('users')
          .where('role', '==', 'Railway Worker')
          .where('zone', '==', requesterZone);
        snapshot = await fallbackQuery.get();
        console.log(`[GET /api/admin/railway-workers] Zone fallback query returned ${snapshot.size} users`);
      }
    }

    let workerList = [];
    let stats = { pending: 0, approved: 0, rejected: 0 };

    snapshot.forEach(doc => {
      const d = doc.data();
      const s = (d.status || '').toUpperCase();

      if (s === 'PENDING') stats.pending++;
      if (s === 'APPROVED') stats.approved++;
      if (s === 'REJECTED') stats.rejected++;

      if (filterStatus) {
        if (s === filterStatus.toUpperCase()) {
          workerList.push({ ...d, uid: doc.id });
        }
      } else {
        workerList.push({ ...d, uid: doc.id });
      }
    });

    res.status(200).json({
      success: true,
      count: workerList.length,
      stats: stats,
      workers: workerList
    });

  } catch (error) {
    console.error('(Workers) Error fetching railway workers:', error);
    res.status(500).send({
      success: false,
      error: error.message
    });
  }
});

// =======================================================
// == API: Get Worker Profile & Assigned Train Runs
// =======================================================
app.get('/api/worker/profile', verifyToken, async (req, res) => {
  try {
    const { uid } = req.user; // Token se worker ki UID mil jayegi

    // 1. Fetch User Profile Data from 'users' collection
    const userDoc = await db.collection('users').doc(uid).get();

    if (!userDoc.exists) {
      return res.status(404).send({ success: false, error: 'Worker profile not found.' });
    }

    const userData = userDoc.data();

    // 2. Fetch Assigned Runs from 'RunInstance' collection
    // Logic: Humein wo saare documents chahiye jahan 'coaches' array ke andar workerId matches
    const runsSnapshot = await db.collection('RunInstance')
      .where('coaches', 'array-contains-any', [
        { workerId: uid }, // Ye query tabhi work karegi agar array objects exactly match karein
      ])
      // NOTE: Firestore mein object search thoda tricky hota hai, 
      // isliye safe side ke liye hum saare 'Scheduled' aur 'Active' runs nikal kar filter karenge.
      .get();

    // Kyunki Firestore 'array-contains' objects ke partial fields par kaam nahi karta, 
    // isliye hum better approach use karenge:

    const allRunsSnapshot = await db.collection('RunInstance')
      .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running'])
      .get();

    let assignedRuns = [];

    allRunsSnapshot.forEach(doc => {
      const runData = doc.data();
      // Check if this worker is in the coaches array of this run
      const assignedCoach = runData.coaches.find(c => c.workerId === uid);

      if (assignedCoach) {
        assignedRuns.push({
          runInstanceId: runData.runInstanceId,
          instanceId: runData.instanceId,
          trainNo: runData.trainNo,
          trainName: runData.trainName,
          departureDate: runData.departureDate,
          outboundTrainNo: runData.outboundTrainNo,
          inboundTrainNo: runData.inboundTrainNo,
          status: runData.status,
          // Coach specific details for the worker
          myCoach: {
            coachPosition: assignedCoach.coachPosition,
            coachType: assignedCoach.coachType,
            attendanceStatus: assignedCoach.attendanceStatus
          }
        });
      }
    });

    // 3. Send Merged Response
    res.status(200).json({
      success: true,
      profile: {
        fullName: userData.fullName,
        email: userData.email,
        mobile: userData.mobile,
        designation: userData.designation,
        division: userData.division,
        zone: userData.zone,
        status: userData.status,
        uid: userData.uid,
        userType: userData.userType
      },
      assignedRuns: assignedRuns.sort((a, b) => new Date(b.departureDate) - new Date(a.departureDate)) // Latest runs first
    });

  } catch (error) {
    console.error('(Worker Profile) Error:', error);
    res.status(500).send({
      success: false,
      error: 'Failed to fetch profile',
      details: error.message
    });
  }
});

// =======================================================
// == API: Worker Statistics
// =======================================================
app.get('/api/worker/statistics', verifyToken, async (req, res) => {
  try {
    const { uid } = req.user;
    const now = new Date();

    // 1. Find active run instance for this worker
    const runsSnapshot = await db.collection('RunInstance')
      .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running'])
      .get();

    let activeRun = null;
    runsSnapshot.forEach(doc => {
      const runData = doc.data();
      const assignedCoach = runData.coaches.find(c => c.workerId === uid);
      if (assignedCoach && !activeRun) {
        activeRun = { ...runData, runInstanceId: runData.runInstanceId || doc.id, coachType: assignedCoach.coachType };
      }
    });

    // 2. Task counts — replicate board counting logic
    let dueTasks = 0, overdueTasks = 0, completedTasks = 0, upcomingTasks = 0;

    if (activeRun) {
      const runInstanceId = activeRun.runInstanceId;
      const coachType = activeRun.coachType || 'S2';

      // Fetch completed tasks from obhs_tasks
      const completedSnapshot = await db.collection('obhs_tasks')
        .where('runInstanceId', '==', runInstanceId)
        .get();
      const completedTaskIds = new Set();
      completedSnapshot.forEach(doc => { completedTaskIds.add(doc.data().uid); });

      // Build task definitions same as board endpoint
      const tripStartDate = new Date(activeRun.createdAt || now);
      const toiletFreqs = [
        { freq: 'Hour 1', startRel: 0, duration: 1 },
        { freq: 'Hour 2', startRel: 1, duration: 1 },
        { freq: 'Hour 3', startRel: 2, duration: 1 },
        { freq: 'Hour 5', startRel: 4, duration: 2 },
        { freq: 'Hour 7', startRel: 6, duration: 2 },
        { freq: 'Hour 9', startRel: 8, duration: 2 },
      ];
      const coachCleaningFreqs = [
        { freq: 'Train Start', startRel: 0, duration: 3 },
        { freq: 'Mid Journey', startRel: 6, duration: 4 },
        { freq: 'Train End', startRel: 12, duration: 3 },
      ];
      const linenFreqs = [
        { freq: 'Initial Distribution', startRel: 0, duration: 4 },
        { freq: 'Night Return Check', startRel: 10, duration: 4 },
      ];

      const taskDefs = [];
      for (const coach of [coachType]) {
        for (const f of toiletFreqs) taskDefs.push({ type: 'Toilet Cleaning', coach, ...f });
        for (const f of coachCleaningFreqs) taskDefs.push({ type: 'Coach Cleaning', coach, ...f });
        for (const f of linenFreqs) taskDefs.push({ type: 'Linen Distribution', coach, ...f });
      }

      taskDefs.forEach(task => {
        const startTime = new Date(tripStartDate.getTime() + task.startRel * 60 * 60 * 1000);
        const endTime = new Date(startTime.getTime() + task.duration * 60 * 60 * 1000);
        const suffix = task.freq ? `_${task.freq.replace(/\s+/g, '')}` : '';
        const generatedTaskId = `${runInstanceId}_${task.coach}_${task.type.replace(/\s+/g, '')}${suffix}`;

        if (completedTaskIds.has(generatedTaskId)) {
          completedTasks++;
        } else if (now > endTime) {
          overdueTasks++;
        } else if (now >= startTime) {
          dueTasks++;
        } else {
          upcomingTasks++;
        }
      });
    }

    // 3. Attendance counts for today
    const today = now.toISOString().split('T')[0];
    let attendancePercentage = 0;
    try {
      const attendanceSnapshot = await db.collection('obhs_attendance')
        .where('workerId', '==', uid)
        .get();
      let markedCount = 0;
      attendanceSnapshot.forEach(doc => {
        const d = doc.data();
        const docDate = d.createdAt ? d.createdAt.split('T')[0] : '';
        if (docDate !== today) return;
        if (d.isStartMarked) markedCount++;
        if (d.isMidMarked) markedCount++;
        if (d.isEndMarked) markedCount++;
      });
      const expected = 3;
      attendancePercentage = Math.round((markedCount / expected) * 100);
    } catch (_) { /* ignore */ }

    // 4. Complaints count for this worker
    let complaintsRaised = 0;
    try {
      const complaintsSnapshot = await db.collection('obhs_complaints')
        .where('submittedBy.uid', '==', uid)
        .get();
      complaintsRaised = complaintsSnapshot.size;
    } catch (_) { /* ignore */ }

    // 5. Average rating
    let averageRating = 0;
    try {
      const ratingSnapshot = await db.collection('ratings')
        .where('workerId', '==', uid)
        .get();
      let total = 0, count = 0;
      ratingSnapshot.forEach(doc => {
        const r = doc.data().rating || 0;
        total += r;
        count++;
      });
      averageRating = count > 0 ? Math.round((total / count) * 10) / 10 : 0;
    } catch (_) { /* ignore */ }

    res.status(200).json({
      success: true,
      statistics: {
        dueTasks,
        overdueTasks,
        completedTasks,
        upcomingTasks,
        tasksCompleted: completedTasks,
        attendancePercentage,
        complaintsRaised,
        averageRating
      }
    });

  } catch (error) {
    console.error('(Worker Statistics) Error:', error);
    res.status(500).json({ success: false, error: 'Failed to fetch statistics', details: error.message });
  }
});

// =======================================================
// == FORGOT PASSWORD WORKFLOW (Mobile)
// =======================================================

app.post('/api/auth/forgot-password/send-otp', async (req, res) => {
  try {
    const { mobile } = req.body;
    if (!mobile) return res.status(400).send({ error: "Mobile number is required." });

    if (!/^\d{10}$/.test(mobile)) {
      return res.status(400).send({ error: "Invalid mobile number. Must be 10 digits." });
    }

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('mobile', '==', mobile).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).send({ error: "Mobile number not registered." });
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    otpStore.set(`RESET_${mobile}`, otp);

    if (!process.env.TWILIO_PHONE_NUMBER) {
      console.error('❌ TWILIO_PHONE_NUMBER not configured');
      return res.status(500).send({ error: "SMS service not configured" });
    }

    const formattedPhone = `+91${mobile}`;
    await client.messages.create({
      body: `Your Password Reset OTP is: ${otp}. Do not share this with anyone.`,
      from: process.env.TWILIO_PHONE_NUMBER,
      to: formattedPhone
    });

    console.log(`(ForgotPwd) OTP ${otp} sent to ${mobile}`);
    res.status(200).json({ message: "OTP sent to your registered mobile number." });

  } catch (error) {
    console.error('(ForgotPwd) Error sending OTP:', error);
    if (error.code === 21211) {
      return res.status(400).send({ error: "Invalid phone number format." });
    }
    res.status(500).send({ error: 'Failed to send OTP', details: error.message });
  }
});

// --- Step 2: Verify OTP & Get Reset Token ---
app.post('/api/auth/forgot-password/verify-otp', async (req, res) => {
  try {
    const { mobile, otp } = req.body;

    const storedOtp = await otpStore.get(`RESET_${mobile}`);

    if (!storedOtp) return res.status(400).json({ error: "OTP expired or invalid request." });
    if (storedOtp !== otp) return res.status(400).json({ error: "Invalid OTP." });

    await otpStore.delete(`RESET_${mobile}`);

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('mobile', '==', mobile).limit(1).get();

    if (snapshot.empty) return res.status(404).send({ error: "User not found." });
    const userDoc = snapshot.docs[0];
    const resetToken = jwt.sign(
      { uid: userDoc.id, purpose: 'password_reset' },
      process.env.JWT_SECRET,
      { expiresIn: '10m' }
    );

    console.log(`(ForgotPwd) OTP Verified for ${mobile}. Sending Reset Token.`);

    res.status(200).json({
      message: "OTP Verified. Please proceed to reset password.",
      resetToken: resetToken
    });

  } catch (error) {
    console.error('(ForgotPwd) Error verifying OTP:', error);
    res.status(500).send({ error: 'Failed to verify OTP', details: error.message });
  }
});

app.post('/api/auth/forgot-password/reset', async (req, res) => {
  try {
    const { newPassword, resetToken } = req.body;

    if (!newPassword || !resetToken) {
      return res.status(400).send({ error: "New Password and Reset Token are required." });
    }

    let decoded;
    try {
      decoded = jwt.verify(resetToken, process.env.JWT_SECRET);
    } catch (err) {
      return res.status(403).send({ error: "Invalid or expired reset session. Please try again." });
    }

    if (decoded.purpose !== 'password_reset') {
      return res.status(403).send({ error: "Invalid token type." });
    }

    const uid = decoded.uid;
    await db.collection('users').doc(uid).update({
      password: newPassword,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(ForgotPwd) Password reset success for User UID: ${uid}`);
    res.status(200).send({ message: "Password has been reset successfully. You can now login." });

  } catch (error) {
    console.error('(ForgotPwd) Error resetting password:', error);
    res.status(500).send({ error: 'Failed to reset password', details: error.message });
  }
});

// =======================================================
// == FORGOT PASSWORD WORKFLOW (EMAIL)
// =======================================================

app.post('/api/auth/forgot-password/email/send-otp', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) return res.status(400).send({ error: "Email is required." });
    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('email', '==', email).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).send({ error: "Email address not registered." });
    }
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    otpStore.set(`RESET_EMAIL_${email}`, otp);

    const { data, error } = await resend.emails.send({
      from: 'Swachh Railways <auth@swachhrailways.com>',
      to: [email],
      subject: 'Password Reset OTP - Swachh Railways',
      html: `
        <div style="font-family: Arial, sans-serif; border: 1px solid #ddd; padding: 20px; border-radius: 10px;">
          <h2 style="color: #2c3e50;">Password Reset Request</h2>
          <p>Your One-Time Password (OTP) for password reset is:</p>
          <h1 style="color: #e67e22; letter-spacing: 5px;">${otp}</h1>
          <p>This code is valid for <b>5 minutes</b>. Please do not share this with anyone.</p>
          <hr style="border: 0; border-top: 1px solid #eee;" />
          <p style="font-size: 12px; color: #7f8c8d;">If you didn't request this, please ignore this email.</p>
        </div>
      `,
    });

    if (error) {
      console.error('(Resend) ERROR:', error);
      return res.status(400).json({ error: "Failed to send email via Resend", details: error });
    }

    console.log(`(ForgotPwd) Email OTP ${otp} sent to ${email}`);
    res.status(200).json({ message: "OTP sent to your registered email." });

  } catch (error) {
    console.error('(ForgotPwd) Error sending Email OTP:', error);
    res.status(500).send({ error: 'Failed to send OTP email', details: error.message });
  }
});

app.post('/api/auth/forgot-password/email/verify-otp', async (req, res) => {
  try {
    const { email, otp } = req.body;

    const storedOtp = await otpStore.get(`RESET_EMAIL_${email}`);

    if (!storedOtp) return res.status(400).json({ error: "OTP expired or invalid request." });
    if (storedOtp !== otp) return res.status(400).json({ error: "Invalid OTP." });

    await otpStore.delete(`RESET_EMAIL_${email}`);

    const usersRef = db.collection('users');
    const snapshot = await usersRef.where('email', '==', email).limit(1).get();

    if (snapshot.empty) return res.status(404).send({ error: "User not found." });
    const userDoc = snapshot.docs[0];

    const resetToken = jwt.sign(
      { uid: userDoc.id, purpose: 'password_reset' },
      process.env.JWT_SECRET,
      { expiresIn: '10m' }
    );

    console.log(`(ForgotPwd) Email OTP Verified for ${email}. Sending Token.`);

    res.status(200).json({
      message: "OTP Verified. Please proceed to reset password.",
      resetToken: resetToken
    });

  } catch (error) {
    console.error('(ForgotPwd) Error verifying Email OTP:', error);
    res.status(500).send({ error: 'Failed to verify OTP', details: error.message });
  }
});

// GET /api/users/workers
app.get('/api/users/workers', verifyToken, async (req, res) => {
  try {
    const snapshot = await db.collection('users')
      .where('role', 'in', ['Worker', 'Railway Worker', 'janitor', 'Janitor'])
      // .where('status', '==', 'APPROVED') // We might want only approved
      .get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, workers: [] });
    }

    const workersList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      workersList.push({
        uid: data.uid || doc.id,
        fullName: data.fullName || '',
        email: data.email || '',
        mobile: data.mobile || '',
        role: data.role || '',
        designation: data.designation || '',
        status: data.status || 'PENDING'
      });
    });

    res.status(200).json({ count: workersList.length, workers: workersList });
  } catch (error) {
    console.error('(GetWorkers) Error:', error);
    res.status(500).send({ error: 'Failed to fetch workers', details: error.message });
  }
});
app.get('/api/users/railway-supervisors', verifyToken, async (req, res) => {
  try {
    const { zone, division, role } = req.user;

    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');

    if (!zone) {
      return res.status(400).send({ error: "Your user profile is missing Zone." });
    }

    let query = db.collection('users')
      .where('role', '==', 'Railway Supervisor')
      .where('status', '==', 'APPROVED');

    if (isMaster) {
      query = query.where('zone', '==', zone);
      console.log(`(GetSupervisors) Master access: Fetching entire Zone ${zone}`);
    }
    else {
      if (!division) {
        return res.status(400).send({ error: "Your user profile is missing Division." });
      }
      query = query.where('zone', '==', zone)
        .where('division', '==', division);
      console.log(`(GetSupervisors) Admin/Sup access: Fetching Division ${division}`);
    }

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, supervisors: [] });
    }

    const supervisorList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      supervisorList.push({
        uid: data.uid,
        fullName: data.fullName,
        division: data.division,
        depot: data.depot || ""
      });
    });

    res.status(200).json({
      count: supervisorList.length,
      supervisors: supervisorList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: `Firebase Console check karo. ${error.message}`
      });
    }
    console.error('(User) Error fetching supervisors:', error);
    res.status(500).send({ error: 'Failed to fetch supervisors', details: error.message });
  }
});


app.get('/api/users/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const requester = req.user;

    if (!uid) return res.status(400).send({ error: "User ID (UID) is required." });

    const docRef = db.collection('users').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "User not found." });

    const targetUser = doc.data();

    if (requester.userType === 'railway') {
      const reqRole = (requester.role || '').toLowerCase();

      if (!reqRole.includes('master')) {
        if (targetUser.division !== requester.division) {
          return res.status(403).send({ error: "Access Denied: You cannot view users of another Division." });
        }
      }
      else {
        if (targetUser.zone !== requester.zone) {
          return res.status(403).send({ error: "Access Denied: You cannot view users of another Zone." });
        }
      }
    }
    else if (requester.userType === 'contractor') {
      if (targetUser.entityId !== requester.entityId) {
        return res.status(403).send({ error: "Access Denied: Different Company." });
      }
    }
    delete targetUser.password;
    res.status(200).json(targetUser);

  } catch (error) {
    console.error('(User) Error fetching single user:', error);
    res.status(500).send({ error: 'Failed to fetch user', details: error.message });
  }
});

// =======================================================
// == 2. ENTITY (COMPANY) REGISTRATION APIs
// =======================================================

// =======================================================
// == API 2.1: Create a new Entity/Company
// =======================================================
app.post('/api/contractors', verifyToken, async (req, res) => {
  try {
    const {
      companyName,
      registrationType, panNumber, gstinNumber,
      registeredAddress, contactNumber, email,
      alternateContact, website, yearOfEstablishment, gemId,
      // zone,
      // division
    } = req.body;

    const { uid, name, fullName, email: userEmail, role } = req.user;
    const creatorName = fullName || name || userEmail || role || 'Unknown';

    if (!companyName || !registrationType) {
      return res.status(400).send({
        error: "companyName and registrationType are required."
      });
    }

    const existing = await db.collection('entities').where('companyName', '==', companyName).get();
    if (!existing.empty) {
      return res.status(400).send({ error: "Company with this name already exists." });
    }

    const docRef = db.collection('entities').doc();

    const newEntity = {
      uid: docRef.id,
      companyName,
      registrationType,
      // zone: zone || null,
      // division: division || null,

      panNumber: panNumber || null,
      gstinNumber: gstinNumber || null,
      registeredAddress: registeredAddress || null,
      contactNumber: contactNumber || null,
      email: email || null,
      alternateContact: alternateContact || null,
      website: website || null,
      yearOfEstablishment: yearOfEstablishment || null,
      gemId: gemId || null,

      status: 'PENDING',
      createdBy: uid,
      createdByName: creatorName,
      createdAt: new Date().toISOString(),
      updatedBy: null,
      updatedByName: null,
      updatedAt: null,
      approvedBy: null,
      approvedByName: null,
      approvedAt: null,
      rejectedBy: null,
      rejectedByName: null,
      rejectedAt: null
    };

    await docRef.set(newEntity);

    console.log(`(Entity) Created: ${companyName} by ${creatorName}`);
    res.status(201).send({
      message: 'Entity created successfully and is pending approval.',
      uid: docRef.id,
      entity: newEntity
    });

  } catch (error) {
    console.error('(Entity) Error creating entity:', error);
    res.status(500).send({ error: 'Failed to create entity', details: error.message });
  }
});
// =======================================================
// == UPDATE ENTITY
// =======================================================
app.put('/api/contractors/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const updates = req.body;

    const { uid: userId, name, fullName, email, role } = req.user;
    const editorName = fullName || name || email || role || 'Unknown';

    const docRef = db.collection('entities').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Entity not found." });

    const currentData = doc.data();

    updates.updatedBy = userId;
    updates.updatedByName = editorName;
    updates.updatedAt = new Date().toISOString();

    // RE-APPROVAL LOGIC: Reset to PENDING on edit
    if (!updates.status || (updates.status !== 'APPROVED' && updates.status !== 'REJECTED')) {
      updates.status = 'PENDING';
    }

    if (updates.status) {

      if (updates.status === 'APPROVED' && currentData.status !== 'APPROVED') {
        updates.approvedBy = userId;
        updates.approvedByName = editorName;
        updates.approvedAt = new Date().toISOString();

        updates.rejectedBy = null; updates.rejectedByName = null; updates.rejectedAt = null;
        updates.suspendedBy = null; updates.suspendedByName = null; updates.suspendedAt = null;
        updates.approved_at = null;

        console.log(`(Entity) Approved: ${currentData.companyName} by ${editorName}`);
      }

      else if (updates.status === 'REJECTED') {
        updates.rejectedBy = userId;
        updates.rejectedByName = editorName;
        updates.rejectedAt = new Date().toISOString();

        updates.approvedBy = null; updates.approvedByName = null; updates.approvedAt = null;
        updates.approved_at = null;
      }

      else if (updates.status === 'SUSPENDED') {
        console.log(`(Entity) Suspending Entity: ${currentData.companyName}`);

        updates.suspendedBy = userId;
        updates.suspendedByName = editorName;
        updates.suspendedAt = new Date().toISOString();

        const usersSnapshot = await db.collection('users').where('entityId', '==', uid).get();
        if (!usersSnapshot.empty) {
          const batch = db.batch();
          usersSnapshot.forEach(doc => batch.update(doc.ref, { status: 'SUSPENDED' }));
          await batch.commit();
        }

        const contractsSnapshot = await db.collection('contracts').where('entityId', '==', uid).get();
        if (!contractsSnapshot.empty) {
          const batch2 = db.batch();
          contractsSnapshot.forEach(doc => batch2.update(doc.ref, { status: 'SUSPENDED' }));
          await batch2.commit();
        }
      }
    }

    await docRef.update(updates);

    console.log(`(Entity) Updated: ${uid} | Status: ${updates.status || 'Edited'} | By: ${editorName}`);
    res.status(200).send({
      message: 'Entity updated successfully',
      updates: updates
    });

  } catch (error) {
    console.error('(Entity) Error updating:', error);
    res.status(500).send({ error: 'Failed to update entity', details: error.message });
  }
});

// =======================================================
// == API 2.2: Get ENTITIES 
// =======================================================
app.get('/api/contractors', verifyToken, async (req, res) => {
  try {
    const queryStatus = req.query.status;
    const { userType, entityId, division, zone, role } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    console.log(`(Entity) User: ${userType} | Role: "${role}" | Zone: "${zone}"`);

    let finalEntityIds = [];
    let isFilteredByContract = false;

    if (userType === 'railway') {
      let contractQuery = db.collection('contracts');

      if (userRole === 'railway master') {
        if (!zone) return res.status(403).send({ error: "Zone missing in profile." });
        contractQuery = contractQuery.where('zone', '==', zone);
        isFilteredByContract = true;
      }
      else if (userRole.includes('admin') || userRole.includes('supervisor')) {
        if (!division) return res.status(403).send({ error: "Division missing in profile." });
        contractQuery = contractQuery.where('division', '==', division);
        isFilteredByContract = true;
      }

      if (isFilteredByContract) {
        const contractSnap = await contractQuery.get();
        if (!contractSnap.empty) {
          finalEntityIds = [...new Set(contractSnap.docs.map(doc => doc.data().entityId))];
        }
        // When no contracts exist yet, skip entityId filter so all approved entities show up
      }
    }

    let entityQuery = db.collection('entities');

    if (userType === 'contractor') {
      if (!entityId) return res.status(403).send({ error: "Entity ID missing." });
      entityQuery = entityQuery.where('uid', '==', entityId);
    }
    else if (isFilteredByContract && finalEntityIds.length > 0) {
      entityQuery = entityQuery.where('uid', 'in', finalEntityIds.slice(0, 30));
    }

    if (queryStatus) {
      entityQuery = entityQuery.where('status', '==', queryStatus);
    } else if (userType === 'railway') {
      entityQuery = entityQuery.where('status', '==', 'APPROVED');
    }

    const snapshot = await entityQuery.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, contractors: [] });
    }

    const formatOptions = {
      hour: 'numeric', minute: 'numeric', hour12: true,
      day: '2-digit', month: '2-digit', year: '2-digit',
      timeZone: 'Asia/Kolkata'
    };

    const safeFormat = (val) => {
      if (!val) return null;
      let dateObj = (typeof val.toDate === 'function') ? val.toDate() : new Date(val);
      return isNaN(dateObj.getTime()) ? val : dateObj.toLocaleString('en-IN', formatOptions);
    };

    const contractorList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      contractorList.push({
        ...data,
        uid: doc.id,
        createdAt: safeFormat(data.createdAt),
        submitted_at: safeFormat(data.submitted_at),
        approved_at: safeFormat(data.approvedAt || data.approved_at),
        updatedAt: safeFormat(data.updatedAt)
      });
    });

    res.status(200).json({ count: contractorList.length, contractors: contractorList });

  } catch (error) {
    console.error('(Entity) Error fetching entities:', error);
    res.status(500).send({ error: 'Failed to fetch entities', details: error.message });
  }
});
// =======================================================
// == API: Master Approve Entity
// =======================================================
app.post('/api/master/approveContractor/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;

    const { uid: adminId, name, fullName, role } = req.user;
    const adminName = fullName || name || role || 'Master Admin';

    if (!uid) return res.status(400).send({ error: "Entity ID is required." });

    const docRef = db.collection('entities').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Entity not found." });

    const entityName = doc.data().companyName;

    await docRef.update({
      status: 'APPROVED',

      approvedBy: adminId,
      approvedByName: adminName,
      approvedAt: new Date().toISOString(),

      updatedBy: adminId,
      updatedByName: adminName,
      updatedAt: new Date().toISOString(),

      rejectedBy: null,
      rejectedByName: null,
      rejectedAt: null,

      suspendedBy: null,
      suspendedByName: null,
      suspendedAt: null,

      approved_at: null
    });

    console.log(`(Master) Entity ${entityName} (${uid}) APPROVED by ${adminName}.`);

    res.status(200).send({
      message: `Entity ${entityName} has been approved.`,
      approvedBy: adminName
    });

  } catch (error) {
    console.error('(Master) Error approving entity:', error);
    res.status(500).send({ error: 'Failed to approve entity', details: error.message });
  }
});

// =======================================================
// == API 2.4: Master Rejects an Entity 
// =======================================================
app.post('/api/master/rejectContractor/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;

    const { uid: adminId, name, fullName, role } = req.user;
    const adminName = fullName || name || role || 'Master Admin';

    if (!uid) return res.status(400).send({ error: "Entity ID is required." });

    const docRef = db.collection('entities').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Entity not found." });

    const entityName = doc.data().companyName;

    await docRef.update({
      status: 'REJECTED',

      rejectedBy: adminId,
      rejectedByName: adminName,
      rejectedAt: new Date().toISOString(),
      updatedBy: adminId,
      updatedByName: adminName,
      updatedAt: new Date().toISOString(),

      approvedBy: null, approvedByName: null, approvedAt: null, approved_at: null,
      suspendedBy: null, suspendedAt: null,

      reviewed_at: null
    });

    console.log(`(Master) Entity ${entityName} (${uid}) REJECTED by ${adminName}.`);

    res.status(200).send({
      message: `Entity ${entityName} has been rejected.`,
      rejectedBy: adminName
    });

  } catch (error) {
    console.error('(Master) Error rejecting entity:', error);
    res.status(500).send({ error: 'Failed to reject entity', details: error.message });
  }
});


// --- API 2.5: Admin Suspends an Entity ---
app.post('/api/admin/suspendContractor/:uid', async (req, res) => {
  try {
    const { uid } = req.params;
    const { suspensionReason } = req.body;

    if (!uid) return res.status(400).send({ error: "Entity ID is required." });

    const docRef = db.collection('entities').doc(uid);
    const doc = await docRef.get();
    if (!doc.exists) return res.status(404).send({ error: "Entity not found." });

    if (doc.data().status !== 'APPROVED') {
      return res.status(400).send({ error: `Cannot suspend an entity with status: ${doc.data().status}` });
    }

    await docRef.update({
      status: 'SUSPENDED',
      suspended_at: admin.firestore.FieldValue.serverTimestamp(),
      suspensionReason: suspensionReason || null
    });

    console.log(`(Admin) Entity ${uid} has been SUSPENDED.`);
    res.status(200).send({ message: `Entity ${uid} has been suspended.` });
  } catch (error) {
    console.error('(Admin) Error suspending entity:', error);
    res.status(500).send({ error: 'Failed to suspend entity', details: error.message });
  }
});

// --- API 2.6: Get complete details for ONE Entity (Info + Contracts) ---
app.get('/api/contractors/details/:uid', async (req, res) => {
  try {
    const { uid } = req.params;
    if (!uid) {
      return res.status(400).send({ error: "Entity ID (UID) is required." });
    }

    const entityRef = db.collection('entities').doc(uid);
    const contractsRef = db.collection('contracts');

    const entityDoc = await entityRef.get();
    if (!entityDoc.exists) {
      return res.status(404).send({ error: "Entity not found." });
    }
    const entityData = entityDoc.data();

    const contractsSnapshot = await contractsRef.where('entityId', '==', uid).get();

    const contractsList = [];
    if (!contractsSnapshot.empty) {
      contractsSnapshot.forEach(doc => {
        contractsList.push(doc.data());
      });
    }

    res.status(200).json({
      details: entityData,
      contracts: contractsList
    });

  } catch (error) {
    console.error('(Entity) Error fetching complete details:', error);
    res.status(500).send({ error: 'Failed to fetch details', details: error.message });
  }
});


// =======================================================
// == 3. CONTRACT MANAGEMENT APIs
// =======================================================

// --- API 3.1: Create a new Contract ---
app.post('/api/contracts', verifyToken, async (req, res) => {
  try {
    const {
      contractNumber,
      contractName,
      entityId,
      zone,
      division,
      depot,
      startDate,
      endDate,
      workCategories,
      remarks,
      status,
      repName,
      repDesignation,
      repMobile,
      repEmail,
      repIdProofType,
      repIdProofNumber
    } = req.body;

    const { uid, name, fullName, email, role } = req.user;
    const creatorName = fullName || name || email || role || 'Unknown';

    if (!entityId) {
      return res.status(400).send({ error: "Please select a Contractor (Entity) to create a contract." });
    }

    if (!contractNumber || !contractName || !zone || !startDate || !endDate || !workCategories || !repName || !repMobile || !repEmail) {
      return res.status(400).send({ error: "Please fill all other mandatory fields (*) like Contract No, Name, Dates, etc." });
    }

    let duplicateQuery = db.collection('contracts')
      .where('entityId', '==', entityId)
      .where('zone', '==', zone)
      .where('division', '==', division || null);

    if (depot) {
      duplicateQuery = duplicateQuery.where('depot', '==', depot);
    }

    const duplicateSnap = await duplicateQuery.get();

    if (!duplicateSnap.empty) {
      const locationName = depot ? `Depot: ${depot}` : `Division: ${division}`;
      return res.status(400).send({
        error: `Restriction: This Contractor already has a contract in this ${locationName}. Same company cannot have multiple contracts in the same division/depot.`
      });
    }

    const entityDoc = await db.collection('entities').doc(entityId).get();
    if (!entityDoc.exists) {
      return res.status(404).send({ error: "Selected Contractor does not exist." });
    }

    const entityData = entityDoc.data();
    if (entityData.status !== 'APPROVED') {
      return res.status(400).send({ error: "Selected Contractor is not Active/Approved." });
    }

    const entityName = entityData.companyName;

    const representative = {
      name: repName,
      designation: repDesignation || null,
      mobile: repMobile,
      email: repEmail,
      idProofType: repIdProofType || null,
      idProofNumber: repIdProofNumber || null
    };

    const docRef = db.collection('contracts').doc();

    await docRef.set({
      uid: docRef.id,
      contractNumber,
      contractName,
      entityId,
      entityName: entityName,
      zone,
      division: division || null,
      depot: depot || null,
      startDate,
      endDate,
      workCategories,
      remarks: remarks || null,
      status: status || 'active',
      representative: representative,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      createdBy: uid,
      createdByName: creatorName
    });

    res.status(201).send({
      message: 'Contract created successfully',
      uid: docRef.id
    });

  } catch (error) {
    console.error('(Contract) Error creating contract:', error);
    res.status(500).send({ error: 'Failed to create contract', details: error.message });
  }
});
// =======================================================
// == UPDATE ENTITY
// =======================================================
app.put('/api/contractors/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const updates = req.body;

    const { uid: userId, name, fullName, email, role } = req.user;

    const editorName = fullName || name || email || role || 'Unknown';

    const docRef = db.collection('entities').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Entity not found." });

    const currentData = doc.data();

    console.log(`(Entity) Update Request for ${uid} | New Status: ${updates.status}`);

    updates.updatedBy = userId;
    updates.updatedByName = editorName;
    updates.updatedAt = new Date().toISOString();

    if (updates.status) {

      if (updates.status === 'APPROVED') {

        if (currentData.status !== 'APPROVED') {
          updates.approvedBy = userId;
          updates.approvedByName = editorName;
          updates.approvedAt = new Date().toISOString();
          updates.approved_at = null;

          updates.rejectedBy = null; updates.rejectedByName = null; updates.rejectedAt = null;
          updates.suspendedBy = null; updates.suspendedAt = null;

          console.log(`(Entity) Marking APPROVED by ${editorName}`);
        }
      }

      else if (updates.status === 'REJECTED') {
        updates.rejectedBy = userId;
        updates.rejectedByName = editorName;
        updates.rejectedAt = new Date().toISOString();

        updates.approvedBy = null; updates.approvedByName = null; updates.approvedAt = null;
        console.log(`(Entity)  Marking REJECTED by ${editorName}`);
      }

      else if (updates.status === 'SUSPENDED') {
        updates.suspendedBy = userId;
        updates.suspendedAt = new Date().toISOString();

        const usersSnapshot = await db.collection('users').where('entityId', '==', uid).get();
        if (!usersSnapshot.empty) {
          const batch = db.batch();
          usersSnapshot.forEach(doc => batch.update(doc.ref, { status: 'SUSPENDED' }));
          await batch.commit();
        }
        console.log(`(Entity)Marking SUSPENDED by ${editorName}`);
      }
    }

    await docRef.update(updates);

    res.status(200).send({
      message: 'Entity updated successfully',
      updates: updates
    });

  } catch (error) {
    console.error('(Entity) Error updating:', error);
    res.status(500).send({ error: 'Failed to update entity', details: error.message });
  }
});
// =======================================================
// == API 3.2: Update Existing Contract
// =======================================================
app.put('/api/contracts/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const updates = req.body;
    const { uid: userId, name, fullName, email, role } = req.user;

    const editorName = fullName || name || email || role || 'Unknown';

    const docRef = db.collection('contracts').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: "Contract not found." });
    }

    updates.updatedBy = userId;
    updates.updatedByName = editorName;
    updates.updatedAt = new Date().toISOString();

    // RE-APPROVAL LOGIC: Reset to PENDING on edit
    if (!updates.status || (updates.status !== 'APPROVED' && updates.status !== 'REJECTED')) {
      updates.status = 'PENDING';
    }

    await docRef.update(updates);

    console.log(`(Contract) Updated: ${uid} by ${editorName}`);

    res.status(200).send({
      message: 'Contract updated successfully',
      updates: updates
    });

  } catch (error) {
    console.error('(Contract) Error updating contract:', error);
    res.status(500).send({ error: 'Failed to update contract', details: error.message });
  }
});
// =======================================================
// == API 3.3: Get all Contracts
// =======================================================
app.get('/api/contracts', verifyToken, async (req, res) => {
  try {
    const { status, division: queryDivision, zone: queryZone, entityId } = req.query;
    const { userType, zone: userZone, division: userDivision, role } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    let query = db.collection('contracts');

    console.log(`(Contract Fetch) User: ${userRole} | Zone: ${userZone} | Div: ${userDivision}`);
    if (userType === 'railway') {

      if (userRole === 'company master' || userRole === 'super admin') {
        if (queryZone) query = query.where('zone', '==', queryZone);
        if (queryDivision) query = query.where('division', '==', queryDivision);
      }

      else if (userRole === 'railway master') {
        if (!userZone) return res.status(403).send({ error: "Zone missing in profile." });

        query = query.where('zone', '==', userZone);

        if (queryDivision) query = query.where('division', '==', queryDivision);
      }

      else if (userRole.includes('admin') || userRole.includes('supervisor')) {
        if (!userDivision) return res.status(403).send({ error: "Division missing in profile." });

        query = query.where('division', '==', userDivision);
      }
    }

    else if (userType === 'contractor') {
      const contractorEntityId = req.user.entityId;
      if (!contractorEntityId) return res.status(403).send({ error: "Entity linkage missing." });
      query = query.where('entityId', '==', contractorEntityId);
    }

    if (status) query = query.where('status', '==', status);
    if (entityId && userType !== 'contractor') query = query.where('entityId', '==', entityId);

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, contracts: [] });
    }

    const contractList = [];
    snapshot.forEach(doc => {
      contractList.push({ ...doc.data(), uid: doc.id });
    });

    res.status(200).json({
      count: contractList.length,
      contracts: contractList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Index required.',
        details: 'Firebase console mein jaakar index create karein.'
      });
    }
    console.error('(Contract) Error:', error);
    res.status(500).send({ error: 'Failed to fetch contracts', details: error.message });
  }
});

//API 3.4: Get details for a SINGLE Contract
app.get('/api/contracts/:uid', async (req, res) => {
  try {
    const { uid } = req.params;

    if (!uid) {
      return res.status(400).send({ error: "Contract ID (UID) is required." });
    }

    const docRef = db.collection('contracts').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: "Contract not found." });
    }

    res.status(200).json(doc.data());

  } catch (error) {
    console.error('(Contract) Error fetching single contract:', error);
    res.status(500).send({ error: 'Failed to fetch contract', details: error.message });
  }
});

//API 3.5: Get details for a SINGLE Contract by Contract NUMBER
app.get('/api/contracts/number/:contractNumber', async (req, res) => {
  try {
    const { contractNumber } = req.params;

    if (!contractNumber) {
      return res.status(400).send({ error: "Contract Number is required." });
    }

    const contractsRef = db.collection('contracts');
    const snapshot = await contractsRef.where('contractNumber', '==', contractNumber).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).send({ error: "Contract not found with this number." });
    }

    const contractData = snapshot.docs[0].data();
    res.status(200).json(contractData);

  } catch (error) {
    console.error('(Contract) Error fetching single contract by number:', error);
    res.status(500).send({ error: 'Failed to fetch contract', details: error.message });
  }
});

//API 3.6: Get Contracts for a specific ENTITY (Filtered by Category)
app.get('/api/contracts/by-entity/:entityId', async (req, res) => {
  try {
    const { entityId } = req.params;
    const { division, zone, status, category } = req.query;

    if (!entityId) {
      return res.status(400).send({ error: "Entity ID is required." });
    }

    let query = db.collection('contracts').where('entityId', '==', entityId);

    if (division) query = query.where('division', '==', division);
    if (zone) query = query.where('zone', '==', zone);
    if (status) query = query.where('status', '==', status);

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, contracts: [] });
    }

    const contractList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;
      if (category) {
        const cats = data.workCategories;
        if (Array.isArray(cats)) {
          const found = cats.some(c => c.toLowerCase().includes(category.toLowerCase()));
          if (!found) include = false;
        }
        else if (typeof cats === 'string') {
          if (!cats.toLowerCase().includes(category.toLowerCase())) include = false;
        }
        else {
          include = false;
        }
      }

      if (include) {
        contractList.push(data);
      }
    });
    res.status(200).json({
      count: contractList.length,
      contracts: contractList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Query requires an index.' });
    }
    console.error('(Contract) Error fetching contracts:', error);
    res.status(500).send({ error: 'Failed to fetch contracts' });
  }
});

// =======================================================
// == 4. TRAIN MANAGEMENT APIs
// =======================================================

// =======================================================
// == API 4.1: Create a new Train (Final Updated Version)
// =======================================================

const convertToDecimalDays = (timeStr) => {
  if (!timeStr || typeof timeStr !== 'string' || !timeStr.includes(':')) return 0;
  const parts = timeStr.split(':');
  const d = parseFloat(parts[0]) || 0;
  const h = parseFloat(parts[1]) || 0;
  const m = parseFloat(parts[2]) || 0;
  return d + (h / 24) + (m / 1440);
};

app.post('/api/trains', verifyToken, async (req, res) => {
  try {
    // ─── ROLE GUARD: Only Admins and Masters can create trains ───────────
    const callerRole = (req.user.role || '').toUpperCase();
    const blockedRoles = ['CONTRACTOR SUPERVISOR', 'CTS', 'RAILWAY SUPERVISOR', 'WORKER', 'RAILWAY WORKER'];
    if (blockedRoles.includes(callerRole)) {
      return res.status(403).json({
        error: 'Access Denied',
        details: `Your role (${req.user.role}) is not authorized to create trains. Only Admin-level users can manage trains.`
      });
    }
    // ─── END ROLE GUARD ──────────────────────────────────────────────────

    const allowedFields = [
      'trainNo', 'trainName', 'origin', 'destination', 'days',
      'zone', 'division', 'depot', 'status', 'TrainApplicableFor',
      'outboundTrainNo', 'inboundTrainNo', 'returnOffset', 'cycleLength',
      'outboundDurationStr', 'inboundDurationStr', 'layoverDestStr', 'layoverOriginStr',
      'journeyStartTime'
    ];

    const bodyKeys = Object.keys(req.body);
    for (const key of bodyKeys) {
      if (!allowedFields.includes(key)) {
        return res.status(400).send({
          error: `Invalid field name.`,
          details: `The field '${key}' is not allowed.`
        });
      }
    }

    const {
      trainNo, trainName, origin, destination, days,
      zone, division, depot, status, TrainApplicableFor,
      outboundTrainNo, inboundTrainNo, returnOffset, cycleLength,
      outboundDurationStr, inboundDurationStr, layoverDestStr, layoverOriginStr,
      journeyStartTime
    } = req.body;

    const isOBHSEnabled = TrainApplicableFor && TrainApplicableFor.includes('OBHS');

    let requiredInstances = 0;
    let finalCycleTime = 0;

    if (isOBHSEnabled) {
      if (!outboundTrainNo || !inboundTrainNo || !days || days.length === 0) {
        return res.status(400).send({
          error: 'OBHS Validation Failed',
          details: 'For OBHS; Outbound No, Inbound No, and Departure Days are mandatory.'
        });
      }

      // PDF Formula Implementation: Cycle Time = Outbound + LayoverDest + Inbound + LayoverOrigin
      const outbound = convertToDecimalDays(outboundDurationStr);
      const inbound = convertToDecimalDays(inboundDurationStr);
      const layoverD = convertToDecimalDays(layoverDestStr);
      const layoverO = convertToDecimalDays(layoverOriginStr);

      const calculatedC = outbound + inbound + layoverD + layoverO;
      finalCycleTime = calculatedC > 0 ? calculatedC : (Number(cycleLength) || 0);

      // Frequency (F) = Interval between departures
      const numDays = (days.includes('All Days') || days.includes('Daily')) ? 7 : days.length;
      const F = 7 / numDays;

      // Required Instances = Ceiling(Cycle / Frequency)
      requiredInstances = Math.ceil(finalCycleTime / F);
      if (requiredInstances <= 0) requiredInstances = 1;
    }

    const { uid, name, email, role } = req.user;
    const userName = name || email || role || 'Unknown';

    if (trainNo) {
      const existingTrain = await db.collection('trains').where('trainNo', '==', trainNo).get();
      if (!existingTrain.empty) {
        return res.status(400).send({ error: 'A train with this number already exists.' });
      }
    }

    const docRef = db.collection('trains').doc();

    const newTrain = {
      uid: docRef.id,
      trainNo: trainNo || null,
      trainName: trainName || null,
      origin: origin || null,
      destination: destination || null,
      days: days || null,
      zone: zone || null,
      division: division || null,
      depot: depot || null,
      status: status || 'active',
      TrainApplicableFor: TrainApplicableFor || [],
      outboundTrainNo: isOBHSEnabled ? outboundTrainNo : null,
      inboundTrainNo: isOBHSEnabled ? inboundTrainNo : null,
      cycleLength: isOBHSEnabled ? Number(finalCycleTime.toFixed(4)) : (cycleLength || null),
      requiredInstances: isOBHSEnabled ? requiredInstances : null,
      journeyStartTime: isOBHSEnabled ? (journeyStartTime || null) : null,
      outboundDurationStr: isOBHSEnabled ? (outboundDurationStr || null) : null,
      inboundDurationStr: isOBHSEnabled ? (inboundDurationStr || null) : null,
      layoverDestStr: isOBHSEnabled ? (layoverDestStr || null) : null,
      layoverOriginStr: isOBHSEnabled ? (layoverOriginStr || null) : null,
      createdBy: uid,
      createdByName: userName,
      createdAt: new Date().toISOString(),
      updatedBy: null,
      updatedByName: null,
      updatedAt: null
    };

    await docRef.set(newTrain);

    if (isOBHSEnabled && requiredInstances > 0) {
      const batch = db.batch();
      for (let i = 0; i < requiredInstances; i++) {
        const instanceLetter = String.fromCharCode(65 + i);
        const instanceId = `${trainNo}-${trainName}-Inst-${instanceLetter}`;

        batch.set(db.collection('TrainPairs').doc(instanceId), {
          instanceId: instanceId,
          instanceName: `Instance ${instanceLetter}`,
          trainNo: trainNo,
          trainName: trainName,
          status: 'Inactive',
          inboundTrainNo: inboundTrainNo,
          outboundTrainNo: outboundTrainNo,
          rotationPattern: `Round-Robin`,
          parentTrainId: docRef.id,
          createdAt: new Date().toISOString()
        });
      }
      await batch.commit();
    }

    res.status(201).send({
      message: 'Train and OBHS Pool created successfully',
      uid: docRef.id,
      calculatedInstances: requiredInstances,
      data: newTrain
    });

  } catch (error) {
    console.error('(Train) Error:', error);
    res.status(500).send({ error: 'Failed to create train', details: error.message });
  }
});

// == HELPER: Frequency-based Task Generation Engine
// =======================================================
// Generates tasks at specific times of day based on Railway Board policy:
//   06-09: hourly   09-20: every 2hr   20-22: hourly   22-06: passenger-request
// Cleaner times: 06, 07, 08, 09, 11, 13, 15, 17, 19, 20, 21
// Garbage: 07, 13, 19 (plus pre-terminal)
// Water check: 08, 16
// Safety: 10, 18
// Petty Repair: 10, 18
async function generateTaskInstancesForRun(runData) {
  const batch = db.batch();
  let taskCount = 0;

  const departureDate = runData.departureDate || new Date().toISOString().split('T')[0];

  // Cleaning frequency schedule (24h format)
  // Cleaning frequency schedule (24h format) - Updated to 1-hour intervals
  const cleaningTimes = [
    '06:00', '07:00', '08:00', '09:00', '10:00', '11:00', '12:00',
    '13:00', '14:00', '15:00', '16:00', '17:00', '18:00', '19:00',
    '20:00', '21:00'
  ];
  const garbageTimes = ['07:00', '10:00', '13:00', '16:00', '19:00'];
  const waterCheckTimes = ['08:00', '12:00', '16:00', '20:00'];
  const safetyInspectionTimes = ['10:00', '14:00', '18:00'];
  const pettyRepairTimes = ['10:00', '14:00', '18:00'];

  for (const coach of runData.coaches) {
    if (!coach.workerId) continue;
    const coachNo = coach.coachPosition || coach.coachNo || 'N/A';
    const workerId = coach.workerId;
    const workerName = coach.workerName || 'Unknown';

    // 1. Generate Toilet/Cleaning Task HEADERS (one per scheduled time)
    for (const timeSlot of cleaningTimes) {
      const hour = parseInt(timeSlot.split(':')[0]);
      let taskTypeName = 'toilet_cleaning';
      let displayName = 'Toilet Cleaning & Disinfection';
      if (hour >= 9 && hour < 20) {
        taskTypeName = 'coach_cleaning';
        displayName = 'Compartment Cleaning';
      }

      const headerId = `${runData.runInstanceId}_${taskTypeName}_${timeSlot.replace(':', '')}`;
      const headerRef = db.collection('task_headers').doc(headerId);
      batch.set(headerRef, {
        headerId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        taskType: taskTypeName,
        displayName,
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        status: 'PLANNED',
        taskSource: 'SYSTEM', // Categorization
        workerId,
        workerName,
        coachNo,
        childTaskIds: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });

      // Create child task detail for this coach/header
      const detailId = `${headerId}_${coachNo}`;
      const detailRef = db.collection('task_details').doc(detailId);
      batch.set(detailRef, {
        detailId,
        headerId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId,
        workerName,
        taskType: taskTypeName,
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        toiletStatus: null,
        washBasinStatus: null,
        dustbinStatus: null,
        consumableStatus: null,
        status: 'PLANNED',
        beforePhoto: null,
        afterPhoto: null,
        gpsLatitude: null,
        gpsLongitude: null,
        employeeId: workerId,
        deviceId: null,
        mobileNumber: null,
        checklist: null,
        remarks: null,
        completionTime: null,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });

      // Update header with child task id
      batch.update(headerRef, {
        childTaskIds: admin.firestore.FieldValue.arrayUnion(detailId)
      });

      taskCount++;
    }

    // 2. Generate Garbage Tasks
    for (const timeSlot of garbageTimes) {
      const garbageId = `${runData.runInstanceId}_garbage_${timeSlot.replace(':', '')}_${coachNo}`;
      const garbageRef = db.collection('garbage_tasks').doc(garbageId);
      batch.set(garbageRef, {
        id: garbageId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId,
        workerName,
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        isPreTerminal: false,
        status: 'PENDING',
        beforePhoto: null,
        afterPhoto: null,
        gpsLatitude: null,
        gpsLongitude: null,
        completedAt: null,
        createdAt: new Date().toISOString()
      });
      taskCount++;
    }

    // 3. Generate Water Check Tasks
    for (const timeSlot of waterCheckTimes) {
      const waterId = `${runData.runInstanceId}_water_${timeSlot.replace(':', '')}_${coachNo}`;
      const waterRef = db.collection('water_checks').doc(waterId);
      batch.set(waterRef, {
        id: waterId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId,
        workerName,
        checkTime: timeSlot,
        checkDate: departureDate,
        waterStatus: 'pending',
        lowWaterAlert: false,
        wateringPointSchedule: null,
        status: 'PENDING',
        photoUrl: null,
        completedAt: null,
        createdAt: new Date().toISOString()
      });
      taskCount++;
    }

    // 4. Generate Safety Check Tasks
    for (const timeSlot of safetyInspectionTimes) {
      const safetyId = `${runData.runInstanceId}_safety_${timeSlot.replace(':', '')}_${coachNo}`;
      const safetyRef = db.collection('safety_checks').doc(safetyId);
      batch.set(safetyRef, {
        id: safetyId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId,
        workerName,
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        fireExtinguisherStatus: null,
        fsdsStatus: null,
        cctvStatus: null,
        emergencyEquipmentStatus: null,
        photos: [],
        deficiencyReports: [],
        status: 'PENDING',
        remarks: null,
        completedAt: null,
        createdAt: new Date().toISOString()
      });
      taskCount++;
    }

    // 5. Generate Petty Repair Inspection Tasks
    for (const timeSlot of pettyRepairTimes) {
      const repairId = `${runData.runInstanceId}_repair_${timeSlot.replace(':', '')}_${coachNo}`;
      const repairRef = db.collection('petty_repairs').doc(repairId);
      batch.set(repairRef, {
        id: repairId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId,
        workerName,
        inspectionTime: timeSlot,
        inspectionDate: departureDate,
        items: {
          latches: 'ok',
          windows: 'ok',
          doors: 'ok',
          seats: 'ok',
          lights: 'ok',
          fans: 'ok',
          taps: 'ok',
          flush: 'ok'
        },
        isEscalated: false,
        escalatedTo: null,
        status: 'PENDING',
        remarks: null,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
      taskCount++;
    }
  }

  // ─── ATTENDANT LINEN TASKS (AC Coaches Only) ─────────────────────────────
  // Janitors get cleaning tasks (above). Attendants on AC coaches get linen
  // and berth-related tasks as per OBHS worker type rules.
  const linenChangeTimes = ['06:00', '14:00', '22:00'];
  const berthInspectionTimes = ['08:00', '18:00'];

  for (const coach of runData.coaches) {
    const isAC = /^[ABHME]/i.test(coach.coachPosition || '') ||
                 (coach.coachType && coach.coachType.toUpperCase().includes('AC'));

    if (!isAC || !coach.attendantId) continue;

    const coachNo = coach.coachPosition || coach.coachNo || 'N/A';
    const attendantId = coach.attendantId;
    const attendantName = coach.attendantName || 'Unknown Attendant';

    // 1. Linen Change Tasks (3 times a day for AC coaches)
    for (const timeSlot of linenChangeTimes) {
      const linenId = `${runData.runInstanceId}_linen_${timeSlot.replace(':', '')}_${coachNo}`;
      const linenRef = db.collection('linen_tasks').doc(linenId);
      batch.set(linenRef, {
        id: linenId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        coachType: coach.coachType || 'AC',
        assignedTo: attendantId,
        assignedToName: attendantName,
        workerType: 'ATTENDANT',
        taskType: 'LINEN_CHANGE',
        displayName: 'Linen Change & Berth Setup',
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        status: 'PENDING',
        linenItemsChecklist: {
          bedsheet: false,
          pillowCover: false,
          blanket: false,
          towel: false
        },
        beforePhoto: null,
        afterPhoto: null,
        remarks: null,
        completedAt: null,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
      taskCount++;
    }

    // 2. Berth Inspection Tasks (2 times a day for AC coaches)
    for (const timeSlot of berthInspectionTimes) {
      const berthId = `${runData.runInstanceId}_berth_${timeSlot.replace(':', '')}_${coachNo}`;
      const berthRef = db.collection('berth_inspections').doc(berthId);
      batch.set(berthRef, {
        id: berthId,
        runInstanceId: runData.runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        coachType: coach.coachType || 'AC',
        assignedTo: attendantId,
        assignedToName: attendantName,
        workerType: 'ATTENDANT',
        taskType: 'BERTH_INSPECTION',
        displayName: 'Berth & Cabin Inspection',
        scheduledTime: timeSlot,
        scheduledDate: departureDate,
        status: 'PENDING',
        berthChecklist: {
          cushionCondition: null,
          berthLatch: null,
          readingLight: null,
          acVent: null,
          windowBlind: null
        },
        deficiencies: [],
        beforePhoto: null,
        afterPhoto: null,
        remarks: null,
        completedAt: null,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
      taskCount++;
    }

    console.log(`[New Engine] Generated ${linenChangeTimes.length + berthInspectionTimes.length} attendant tasks for Attendant ${attendantName} on AC Coach ${coachNo}`);
  }
  // ─── END ATTENDANT TASKS ──────────────────────────────────────────────────


  if (taskCount > 0) {
    await batch.commit();
    console.log(`[New Engine] Generated ${taskCount} tasks for Run ${runData.runInstanceId}`);
  }
}

// ─── Generate terminal garbage tasks (when train approaches destination) ───
async function generatePreTerminalGarbageTasks(runInstanceId) {
  try {
    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) return;
    const runData = runDoc.data();
    const batch = db.batch();
    let count = 0;

    for (const coach of (runData.coaches || [])) {
      const effectiveWorkerId = coach.workerId || coach.janitorId || null;
      if (!effectiveWorkerId) continue;
      const coachNo = coach.coachPosition || coach.coachNo || 'N/A';
      const timeSlot = 'terminal';
      const garbageId = `${runInstanceId}_garbage_${timeSlot}_${coachNo}`;
      const garbageRef = db.collection('garbage_tasks').doc(garbageId);
      batch.set(garbageRef, {
        id: garbageId,
        runInstanceId,
        trainNo: runData.trainNo,
        coachNo,
        workerId: effectiveWorkerId,
        workerName: coach.workerName || coach.janitorName || 'Unknown',
        scheduledTime: timeSlot,
        scheduledDate: new Date().toISOString().split('T')[0],
        isPreTerminal: true,
        status: 'PENDING',
        beforePhoto: null,
        afterPhoto: null,
        gpsLatitude: null,
        gpsLongitude: null,
        completedAt: null,
        createdAt: new Date().toISOString()
      });
      count++;
    }

    if (count > 0) {
      await batch.commit();
      console.log(`Generated ${count} pre-terminal garbage tasks for ${runInstanceId}`);
    }
  } catch (error) {
    console.error('(PreTerminalGarbage) Error:', error);
  }
}

// =======================================================
// == API 4.2: Create Run Instance (Updated with PDF Rotation Logic)
// =======================================================
app.post('/api/run-instances', verifyToken, async (req, res) => {
  try {
    const { instanceId, coaches, departureDate } = req.body;

    if (!instanceId || !coaches || !Array.isArray(coaches) || !departureDate) {
      return res.status(400).send({
        error: 'Missing required fields',
        details: 'instanceId, coaches array, and departureDate are mandatory.'
      });
    }

    const existingCheck = await db.collection('RunInstance')
      .where('instanceId', '==', instanceId)
      .where('departureDate', '==', departureDate)
      .limit(1)
      .get();

    if (!existingCheck.empty) {
      return res.status(400).send({
        error: 'Rotation Conflict',
        details: `Instance '${instanceId}' is already assigned to a run on ${departureDate}.`
      });
    }

    const trainPairDoc = await db.collection('TrainPairs').doc(instanceId).get();
    if (!trainPairDoc.exists) {
      return res.status(404).send({ error: 'Train Instance (Rake) not found in Pool' });
    }
    const pairData = trainPairDoc.data();

    // Fetch parent train to get division, zone, and depot
    const parentTrainDoc = await db.collection('trains').doc(pairData.parentTrainId).get();
    const trainData = parentTrainDoc.exists ? parentTrainDoc.data() : {};

    const coachesWithNames = await Promise.all(coaches.map(async (c) => {
      const actualJanitorId = c.janitorId || c.workerId || null;
      let janitorName = c.janitorName || "Unknown Worker";

      if (actualJanitorId && (!c.janitorName || c.janitorName === "Unknown Worker")) {
        const workerDoc = await db.collection('users').doc(actualJanitorId).get();
        if (workerDoc.exists) {
          janitorName = workerDoc.data().fullName || workerDoc.data().name || "Unknown Worker";
        }
      }

      let actualAttendantName = c.attendantName || null;
      if (c.attendantId && !actualAttendantName) {
        const attDoc = await db.collection('users').doc(c.attendantId).get();
        if (attDoc.exists) {
          actualAttendantName = attDoc.data().fullName || attDoc.data().name || 'Unknown';
        }
      }

      return {
        coachPosition: c.coachPosition || null,
        coachType: c.coachType || null,
        janitorId: actualJanitorId,
        janitorName: janitorName,
        workerId: actualJanitorId,
        workerName: janitorName,
        attendantId: c.attendantId || null,
        attendantName: actualAttendantName,
        janitorTasks: c.janitorTasks || [],
        attendantTasks: c.attendantTasks || [],
        tasks: c.tasks || c.janitorTasks || [],
        attendanceStatus: 'Pending'
      };
    }));

    const { uid, name, email, role } = req.user;

    // Enforce maximum 2 coaches per worker (Railway Board policy)
    const workerCoachCount = {};
    for (const c of coachesWithNames) {
      if (c.workerId) {
        workerCoachCount[c.workerId] = (workerCoachCount[c.workerId] || 0) + 1;
        if (workerCoachCount[c.workerId] > 2) {
          return res.status(400).send({
            error: 'Coach Assignment Limit',
            details: `Worker ${c.workerId} has been assigned to more than 2 coaches. Maximum 2 coaches per janitor allowed.`
          });
        }
      }

      if (c.coachPosition) {
        const isAC = /^[ABHME]/i.test(c.coachPosition) || (c.coachType && c.coachType.toUpperCase().includes('AC'));
        if (isAC && !c.attendantId) {
          return res.status(400).send({
            error: 'Train Formation Error',
            details: `AC Coach ${c.coachPosition} must have an Attendant assigned.`
          });
        }
      }
    }

    const userName = name || email || role || 'Unknown';
    const runInstanceRef = db.collection('RunInstance').doc();

    const newRunData = {
      runInstanceId: runInstanceRef.id,
      instanceId: instanceId,
      departureDate: departureDate,
      trainNo: pairData.trainNo,
      trainName: pairData.trainName,
      inboundTrainNo: pairData.inboundTrainNo,
      outboundTrainNo: pairData.outboundTrainNo,
      parentTrainId: pairData.parentTrainId,
      division: trainData.division || null,
      zone: trainData.zone || null,
      depot: trainData.depot || null,
      journeyStartTime: pairData.journeyStartTime || null,
      journeyEndTime: pairData.journeyEndTime || null,
      scheduledDeparture: `${departureDate}T${pairData.journeyStartTime || '00:00:00'}.000Z`,

      numberOfCoaches: coachesWithNames.length,
      coaches: coachesWithNames,

      status: 'PLANNED',
      attendanceCaptured: false,
      taskExecutionScore: 0,
      actualDeparture: null,
      actualArrival: null,

      createdAt: new Date().toISOString(),
      createdBy: uid,
      createdByName: userName
    };

    await runInstanceRef.set(newRunData);

    await db.collection('TrainPairs').doc(instanceId).update({
      status: 'Active',
      lastAssignedDate: departureDate
    });

    res.status(201).send({
      message: 'Run Instance (Run Calendar Entry) created successfully',
      id: runInstanceRef.id,
      data: newRunData
    });

  } catch (error) {
    console.error('(RunInstance) Error:', error);
    res.status(500).send({ error: 'Failed to create run instance', details: error.message });
  }
});

// =======================================================
// == API 4.5: Update Run Instance (With Worker Names)
// =======================================================
app.put('/api/run-instances/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const { coaches, status } = req.body;

    if (!runInstanceId) {
      return res.status(400).send({ error: "runInstanceId is required in parameters." });
    }

    const runInstanceRef = db.collection('RunInstance').doc(runInstanceId);
    const doc = await runInstanceRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: "Run Instance not found." });
    }

    const { uid, name, email, role } = req.user;
    const userName = name || email || role || 'Unknown';

    const updateData = {
      updatedBy: uid,
      updatedByName: userName,
      updatedAt: new Date().toISOString()
    };

    if (coaches && Array.isArray(coaches)) {

      // AC coach prefix check for attendant validation
      const isACCoach = (coachType) => {
        if (!coachType) return false;
        const upper = coachType.toUpperCase();
        return ['A', 'B', 'H', 'M', 'C', 'E', 'AC', 'A1', 'B1', 'H1', 'M1', 'C1', 'E1',
          '2AC', '3AC', '1AC', 'AC', '2A', '3A', '1A', 'EA', 'EC', 'AB', 'HA', 'MA']
          .includes(upper) || /^[ABHMC]\d*$/.test(upper);
      };

      // Cache for worker names to avoid redundant DB calls
      const workerNamesCache = new Map();

      // Bulk Fetch Technique: Collect all unique worker IDs and fetch them in one go
      const uniqueWorkerIds = new Set();
      coaches.forEach(c => {
        if (c.janitorId || c.workerId) uniqueWorkerIds.add(c.janitorId || c.workerId);
        if (c.attendantId) uniqueWorkerIds.add(c.attendantId);
      });

      const workerIdList = Array.from(uniqueWorkerIds).filter(id => id && typeof id === 'string');
      
      if (workerIdList.length > 0) {
        // Fetch in batches of 30 (Firestore limit for 'in' queries)
        const batches = [];
        for (let i = 0; i < workerIdList.length; i += 30) {
          batches.push(workerIdList.slice(i, i + 30));
        }

        await Promise.all(batches.map(async (batch) => {
          const workersSnap = await db.collection('users')
            .where(admin.firestore.FieldPath.documentId(), 'in', batch)
            .get();
          
          workersSnap.forEach(doc => {
            const data = doc.data();
            workerNamesCache.set(doc.id, data.fullName || data.name || "Unknown Worker");
          });
        }));
      }

      const coachesWithNames = await Promise.all(coaches.map(async (c) => {
        const actualJanitorId = c.janitorId || c.workerId || null;
        let janitorName = c.janitorName || "Unknown Worker";

        if (actualJanitorId) {
          if (workerNamesCache.has(actualJanitorId)) {
            janitorName = workerNamesCache.get(actualJanitorId);
          } else if (!c.janitorName || c.janitorName === "Unknown Worker") {
            try {
              const workerDoc = await db.collection('users').doc(actualJanitorId).get();
              if (workerDoc.exists) {
                janitorName = workerDoc.data().fullName || workerDoc.data().name || "Unknown Worker";
                workerNamesCache.set(actualJanitorId, janitorName);
              }
            } catch (err) {
              console.error(`Error fetching janitor ${actualJanitorId}:`, err.message);
            }
          }
        }

        // Resolve attendant name if attendantId provided
        let actualAttendantName = c.attendantName || null;
        if (c.attendantId) {
          if (workerNamesCache.has(c.attendantId)) {
            actualAttendantName = workerNamesCache.get(c.attendantId);
          } else if (!actualAttendantName || actualAttendantName === 'Unknown') {
            try {
              const attDoc = await db.collection('users').doc(c.attendantId).get();
              if (attDoc.exists) {
                actualAttendantName = attDoc.data().fullName || attDoc.data().name || 'Unknown';
                workerNamesCache.set(c.attendantId, actualAttendantName);
              }
            } catch (err) {
              console.error(`Error fetching attendant ${c.attendantId}:`, err.message);
            }
          }
        }

        // Relaxed Validation: Only log warnings for inconsistencies instead of throwing errors 
        // to prevent UI "loading issues" or save failures during incremental assignments.
        const cType = c.coachType || '';
        const cPos = c.coachPosition || '';
        const isAC = isACCoach(cType) || /^[ABHMC]/i.test(cPos);

        if (c.attendantId && !isAC) {
          console.warn(`[Assignment Warning] Attendant ${c.attendantId} assigned to non-AC coach ${cPos}`);
        }
        if (isAC && !c.attendantId && (status === 'Active' || status === 'Ready')) {
          console.warn(`[Assignment Warning] AC Coach ${cPos} has no attendant in ${status} journey`);
        }

        return {
          coachPosition: cPos || null,
          coachNo: c.coachNo || cPos || null,
          coachType: cType || null,
          janitorId: actualJanitorId,
          janitorName: janitorName,
          workerId: actualJanitorId,
          workerName: janitorName,
          attendantId: c.attendantId || null,
          attendantName: actualAttendantName,
          janitorTasks: c.janitorTasks || [],
          attendantTasks: c.attendantTasks || [],
          tasks: c.tasks || c.janitorTasks || [],
          attendanceStatus: c.attendanceStatus || 'Pending'
        };
      }));

      updateData.numberOfCoaches = coachesWithNames.length;
      updateData.coaches = coachesWithNames;
    }

    if (status) {
      if (status.toUpperCase() === 'CLOSED' || status.toUpperCase() === 'COMPLETED') {
        const pendingTasksSnapshot = await db.collection('task_instances')
          .where('runInstanceId', '==', runInstanceId)
          .where('status', 'in', ['PENDING', 'SUBMITTED'])
          .get();

        if (!pendingTasksSnapshot.empty) {
          return res.status(400).send({
            error: `Cannot close run instance. There are still ${pendingTasksSnapshot.size} pending or unreviewed tasks.`
          });
        }
      }
      updateData.status = status;
    }

    await runInstanceRef.update(updateData);

    const updatedDoc = await runInstanceRef.get();
    
    res.status(200).send({
      message: 'Run Instance updated successfully',
      runInstanceId: runInstanceId,
      data: { id: updatedDoc.id, ...updatedDoc.data() }
    });

  } catch (error) {
    console.error('(RunInstance) Error updating:', error);
    res.status(500).send({ error: 'Failed to update run instance', details: error.message });
  }
});

// =======================================================
// == API 4.6: Activate Journey (Set Actual Departure)
// =======================================================
app.post('/api/run-instances/:runInstanceId/activate', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const { actualDeparture } = req.body;

    if (!runInstanceId) {
      return res.status(400).send({ error: 'runInstanceId is required.' });
    }

    const runInstanceRef = db.collection('RunInstance').doc(runInstanceId);
    const doc = await runInstanceRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: 'Run Instance not found.' });
    }

    const currentData = doc.data();
    if (!['ALLOCATED', 'READY'].includes(currentData.status)) {
      return res.status(400).send({
        error: 'Invalid State Transition',
        details: `Cannot activate a journey with status '${currentData.status}'. Must be 'ALLOCATED' or 'READY'.`
      });
    }

    const departureTime = actualDeparture || new Date().toISOString();

    await runInstanceRef.update({
      actualDeparture: departureTime,
      status: 'Active',
      journeyStartedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    // ─── Rule-driven task generation on actual departure ───
    const runDocAfter = await runInstanceRef.get();
    const activatedRunData = runDocAfter.data();
    await generateTaskInstancesForRun(activatedRunData);
    await generatePreTerminalGarbageTasks(runInstanceId);
    // Also generate from task_masters rules (CM/CA configurable)
    await generateTasksFromMasters(activatedRunData);

    res.status(200).send({
      message: 'Journey activated successfully',
      runInstanceId,
      actualDeparture: departureTime,
      status: 'Active'
    });

  } catch (error) {
    console.error('(Journey Activation) Error:', error);
    res.status(500).send({ error: 'Failed to activate journey', details: error.message });
  }
});

// =======================================================
// == API 4.7: Complete Journey (Set Actual Arrival)
// =======================================================
app.post('/api/run-instances/:runInstanceId/complete', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const { actualArrival, delayMinutes } = req.body;

    if (!runInstanceId) {
      return res.status(400).send({ error: 'runInstanceId is required.' });
    }

    const runInstanceRef = db.collection('RunInstance').doc(runInstanceId);
    const doc = await runInstanceRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: 'Run Instance not found.' });
    }

    const currentData = doc.data();
    if (currentData.status !== 'Active') {
      return res.status(400).send({
        error: 'Invalid State Transition',
        details: `Cannot complete a journey with status '${currentData.status}'. Must be 'Active'.`
      });
    }

    const arrivalTime = actualArrival || new Date().toISOString();
    const updateData = {
      actualArrival: arrivalTime,
      status: 'Completed',
      journeyCompletedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    if (delayMinutes !== undefined) {
      updateData.delayMinutes = delayMinutes;
    }

    await runInstanceRef.update(updateData);

    // Generate journey closure tasks
    await generateClosureTasks(runInstanceId, arrivalTime, currentData);

    await logAudit({
      action: 'JOURNEY_COMPLETED',
      entityType: 'RunInstance',
      entityId: runInstanceId,
      actorId: req.user.uid,
      actorRole: req.user.role,
      actorName: req.user.fullName,
      newValue: { actualArrival: arrivalTime, status: 'Completed' },
      details: `Journey completed at ${arrivalTime}${delayMinutes ? ` with ${delayMinutes}min delay` : ''}`
    });

    res.status(200).send({
      message: 'Journey completed successfully',
      runInstanceId,
      actualArrival: arrivalTime,
      status: 'Completed'
    });

  } catch (error) {
    console.error('(Journey Completion) Error:', error);
    res.status(500).send({ error: 'Failed to complete journey', details: error.message });
  }
});

// =======================================================
// == API 4.2: Edit an existing Train (Fixed Version)
// =======================================================
// NOTE: convertToDecimalDays function ko yahan se hata diya hai 
// kyunki wo file mein upar pehle se declared hai.

app.put('/api/trains/:uid', verifyToken, async (req, res) => {
  try {
    // ─── ROLE GUARD: Only Admins and Masters can edit trains ─────────────
    const callerRole = (req.user.role || '').toUpperCase();
    const blockedFromEditing = ['CONTRACTOR SUPERVISOR', 'CTS', 'RAILWAY SUPERVISOR', 'WORKER', 'RAILWAY WORKER'];
    if (blockedFromEditing.includes(callerRole)) {
      return res.status(403).json({
        error: 'Access Denied',
        details: `Your role (${req.user.role}) is not authorized to modify trains.`
      });
    }
    // ─── END ROLE GUARD ──────────────────────────────────────────────────

    const { uid } = req.params;

    const allowedFields = [
      'trainNo', 'trainName', 'origin', 'destination', 'days',
      'zone', 'division', 'depot', 'status', 'TrainApplicableFor',
      'outboundTrainNo', 'inboundTrainNo', 'returnOffset', 'cycleLength',
      'outboundDurationStr', 'inboundDurationStr', 'layoverDestStr', 'layoverOriginStr',
      'journeyStartTime'
    ];

    const bodyKeys = Object.keys(req.body);
    for (const key of bodyKeys) {
      if (!allowedFields.includes(key)) {
        return res.status(400).send({
          error: `Invalid field name.`,
          details: `The field '${key}' is not allowed.`
        });
      }
    }

    const {
      trainNo, trainName, origin, destination, days,
      zone, division, depot, status, TrainApplicableFor,
      outboundTrainNo, inboundTrainNo, returnOffset, cycleLength,
      outboundDurationStr, inboundDurationStr, layoverDestStr, layoverOriginStr
    } = req.body;

    if (!uid) return res.status(400).send({ error: "Train ID (UID) is required." });

    const docRef = db.collection('trains').doc(uid);
    const doc = await docRef.get();
    if (!doc.exists) return res.status(404).send({ error: "Train not found." });

    const existingData = doc.data();

    // Check if OBHS is active
    const isOBHSNow = TrainApplicableFor ? TrainApplicableFor.includes('OBHS') : existingData.TrainApplicableFor.includes('OBHS');

    let newRequiredInstances = existingData.requiredInstances || 0;
    let newFinalCycleTime = existingData.cycleLength || 0;

    if (isOBHSNow) {
      // PDF Formula Calculation for Update
      const outbound = outboundDurationStr !== undefined ? convertToDecimalDays(outboundDurationStr) : convertToDecimalDays(existingData.outboundDurationStr);
      const inbound = inboundDurationStr !== undefined ? convertToDecimalDays(inboundDurationStr) : convertToDecimalDays(existingData.inboundDurationStr);
      const layoverD = layoverDestStr !== undefined ? convertToDecimalDays(layoverDestStr) : convertToDecimalDays(existingData.layoverDestStr);
      const layoverO = layoverOriginStr !== undefined ? convertToDecimalDays(layoverOriginStr) : convertToDecimalDays(existingData.layoverOriginStr);

      const calculatedC = outbound + inbound + layoverD + layoverO;
      newFinalCycleTime = calculatedC > 0 ? calculatedC : (Number(cycleLength) || existingData.cycleLength || 0);

      const finalDays = days || existingData.days || [];
      if (finalDays.length === 0) {
        return res.status(400).send({ error: 'OBHS requires at least one departure day.' });
      }

      const numDays = (finalDays.includes('All Days') || finalDays.includes('Daily')) ? 7 : finalDays.length;
      const F = 7 / numDays;
      newRequiredInstances = Math.ceil(newFinalCycleTime / F);
      if (newRequiredInstances <= 0) newRequiredInstances = 1;
    }

    // Use editorId/editorName already defined at line 3857
    // const { uid: editorId, name, email, role } = req.user;
    // const editorName = name || email || role || 'Unknown';
    const updateData = {};
    if (trainNo !== undefined) updateData.trainNo = trainNo;
    if (trainName !== undefined) updateData.trainName = trainName;
    if (origin !== undefined) updateData.origin = origin;
    if (destination !== undefined) updateData.destination = destination;
    if (days !== undefined) updateData.days = days;
    if (zone !== undefined) updateData.zone = zone;
    if (division !== undefined) updateData.division = division;
    if (depot !== undefined) updateData.depot = depot;
    if (TrainApplicableFor !== undefined) updateData.TrainApplicableFor = TrainApplicableFor;
    if (outboundTrainNo !== undefined) updateData.outboundTrainNo = outboundTrainNo;
    if (inboundTrainNo !== undefined) updateData.inboundTrainNo = inboundTrainNo;
    if (status) updateData.status = status;

    // Fixed: isOBHSNow variable used here
    if (isOBHSNow) {
      updateData.cycleLength = Number(newFinalCycleTime.toFixed(4));
      updateData.requiredInstances = newRequiredInstances;
      if (outboundDurationStr) updateData.outboundDurationStr = outboundDurationStr;
      if (inboundDurationStr) updateData.inboundDurationStr = inboundDurationStr;
      if (layoverDestStr) updateData.layoverDestStr = layoverDestStr;
      if (layoverOriginStr) updateData.layoverOriginStr = layoverOriginStr;
      if (journeyStartTime) updateData.journeyStartTime = journeyStartTime;
    }

    updateData.updatedBy = editorId;
    updateData.updatedByName = editorName;
    updateData.updatedAt = new Date().toISOString();
    updateData.status = 'PENDING';

    await docRef.update(updateData);

    const needsPairUpdate = trainNo || trainName || days || outboundDurationStr || inboundDurationStr || layoverDestStr || layoverOriginStr || outboundTrainNo || inboundTrainNo;

    if (isOBHSNow && needsPairUpdate) {
      const finalTrainNo = trainNo || existingData.trainNo;
      const finalTrainName = trainName || existingData.trainName;
      const finalInbound = inboundTrainNo || existingData.inboundTrainNo;
      const finalOutbound = outboundTrainNo || existingData.outboundTrainNo;

      const oldPairs = await db.collection('TrainPairs').where('parentTrainId', '==', uid).get();
      const deleteBatch = db.batch();
      oldPairs.forEach(doc => deleteBatch.delete(doc.ref));
      await deleteBatch.commit();

      const createBatch = db.batch();
      for (let i = 0; i < newRequiredInstances; i++) {
        const instanceLetter = String.fromCharCode(65 + i);
        const instanceId = `${finalTrainNo}-${finalTrainName}-Inst-${instanceLetter}`;

        createBatch.set(db.collection('TrainPairs').doc(instanceId), {
          instanceId: instanceId,
          instanceName: `Instance ${instanceLetter}`,
          trainNo: finalTrainNo,
          trainName: finalTrainName,
          status: 'Inactive',
          inboundTrainNo: finalInbound,
          outboundTrainNo: finalOutbound,
          rotationPattern: `Round-Robin`,
          parentTrainId: uid,
          updatedAt: new Date().toISOString()
        });
      }
      await createBatch.commit();
    }

    res.status(200).send({
      message: `Train and associated TrainPairs updated successfully.`,
      uid: uid,
      calculatedInstances: newRequiredInstances,
      updatedData: updateData
    });

  } catch (error) {
    console.error('(Train) Error updating train:', error);
    res.status(500).send({ error: 'Failed to update train', details: error.message });
  }
});

// =======================================================
// == API 4.3: Get Train Pairs (D1, D2, etc.) by Train ID
// =======================================================
app.get('/api/train-pairs/train/:parentTrainId', verifyToken, async (req, res) => {
  try {
    const { parentTrainId } = req.params;
    const { status } = req.query;

    if (!parentTrainId) {
      return res.status(400).send({
        error: "Missing parameter",
        details: "parentTrainId is required to fetch pairs."
      });
    }

    let query = db.collection('TrainPairs').where('parentTrainId', '==', parentTrainId);

    if (status) {
      const formattedStatus = status.charAt(0).toUpperCase() + status.slice(1).toLowerCase();
      query = query.where('status', '==', formattedStatus);
    }

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({
        message: "No Train Pairs found for this Train ID.",
        count: 0,
        data: []
      });
    }

    const rawPairs = [];
    snapshot.forEach(doc => {
      rawPairs.push({ id: doc.id, ...doc.data() });
    });

    // Deduplicate by instanceId (safety net — each Inst-A / Inst-B should only appear once)
    const seen = new Set();
    const pairs = rawPairs.filter(p => {
      const key = p.instanceId || p.id;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    pairs.sort((a, b) => {
      const idA = a.instanceId || a.id || '';
      const idB = b.instanceId || b.id || '';
      return idA.localeCompare(idB, undefined, { numeric: true });
    });

    res.status(200).send({
      message: "Train Pairs fetched successfully",
      count: pairs.length,
      data: pairs
    });

  } catch (error) {
    console.error('(TrainPairs) Error fetching data:', error);
    res.status(500).send({
      error: 'Failed to fetch train pairs',
      details: error.message
    });
  }
});


// =======================================================
// == API 4.3: Get Run Instances by Train ID (With Status Filter)
// =======================================================
app.get('/api/run-instances/train/:parentTrainId', verifyToken, async (req, res) => {
  try {
    const { parentTrainId } = req.params;
    const { status } = req.query;
    if (!parentTrainId) {
      return res.status(400).send({
        error: "Missing parameter",
        details: "parentTrainId is required to fetch instances."
      });
    }

    let query = db.collection('RunInstance').where('parentTrainId', '==', parentTrainId);

    if (status) {
      const formattedStatus = status.charAt(0).toUpperCase() + status.slice(1).toLowerCase();
      query = query.where('status', '==', formattedStatus);
    }

    const snapshot = await query.orderBy('createdAt', 'desc').get();

    if (snapshot.empty) {
      return res.status(200).json({
        message: "No Run Instances found matching the criteria.",
        count: 0,
        data: []
      });
    }

    const instances = [];
    snapshot.forEach(doc => {
      instances.push({ id: doc.id, ...doc.data() });
    });

    res.status(200).send({
      message: "Run Instances fetched successfully",
      count: instances.length,
      data: instances
    });

  } catch (error) {
  }
  console.error('(RunInstance) Error fetching data:', error);
  res.status(500).send({
    error: 'Failed to fetch run instances',
    details: error.message
  });
}
);


// =======================================================
// == API 4.5.5: Generate Run Instances Schedule automatically
// =======================================================
app.post('/api/trains/:trainId/generate-schedule', verifyToken, async (req, res) => {
  try {
    const { trainId } = req.params;
    const { startDate, endDate } = req.body;

    if (!startDate || !endDate) {
      return res.status(400).send({ error: "startDate and endDate are required (YYYY-MM-DD)." });
    }

    const trainDoc = await db.collection('trains').doc(trainId).get();
    if (!trainDoc.exists) {
      return res.status(404).send({ error: "Train not found." });
    }

    const trainData = trainDoc.data();
    const days = trainData.days || [];

    if (days.length === 0) {
      return res.status(400).send({ error: "Train has no assigned running days." });
    }

    const trainPairsSnapshot = await db.collection('TrainPairs')
      .where('parentTrainId', '==', trainId)
      .get();

    if (trainPairsSnapshot.empty) {
      return res.status(400).send({ error: "No Rake Instances (TrainPairs) found for this train. Please update train to generate pairs." });
    }

    const trainPairs = [];
    trainPairsSnapshot.forEach(doc => trainPairs.push({ id: doc.id, ...doc.data() }));
    // Sort them so Inst-A comes before Inst-B
    trainPairs.sort((a, b) => a.id.localeCompare(b.id));

    const start = new Date(startDate);
    const end = new Date(endDate);

    if (isNaN(start) || isNaN(end) || start > end) {
      return res.status(400).send({ error: "Invalid date range." });
    }

    const dayMap = {
      'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3,
      'Thursday': 4, 'Friday': 5, 'Saturday': 6
    };

    const validDays = days.map(d => dayMap[d]);

    const generatedInstances = [];
    let pairIndex = 0;

    // Check existing RunInstances to avoid duplicates
    const existingRunsSnapshot = await db.collection('RunInstance')
      .where('parentTrainId', '==', trainId)
      .get();

    const existingDates = new Set();
    existingRunsSnapshot.forEach(doc => {
      existingDates.add(doc.data().departureDate);
    });

    const batch = db.batch();
    const { uid, name, email, role } = req.user;
    const userName = name || email || role || 'System';

    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      if (validDays.includes(d.getDay())) {
        const dateStr = d.toISOString().split('T')[0];

        if (existingDates.has(dateStr)) {
          continue; // Skip if already generated
        }

        const pair = trainPairs[pairIndex % trainPairs.length];
        pairIndex++;

        const runInstanceRef = db.collection('RunInstance').doc();

        const newRunData = {
          runInstanceId: runInstanceRef.id,
          instanceId: pair.id,
          departureDate: dateStr,
          trainNo: pair.trainNo || trainData.trainNo || trainData.trainNumber,
          trainName: pair.trainName || trainData.trainName,
          inboundTrainNo: pair.inboundTrainNo || trainData.inboundTrainNo || null,
          outboundTrainNo: pair.outboundTrainNo || trainData.outboundTrainNo || null,
          parentTrainId: trainId,
          division: trainData.division || null,
          zone: trainData.zone || null,
          depot: trainData.depot || null,
          coaches: [], // Empty initially
          numberOfCoaches: 0,
          createdAt: new Date().toISOString(),
          createdBy: uid,
          createdByName: userName,
          status: 'Active'
        };

        batch.set(runInstanceRef, newRunData);
        generatedInstances.push(newRunData);
      }
    }

    if (generatedInstances.length > 0) {
      await batch.commit();
    }

    res.status(200).send({
      message: `Generated ${generatedInstances.length} new instances successfully.`,
      count: generatedInstances.length,
      data: generatedInstances
    });

  } catch (error) {
    console.error('(Generate Schedule) Error:', error);
    res.status(500).send({ error: 'Failed to generate schedule', details: error.message });
  }
});



// =======================================================
// == API 4.6: Get All Run Instances (Global Fetch)
// =======================================================
app.get('/api/run-instances', verifyToken, async (req, res) => {
  try {
    // Optional: Filter lagane ke liye query params
    const { status } = req.query;

    let query = db.collection('RunInstance');

    if (status) {
      query = query.where('status', '==', status);
    }

    const { role, division: userDiv, trainId: userTrainId, trainIds: userTrainIds } = req.user;
    const roleUpper = (role || '').toUpperCase();

    const snapshot = await query.orderBy('createdAt', 'desc').get();

    if (snapshot.empty) {
      return res.status(200).json({
        success: true,
        count: 0,
        data: []
      });
    }

    const allInstances = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      let isVisible = false;

      // ─── ADMIN RBAC ───
      if (roleUpper === 'SUPER ADMIN' || roleUpper === 'COMPANY MASTER' || roleUpper === 'RAILWAY MASTER') {
        isVisible = true; // Global views
      }
      // ─── DIVISIONAL ADMIN / RAILWAY ADMIN ───
      else if (roleUpper.includes('ADMIN')) {
        if (data.division === userDiv) {
          isVisible = true;
        }
      }
      // ─── SUPERVISOR RBAC ───
      else if (roleUpper === 'RAILWAY SUPERVISOR') {
        const assignedTrains = userTrainIds || (userTrainId ? [userTrainId] : []);
        if (assignedTrains.includes(data.parentTrainId) || assignedTrains.includes(data.trainId)) {
          isVisible = true;
        }
      }
      else if (roleUpper === 'CTS' || roleUpper === 'CONTRACTOR SUPERVISOR') {
        if (data.parentTrainId === userTrainId || data.trainId === userTrainId) {
          isVisible = true;
        }
      }
      // Workers generally don't use this list, but if they do, they see their own runs
      else if (roleUpper.includes('WORKER')) {
        const carriesWorker = (data.coaches || []).some(c => c.workerId === req.user.uid);
        if (carriesWorker) isVisible = true;
      }

      if (isVisible) {
        allInstances.push({ id: doc.id, ...data });
      }
    });

    res.status(200).json({
      success: true,
      message: "Run Instances fetched with RBAC filters",
      count: allInstances.length,
      data: allInstances
    });

  } catch (error) {
    // Indexing error handling
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Firestore Index Missing',
        details: 'Composite index is required for filtering and sorting.'
      });
    }

    console.error('(RunInstance) Error fetching all data:', error);
    res.status(500).send({
      success: false,
      error: 'Failed to fetch all run instances',
      details: error.message
    });
  }
});

// GET /api/v1/obhs-run/:runId — Run instance with nested janitor/attendant/tasks format
app.get('/api/v1/obhs-run/:runId', verifyToken, async (req, res) => {
  try {
    const { runId } = req.params;
    const doc = await db.collection('RunInstance').doc(runId).get();
    if (!doc.exists) return res.status(404).json({ error: 'Run instance not found' });

    const data = doc.data();
    const coaches = (data.coaches || []).map(c => ({
      coachNo: c.coachPosition || c.coachNo || null,
      isAc: c.coachType ? ['A', 'B', 'H', 'M', 'C', 'E', 'AC', 'A1']
        .some(prefix => c.coachType.toUpperCase().startsWith(prefix)) : false,
      janitor: c.workerId ? { workerId: c.workerId, name: c.workerName || 'Unknown' } : null,
      attendant: c.attendantId ? { workerId: c.attendantId, name: c.attendantName || 'Unknown' } : null,
      tasks: c.tasks || []
    }));

    res.status(200).json({
      success: true,
      data: {
        runInstanceId: runId,
        trainNo: data.trainNo,
        trainName: data.trainName,
        status: data.status,
        coaches,
        actualDeparture: data.actualDeparture,
        scheduledArrival: data.scheduledArrival,
        zone: data.zone, division: data.division, depot: data.depot,
        createdAt: data.createdAt,
        updatedAt: data.updatedAt
      }
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch run instance', details: error.message });
  }
});

// DELETE /api/run-instances/:runInstanceId — Delete a run instance
app.delete('/api/run-instances/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const ref = db.collection('RunInstance').doc(runInstanceId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Run instance not found' });
    }
    const data = doc.data();
    await ref.delete();
    // If the instance had a linked TrainPair, reset its status
    if (data && data.instanceId) {
      try {
        await db.collection('TrainPairs').doc(data.instanceId).update({ status: 'Available' });
      } catch (_) { }
    }
    res.status(200).json({ success: true, message: 'Run instance deleted successfully' });
  } catch (error) {
    console.error('(RunInstance) Error deleting:', error);
    res.status(500).json({ error: 'Failed to delete run instance', details: error.message });
  }
});

// =======================================================
// == STATION CLEANING RUN APIs
// =======================================================
app.post('/api/station-runs', verifyToken, async (req, res) => {
  try {
    const data = req.body;
    if (!data.stationId || !data.stationName || !data.date || !data.shift) {
      return res.status(400).json({ error: 'Missing required fields for station run' });
    }

    // Add metadata
    data.status = 'Pending';
    data.createdAt = new Date().toISOString();
    data.updatedAt = data.createdAt;

    // Generate an ID (e.g., SCR-stationCode-date-shift)
    const runId = `SCR-${data.stationName.substring(0, 3).toUpperCase()}-${Date.now()}`;
    data.runInstanceId = runId;

    await db.collection('StationCleaningRuns').doc(runId).set(data);
    res.status(201).json({ success: true, message: 'Station Run created successfully', data });
  } catch (error) {
    console.error('(StationRun) Error creating:', error);
    res.status(500).json({ error: 'Failed to create Station Run', details: error.message });
  }
});

app.get('/api/station-runs', verifyToken, async (req, res) => {
  try {
    const snapshot = await db.collection('StationCleaningRuns').orderBy('createdAt', 'desc').get();
    const runs = [];
    snapshot.forEach(doc => {
      runs.push({ id: doc.id, ...doc.data() });
    });
    res.status(200).json({ success: true, data: runs });
  } catch (error) {
    console.error('(StationRun) Error fetching:', error);
    res.status(500).json({ error: 'Failed to fetch Station Runs', details: error.message });
  }
});

app.put('/api/station-runs/:runId', verifyToken, async (req, res) => {
  try {
    const { runId } = req.params;
    const data = req.body;
    data.updatedAt = new Date().toISOString();

    const ref = db.collection('StationCleaningRuns').doc(runId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Station Run not found' });
    }

    await ref.update(data);
    res.status(200).json({ success: true, message: 'Station Run updated successfully' });
  } catch (error) {
    console.error('(StationRun) Error updating:', error);
    res.status(500).json({ error: 'Failed to update Station Run', details: error.message });
  }
});

// GET /api/station-runs/my-runs — Worker sees only their assigned runs
app.get('/api/station-runs/my-runs', verifyToken, async (req, res) => {
  try {
    const { uid } = req.user;
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    const snapshot = await db.collection('StationCleaningRuns')
      .where('createdAt', '>=', thirtyDaysAgo)
      .orderBy('createdAt', 'desc')
      .limit(100)
      .get();
    const myRuns = [];
    snapshot.forEach(doc => {
      const data = { id: doc.id, ...doc.data() };
      const platforms = data.platforms || [];
      const isAssigned = platforms.some(p => p.janitorId === uid);
      if (isAssigned) myRuns.push(data);
    });
    res.status(200).json({ success: true, data: myRuns });
  } catch (error) {
    console.error('(StationRun) Error fetching worker runs:', error);
    res.status(500).json({ error: 'Failed to fetch worker station runs', details: error.message });
  }
});

app.delete('/api/station-runs/:runId', verifyToken, async (req, res) => {
  try {
    const { runId } = req.params;
    const role = (req.user.role || '').toLowerCase();
    if (!role.includes('admin') && !role.includes('master') && !role.includes('supervisor')) {
      return res.status(403).json({ error: 'Only admins, masters, or supervisors can delete station runs' });
    }

    const ref = db.collection('StationCleaningRuns').doc(runId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Station Run not found' });
    }

    await ref.delete();
    res.status(200).json({ success: true, message: 'Station Run deleted successfully' });
  } catch (error) {
    console.error('(StationRun) Error deleting:', error);
    res.status(500).json({ error: 'Failed to delete Station Run', details: error.message });
  }
});

// =======================================================
// == STATION CLEANING TASKS APIs
// =======================================================

app.post('/api/station-tasks/submit', verifyToken, async (req, res) => {
  try {
    const data = req.body;
    if (!data.runInstanceId || !data.platformNumber) {
      return res.status(400).json({ error: 'Missing runInstanceId or platformNumber' });
    }
    data.status = 'Completed'; // worker submitted
    data.submittedAt = new Date().toISOString();

    // Save to station_tasks collection
    const result = await db.collection('station_tasks').add(data);
    res.status(201).json({ success: true, message: 'Task submitted', taskId: result.id });
  } catch (error) {
    res.status(500).json({ error: 'Failed to submit task', details: error.message });
  }
});

app.get('/api/station-tasks/pending-review', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    let query = db.collection('station_tasks').where('status', '==', 'Completed');
    if (runInstanceId) {
      query = query.where('runInstanceId', '==', runInstanceId);
    }

    const snapshot = await query.get();
    const tasks = [];
    snapshot.forEach(doc => {
      tasks.push({ id: doc.id, ...doc.data() });
    });
    res.status(200).json({ success: true, count: tasks.length, data: tasks });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch pending tasks', details: error.message });
  }
});

// =======================================================
// == OBHS TASK APIs — superseded by newer handlers below
// =======================================================
// Old board + submit handlers removed (they queried task_instances and
// conflicted with the newer OBHS handlers at ~line 7957+ that use obhs_tasks).

// 3. Raise Complaint / Report Issue (Worker) — handled at ~line 8446
// 3. Get Pending Review Tasks for Supervisor
app.get('/api/tasks/pending-review', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;

    let query = db.collection('obhs_tasks')
      .where('status', '==', 'Completed');

    if (runInstanceId) {
      query = query.where('runInstanceId', '==', runInstanceId);
    }

    const snapshot = await query.get();

    const tasks = [];
    snapshot.forEach(doc => {
      const task = doc.data();
      task.id = doc.id;
      tasks.push(task);
    });
    res.status(200).json({ success: true, count: tasks.length, data: tasks });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 4. Approve Task & Provide Supervisor Score
app.post('/api/tasks/:taskId/approve', verifyToken, async (req, res) => {
  try {
    const { taskId } = req.params;
    const { supervisorScore, supervisorComments } = req.body; // Score out of 10

    if (supervisorScore === undefined) {
      return res.status(400).json({ error: "Supervisor score is required" });
    }

    const taskRef = db.collection('obhs_tasks').doc(taskId);
    await taskRef.update({
      status: 'APPROVED',
      supervisorScore: Number(supervisorScore),
      supervisorComments: supervisorComments || '',
      approvedBy: req.user.uid,
      approvedAt: new Date().toISOString()
    });

    res.status(200).json({ success: true, message: 'Task approved and scored' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 5. Reject Task
app.post('/api/tasks/:taskId/reject', verifyToken, async (req, res) => {
  try {
    const { taskId } = req.params;
    const { rejectionReason } = req.body;

    if (!rejectionReason) {
      return res.status(400).json({ error: 'Rejection reason is mandatory' });
    }

    const taskRef = db.collection('obhs_tasks').doc(taskId);
    await taskRef.update({
      status: 'REJECTED',
      rejectionReason: rejectionReason,
      rejectedBy: req.user.uid,
      rejectedAt: new Date().toISOString()
    });

    res.status(200).json({ success: true, message: 'Task rejected' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 6. Passenger Feedback (Consolidated Score Logic)
app.post('/api/tasks/:taskId/passenger-feedback', async (req, res) => {
  try {
    const { taskId } = req.params;
    const { passengerScore, passengerPhone, feedbackText } = req.body; // Score out of 10

    const taskRef = db.collection('task_instances').doc(taskId);
    const doc = await taskRef.get();

    if (!doc.exists) return res.status(404).json({ error: 'Task not found' });

    const taskData = doc.data();
    const supScore = taskData.supervisorScore || 0;

    // Weightage: Passenger = 70%, Supervisor = 30%
    const pScore = Number(passengerScore);
    const consolidatedScore = (pScore * 0.7) + (supScore * 0.3);

    await taskRef.update({
      passengerScore: pScore,
      passengerPhone,
      passengerFeedback: feedbackText || '',
      consolidatedScore: consolidatedScore,
      feedbackReceivedAt: new Date().toISOString()
    });

    res.status(200).json({
      success: true,
      message: 'Feedback submitted successfully',
      consolidatedScore
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// =======================================================
// GET TRAINS (Updated with Strict TrainApplicableFor Filter)
// =======================================================
app.get('/api/trains', verifyToken, async (req, res) => {
  try {
    const { role, zone: userZone, division: userDivision } = req.user;
    const { status, zone: queryZone, division: queryDivision, applicableFor } = req.query;

    const userRole = (role || '').trim().toLowerCase();
    let query = db.collection('trains');

    if (userRole === 'company master' || userRole === 'super admin') {
      if (queryZone) query = query.where('zone', '==', queryZone);
      if (queryDivision) query = query.where('division', '==', queryDivision);
    }
    else if (userRole === 'railway master') {
      if (!userZone) return res.status(403).send({ error: "Railway Master profile mein Zone missing hai." });
      query = query.where('zone', '==', userZone);
      if (queryDivision) query = query.where('division', '==', queryDivision);
    }
    else if (userRole === 'cts' || userRole === 'contractor supervisor') {
      // Contractor Supervisors are mapped to exactly ONE train — return only that
      const { trainId: userTrainId } = req.user;
      if (!userTrainId) {
        return res.status(200).json({ count: 0, trains: [], message: 'No train assigned to your profile.' });
      }
      const trainDoc = await db.collection('trains').doc(userTrainId).get();
      if (!trainDoc.exists) return res.status(200).json({ count: 0, trains: [] });
      return res.status(200).json({ count: 1, trains: [{ uid: trainDoc.id, ...trainDoc.data() }] });
    }
    else if (userRole.includes('admin') || userRole.includes('supervisor')) {
      if (!userDivision) return res.status(403).send({ error: "Supervisor profile mein Division missing hai." });
      query = query.where('division', '==', userDivision);
      if (userZone) query = query.where('zone', '==', userZone);
    }

    if (applicableFor) {
      query = query.where('TrainApplicableFor', 'array-contains', applicableFor);
    }

    if (status) {
      query = query.where('status', '==', status);
    }

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, trains: [] });
    }

    const trainList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      trainList.push({ uid: doc.id, ...data });
    });

    trainList.sort((a, b) => {
      const trainA = String(a.trainNo || "");
      const trainB = String(b.trainNo || "");
      return trainA.localeCompare(trainB, undefined, { numeric: true });
    });

    res.status(200).json({
      count: trainList.length,
      trains: trainList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Firestore Index Missing. Please check Firebase Console link in logs.' });
    }
    console.error('(GetTrains) Error:', error);
    res.status(500).send({ error: 'Failed to fetch trains', details: error.message });
  }
});

// --- API 4.4: Get details for a SINGLE Train by UID ---
app.get('/api/trains/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    if (!uid) return res.status(400).send({ error: "Train ID (UID) is required." });

    const docRef = db.collection('trains').doc(uid);
    const doc = await docRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Train not found." });

    const trainData = doc.data();
    if (!trainData.TrainApplicableFor) {
      trainData.TrainApplicableFor = [];
    }

    res.status(200).json(trainData);
  } catch (error) {
    console.error('(Train) Error fetching single train:', error);
    res.status(500).send({ error: 'Failed to fetch train', details: error.message });
  }
});

// --- API 4.5: Get details for a SINGLE Train by Train NUMBER ---
app.get('/api/trains/number/:trainNo', verifyToken, async (req, res) => {
  try {
    const { trainNo } = req.params;
    if (!trainNo) return res.status(400).send({ error: "Train Number is required." });

    const trainsRef = db.collection('trains');
    const snapshot = await trainsRef.where('trainNo', '==', trainNo).limit(1).get();

    if (snapshot.empty) {
      return res.status(404).send({ error: "Train not found with this number." });
    }

    const doc = snapshot.docs[0];
    const trainData = { uid: doc.id, ...doc.data() };
    if (!trainData.TrainApplicableFor) {
      trainData.TrainApplicableFor = [];
    }

    res.status(200).json(trainData);
  } catch (error) {
    console.error('(Train) Error fetching single train by number:', error);
    res.status(500).send({ error: 'Failed to fetch train', details: error.message });
  }
});

// =======================================================
// == 5. COACH CLEANING FORM APIs (*** WORKFLOW UPDATED ***)
// =======================================================
const generateFormId = async (formType, division) => {
  const div = division ? division.substring(0, 2).toUpperCase() : 'XX';

  let type = 'PC';
  if (formType === 'coach') {
    type = 'CC';
  } else if (formType === 'cts') {
    type = 'CTS';
  }

  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = String(now.getFullYear()).substring(2);
  const dateStr = `${day}${month}${year}`;

  const counterId = `${div}-${type}-${dateStr}`;
  const counterRef = db.collection('counters').doc(counterId);

  let sequentialNumber = 1;

  try {
    await db.runTransaction(async (transaction) => {
      const counterDoc = await transaction.get(counterRef);
      if (!counterDoc.exists) {
        sequentialNumber = 1;
        transaction.set(counterRef, { count: sequentialNumber });
      } else {
        sequentialNumber = counterDoc.data().count + 1;
        transaction.update(counterRef, { count: sequentialNumber });
      }
    });
  } catch (e) {
    console.error("Transaction failure:", e);
    throw new Error("Failed to generate form ID sequence.");
  }

  const sequentialStr = String(sequentialNumber).padStart(2, '0');
  return `${div}-${type}-${dateStr}-${sequentialStr}`;
};

// --- API 5.2: (Contractor) Submits a new Coach Form (FINAL UPDATED with Validation) ---
app.post('/api/coach-forms', verifyToken, async (req, res) => {
  try {
    const contractorSupervisor = req.user;

    const {
      trainId,
      formDateTime, coachCount, machinesUsed, chemicals,
      manpower, submittedTo, signature,
      contractId
    } = req.body;

    if (!trainId || !formDateTime || !chemicals || !manpower || !submittedTo || !signature || !contractId) {
      return res.status(400).send({ error: "Please fill all mandatory fields (including Contract)." });
    }

    const contractDoc = await db.collection('contracts').doc(contractId).get();
    if (!contractDoc.exists) {
      return res.status(404).send({ error: "Selected Contract not found." });
    }

    const contractData = contractDoc.data();
    const status = (contractData.status || "").toLowerCase();
    if (status !== 'active') {
      return res.status(400).send({ error: `Selected Contract is ${contractData.status} (Not Active).` });
    }

    if (contractData.endDate) {
      const today = new Date();
      const endDate = new Date(contractData.endDate);
      endDate.setHours(23, 59, 59, 999);
      if (today > endDate) {
        return res.status(400).send({ error: "Selected Contract has Expired." });
      }
    }

    const categories = contractData.workCategories || [];
    let hasCoachAccess = false;
    if (Array.isArray(categories)) {
      hasCoachAccess = categories.some(cat => cat.toLowerCase().includes('coach'));
    } else if (typeof categories === 'string') {
      hasCoachAccess = categories.toLowerCase().includes('coach');
    }

    if (!hasCoachAccess) {
      return res.status(403).send({ error: "This Contract does not allow Coach Cleaning work." });
    }


    let trainName = null;
    let fetchedTrainNo = "";
    if (trainId) {
      const trainDoc = await db.collection('trains').doc(trainId).get();
      if (trainDoc.exists) {
        const tData = trainDoc.data();
        trainName = tData.trainName || null;
        fetchedTrainNo = tData.trainNo || tData.trainNumber || "";
      }
    }

    let supervisorName = null;
    if (submittedTo && submittedTo.railwayEmployeeId) {
      const userDoc = await db.collection('users').doc(submittedTo.railwayEmployeeId).get();
      if (userDoc.exists) {
        supervisorName = userDoc.data().fullName || null;
      }
    }

    const newFormId = await generateFormId('coach', contractorSupervisor.division);
    const entityName = req.user.entityName || contractData.entityName || contractData.agencyName || "Unknown Agency";

    const docRef = db.collection('coachForms').doc(newFormId);
    await docRef.set({
      uid: newFormId,
      formId: newFormId,
      trainId: trainId,
      trainName: trainName,
      trainNumber: fetchedTrainNo,
      formDateTime: formDateTime,
      coachCount: coachCount || null,
      machinesUsed: machinesUsed || [],
      chemicals: chemicals,
      manpower: manpower || [],
      submittedTo: {
        railwayEmployeeId: submittedTo.railwayEmployeeId || null,
        railwayEmployeeName: supervisorName,
        division: submittedTo.division || null,
        depot: submittedTo.depot || null
      },
      signature: {
        name: signature.name || null,
        date: signature.date || null
      },
      contractId: contractId,

      status: 'SUBMITTED',
      submittedById: contractorSupervisor.uid,
      submittedByName: contractorSupervisor.fullName,
      submittedByZone: contractorSupervisor.zone || null,
      submittedByDivision: contractorSupervisor.division || null,
      submittedByDepot: contractorSupervisor.depot || null,
      submittedByEntityId: contractorSupervisor.entityId || null,
      submittedByEntityName: entityName,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.status(201).send({
      message: 'Coach form submitted successfully.',
      uid: newFormId
    });

  } catch (error) {
    console.error('(CoachForm) Error submitting form:', error);
    res.status(500).send({ error: 'Failed to submit form', details: error.message });
  }
});

// --- API 5.3: Get Coach Forms (STRICT FIXED VERSION) ---
app.get('/api/coach-forms', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone, division, entityId } = req.user;
    const { status, type } = req.query;

    let query = db.collection('coachForms');
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master') || userRole === 'company master';
    const isAdmin = userRole.includes('admin');

    if (userType === 'railway') {
      if (isMaster) {
        if (zone) {
          query = query.where('submittedByZone', '==', zone);
          console.log(`(CoachForm) Railway Master Locked to Zone: ${zone}`);
        } else {
          console.log('(CoachForm) Railway Master (No Zone in Token) - Showing All');
        }
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedTo.railwayEmployeeId', '==', uid);
      }
    }
    else if (userType === 'contractor') {
      if (!entityId) return res.status(403).send({ error: "Company ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) {
        console.log(`(CoachForm) Contractor Master: Viewing All Company Data`);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }

    if (status) {
      query = query.where('status', '==', status);
    }
    else if (type === 'history') {
      query = query.where('status', 'in', ['SCORED', 'LOCKED', 'AUTO-APPROVED', 'REJECTED_BY_RAILWAY']);
    }
    else {

      query = query.where('status', 'in', ['SUBMITTED', 'RE-SUBMITTED', 'APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS']);

      console.log(`(CoachForm) Active View for ${role}`);
    }

    const snapshot = await query.orderBy('createdAt', 'desc').get();
    const list = [];
    snapshot.forEach(d => list.push(d.data()));

    console.log(`(CoachForm) Total Found: ${list.length}`);
    res.status(200).json({ count: list.length, forms: list });

  } catch (e) {
    if (e.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index Missing. Check Firebase Console.' });
    }
    res.status(500).send({ error: e.message });
  }
});
// --- NAYI API 5.4: (Railway) Approves Manpower (Unlocks Scoring) ---
app.post('/api/coach-forms/:formId/approve-manpower', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const railwaySupervisor = req.user;

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to approve this form." });
    }

    if (formData.status === 'SUBMITTED' || formData.status === 'RE-SUBMITTED') {
      await formDocRef.update({
        status: 'APPROVED_BY_RAILWAY',
        manpowerApprovedAt: admin.firestore.FieldValue.serverTimestamp()
      });
      console.log(`(CoachForm) Form ${formId} manpower approved.`);
      return res.status(200).send({ message: 'Form approved for scoring.' });
    }

    res.status(400).send({ message: `Form status is already ${formData.status}.` });

  } catch (error) {
    console.error('(CoachForm) Error approving manpower:', error);
    res.status(500).send({ error: 'Failed to approve manpower', details: error.message });
  }
});

// --- NAYI API 5.5: (Railway) Saves Scoring Draft ---
app.put('/api/coach-forms/:formId/save-scoring-draft', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const {
      workType,
      acwpStatus,
      coachEvaluationTable,
      railwayRemarks
    } = req.body;
    const railwaySupervisor = req.user;

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to rate this form." });
    }

    if (formData.status !== 'APPROVED_BY_RAILWAY' && formData.status !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot save draft. Status is ${formData.status}. Please approve manpower first.` });
    }

    const penaltyMap = { 'A': 0, 'B': 50, 'C': 100, 'D': 200 };
    let totalPenalty = 0;

    const summary = {
      internal: { A: 0, B: 0, C: 0, D: 0 },
      external: { A: 0, B: 0, C: 0, D: 0 },
      intensive: { A: 0, B: 0, C: 0, D: 0 },
      toiletries: { Yes: 0, No: 0, NA: 0 },
      watering: { Yes: 0, No: 0, NA: 0 },
      doorsLocking: { Yes: 0, No: 0, NA: 0 },
      totalCoaches: 0,
    };

    const processedEvaluationTable = coachEvaluationTable ? coachEvaluationTable.map(coach => {
      summary.totalCoaches++;
      const internalPenalty = penaltyMap[coach.internalCleaning] || 0;
      const externalPenalty = penaltyMap[coach.externalCleaning] || 0;
      const intensivePenalty = penaltyMap[coach.intensiveCleaning] || 0;

      if (coach.internalCleaning) summary.internal[coach.internalCleaning]++;
      if (coach.externalCleaning) summary.external[coach.externalCleaning]++;
      if (coach.intensiveCleaning) summary.intensive[coach.intensiveCleaning]++;

      if (coach.toiletries) summary.toiletries[coach.toiletries]++;
      if (coach.watering) summary.watering[coach.watering]++;
      if (coach.doorsLocking) summary.doorsLocking[coach.doorsLocking]++;

      const coachPenalty = internalPenalty + externalPenalty + intensivePenalty;
      totalPenalty += coachPenalty;

      return { ...coach, penalty: coachPenalty };
    }) : [];


    await formDocRef.update({
      status: 'SCORING_IN_PROGRESS',
      ratingDetails: {
        workType: workType || null,
        acwpStatus: acwpStatus || null,
        coachEvaluationTable: processedEvaluationTable || null,
        totalPenalty: totalPenalty,
        summary: summary
      },
      railwayRemarks: railwayRemarks || null,
      scoringLastSavedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(CoachForm) Form ${formId} draft saved with penalty ${totalPenalty}.`);
    res.status(200).send({ message: 'Scoring draft saved successfully.' });

  } catch (error) {
    console.error('(CoachForm) Error saving draft:', error);
    res.status(500).send({ error: 'Failed to save draft', details: error.message });
  }
});


// --- NAYI API 5.6: (Railway) Submits Final Scoring ---
app.post('/api/coach-forms/:formId/submit-scoring', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const {
      workType,
      acwpStatus,
      coachEvaluationTable,
      railwayRemarks,
      railwaySignatureName,
      railwaySignatureDate
    } = req.body;
    const railwaySupervisor = req.user;

    if (!workType || !acwpStatus || !coachEvaluationTable || !railwaySignatureName || !railwaySignatureDate) {
      return res.status(400).send({ error: "Work Type, ACWP Status, Evaluation Table, and Signature are required for final submission." });
    }

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to rate this form." });
    }

    if (formData.status !== 'APPROVED_BY_RAILWAY' && formData.status !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot score form. Status is ${formData.status}.` });
    }

    const penaltyMap = { 'A': 0, 'B': 50, 'C': 100, 'D': 200 };
    let totalPenalty = 0;

    const summary = {
      internal: { A: 0, B: 0, C: 0, D: 0 },
      external: { A: 0, B: 0, C: 0, D: 0 },
      intensive: { A: 0, B: 0, C: 0, D: 0 },
      toiletries: { Yes: 0, No: 0, NA: 0 },
      watering: { Yes: 0, No: 0, NA: 0 },
      doorsLocking: { Yes: 0, No: 0, NA: 0 },
      totalCoaches: 0,
    };

    const processedEvaluationTable = coachEvaluationTable.map(coach => {
      summary.totalCoaches++;
      const internalPenalty = penaltyMap[coach.internalCleaning] || 0;
      const externalPenalty = penaltyMap[coach.externalCleaning] || 0;
      const intensivePenalty = penaltyMap[coach.intensiveCleaning] || 0;

      if (coach.internalCleaning) summary.internal[coach.internalCleaning]++;
      if (coach.externalCleaning) summary.external[coach.externalCleaning]++;
      if (coach.intensiveCleaning) summary.intensive[coach.intensiveCleaning]++;

      if (coach.toiletries) summary.toiletries[coach.toiletries]++;
      if (coach.watering) summary.watering[coach.watering]++;
      if (coach.doorsLocking) summary.doorsLocking[coach.doorsLocking]++;

      const coachPenalty = internalPenalty + externalPenalty + intensivePenalty;
      totalPenalty += coachPenalty;

      return { ...coach, penalty: coachPenalty };
    });

    await formDocRef.update({
      status: 'SCORED',
      scoringInProgress: false,
      ratingDetails: {
        workType: workType,
        acwpStatus: acwpStatus,
        coachEvaluationTable: processedEvaluationTable,
        totalPenalty: totalPenalty,
        summary: summary
      },
      railwaySignature: {
        name: railwaySignatureName || railwaySupervisor.fullName,
        date: railwaySignatureDate || new Date().toISOString()
      },
      railwayRemarks: railwayRemarks || null,
      ratedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(CoachForm) Form ${formId} has been SCORED with penalty ${totalPenalty}.`);
    res.status(200).send({ message: 'Form successfully scored and sent to contractor.' });

  } catch (error) {
    console.error('(CoachForm) Error submitting rating:', error);
    res.status(500).send({ error: 'Failed to submit rating', details: error.message });
  }
});


// --- API 5.7: (Contractor) Approves/Accepts the Rating (Contractor Approved) ---
app.post('/api/coach-forms/:formId/accept-rating', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to accept this rating." });
    }

    const currentStatus = doc.data().status;
    if (currentStatus === 'LOCKED' || currentStatus === 'AUTO-APPROVED') {
      return res.status(400).send({ error: "This his form is already locked." });
    }

    if (currentStatus !== 'SCORED') {
      return res.status(400).send({ error: `Cannot accept rating. Form status is ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'LOCKED',
      completedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(CoachForm) Form ${formId} rating accepted and LOCKED.`);
    res.status(200).send({ message: 'Rating accepted. Form is now locked.' });

  } catch (error) {
    console.error('(CoachForm) Error accepting rating:', error);
    res.status(500).send({ error: 'Failed to accept rating', details: error.message });
  }
});

// --- API 5.8: (Contractor) Re-submits the form (Updated with Train Name, Number & No Auto-Remarks) ---
app.put('/api/coach-forms/:formId/resubmit', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const {
      trainId, coachCount, machinesUsed, chemicals,
      manpower, signature,
      contractorRemarks,
      resubmitSign
    } = req.body;

    if (!resubmitSign) {
      return res.status(400).send({ error: "Resubmit Signature is required to submit the form again." });
    }

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to resubmit this form." });
    }

    if (doc.data().status !== 'REJECTED_BY_RAILWAY') {
      return res.status(400).send({ error: `Cannot resubmit form. Status is ${doc.data().status}` });
    }

    let trainName = doc.data().trainName;
    let trainNumber = doc.data().trainNumber;

    if (trainId) {
      const trainDoc = await db.collection('trains').doc(trainId).get();
      if (trainDoc.exists) {
        const trainData = trainDoc.data();
        trainName = trainData.trainName;
        trainNumber = trainData.trainNumber;
      }
    }

    const updateData = {
      trainId: trainId || doc.data().trainId,
      trainName: trainName,
      trainNumber: trainNumber,
      coachCount: coachCount || doc.data().coachCount,
      machinesUsed: machinesUsed || doc.data().machinesUsed,
      chemicals: chemicals || doc.data().chemicals,
      manpower: manpower || doc.data().manpower,
      signature: signature || doc.data().signature,
      resubmitSignature: resubmitSign,
      status: 'RE-SUBMITTED',

      resubmittedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    if (contractorRemarks) {
      updateData.contractorRemarks = contractorRemarks;
    }

    await formDocRef.update(updateData);

    console.log(`(CoachForm) Form ${formId} re-submitted with Train: ${trainName} (${trainNumber})`);
    res.status(200).send({ message: 'Form has been re-submitted.' });

  } catch (error) {
    console.error('(CoachForm) Error resubmitting form:', error);
    res.status(500).send({ error: 'Failed to resubmit form', details: error.message });
  }
});

app.post('/api/coach-forms/:formId/reject', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const { rejectionComments } = req.body;
    const railwaySupervisor = req.user;

    if (!rejectionComments) {
      return res.status(400).send({ error: "Rejection comments are required." });
    }

    const formDocRef = db.collection('coachForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to reject this form." });
    }

    const currentStatus = doc.data().status;
    if (currentStatus !== 'SUBMITTED' && currentStatus !== 'RE-SUBMITTED' && currentStatus !== 'APPROVED_BY_RAILWAY' && currentStatus !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot reject a form with status: ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'REJECTED_BY_RAILWAY',
      rejectionComments: rejectionComments,
      rejectedAt: admin.firestore.FieldValue.serverTimestamp()

    });

    console.log(`(CoachForm) Form ${formId} has been REJECTED.`);
    res.status(200).send({ message: 'Form successfully rejected.' });

  } catch (error) {
    console.error('(CoachForm) Error rejecting form:', error);
    res.status(500).send({ error: 'Failed to reject form', details: error.message });
  }
});

// --- API 5.10: (Railway) Gets forms that are ready for SCORING ---
app.get('/api/coach-forms/pending-scoring', verifyToken, async (req, res) => {
  try {
    const railwaySupervisorId = req.user.uid;

    let query = db.collection('coachForms');

    query = query.where('submittedTo.railwayEmployeeId', '==', railwaySupervisorId);

    query = query.where('status', 'in', ['APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS']);

    query = query.orderBy('createdAt', 'desc');
    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, forms: [] });
    }

    const formList = [];
    snapshot.forEach(doc => {
      formList.push(doc.data());
    });

    res.status(200).json({
      count: formList.length,
      forms: formList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: `Firebase needs a composite index for this query. ${error.message}`
      });
    }
    console.error('(CoachForm) Error fetching pending-scoring forms:', error);
    res.status(500).send({ error: 'Failed to fetch forms', details: error.message });
  }
});

app.get('/api/coach-forms/submitted', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, entityId: tokenEntityId, division: tokenDivision } = req.user;
    const { entityId, division, contractId, startDate, endDate } = req.query;

    let targetEntityId = entityId;
    if (userType === 'contractor') {
      targetEntityId = tokenEntityId;
    }
    if (!targetEntityId && userType !== 'railway') {
      return res.status(400).send({ error: "Entity ID is required." });
    }

    let query = db.collection('coachForms');
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    query = query.where('submittedByEntityId', '==', targetEntityId);

    if (userType === 'contractor') {
      if (isMaster) {
        if (division) query = query.where('submittedByDivision', '==', division);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', tokenDivision);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }
    else if (userType === 'railway' && division) {
      query = query.where('submittedByDivision', '==', division);
    }

    if (contractId) {
      query = query.where('contractId', '==', contractId);
    }

    const snapshot = await query.orderBy('createdAt', 'desc').get();
    let formList = [];

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;
      if (start || end) {
        if (!data.formDateTime) {
          include = false;
        } else {
          const fDate = new Date(data.formDateTime);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }
      if (include) formList.push(data);
    });

    res.status(200).json({ count: formList.length, forms: formList });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index Missing', details: error.message });
    }
    console.error('(CoachForm) Error:', error);
    res.status(500).send({ error: error.message });
  }
});


// =======================================================
// == 5.14 STATS API (Updated with Coach, Premises, and CTS)
// =======================================================

app.get('/api/all-forms/stats', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone: userZone, division: userDiv, entityId: userEntityId } = req.user;

    const {
      startDate, endDate, zone, division, depot,
      contractId, entityId,
      formType
    } = req.query;

    console.log(`(Stats) Request by ${role} - Filter: ${formType || 'ALL'}`);

    const buildQuery = (collectionName) => {
      let query = db.collection(collectionName);
      const userRole = (role || '').toLowerCase();
      const isMaster = userRole.includes('master');
      const isAdmin = userRole.includes('admin');

      if (userType === 'railway') {
        if (isMaster) {
          const targetZone = zone || userZone;
          if (targetZone) query = query.where('submittedByZone', '==', targetZone);
        } else if (isAdmin) {
          query = query.where('submittedByDivision', '==', userDiv);
        } else {
          query = query.where('submittedTo.railwayEmployeeId', '==', uid);
        }
      }
      else if (userType === 'contractor') {
        if (!userEntityId) throw new Error("Company ID missing.");
        query = query.where('submittedByEntityId', '==', userEntityId);

        if (isMaster) { /* See All for that entity */ }
        else if (isAdmin) {
          query = query.where('submittedByDivision', '==', userDiv);
        }
        else {
          query = query.where('submittedById', '==', uid);
        }
      }

      if (division) query = query.where('submittedByDivision', '==', division);
      if (depot) query = query.where('submittedByDepot', '==', depot);
      if (contractId) query = query.where('contractId', '==', contractId);
      if (entityId && userType === 'railway') query = query.where('submittedByEntityId', '==', entityId);

      return query;
    };

    let fetchCoach = true;
    let fetchPremises = true;
    let fetchCTS = true;

    if (formType === 'coach') {
      fetchPremises = false;
      fetchCTS = false;
    } else if (formType === 'premises') {
      fetchCoach = false;
      fetchCTS = false;
    } else if (formType === 'cts') {
      fetchCoach = false;
      fetchPremises = false;
    }

    const promises = [];

    if (fetchCoach) {
      promises.push(buildQuery('coachForms').get());
    } else {
      promises.push(Promise.resolve({ empty: true, forEach: () => { } }));
    }

    if (fetchPremises) {
      promises.push(buildQuery('premisesForms').get());
    } else {
      promises.push(Promise.resolve({ empty: true, forEach: () => { } }));
    }

    if (fetchCTS) {
      promises.push(buildQuery('ctsForms').get());
    } else {
      promises.push(Promise.resolve({ empty: true, forEach: () => { } }));
    }

    const [coachSnapshot, premiseSnapshot, ctsSnapshot] = await Promise.all(promises);

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    const stats = {
      total: 0,
      pending: 0,
      manpowerApproved: 0,
      rejected: 0,
      scoringProgress: 0,
      autoApproved: 0,
      locked: 0
    };

    const processDoc = (doc) => {
      if (!doc || !doc.exists) return;

      const data = doc.data();
      let include = true;

      if (start || end) {
        const dateToCompare = data.formDateTime || data.createdAt?.toDate?.()?.toISOString();
        if (!dateToCompare) {
          include = false;
        } else {
          const fDate = new Date(dateToCompare);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }

      if (include) {
        stats.total++;
        const s = data.status;

        if (['SUBMITTED', 'RE-SUBMITTED'].includes(s)) stats.pending++;
        else if (s === 'APPROVED_BY_RAILWAY') stats.manpowerApproved++;
        else if (s === 'REJECTED_BY_RAILWAY') stats.rejected++;
        else if (['SCORED', 'SCORING_IN_PROGRESS'].includes(s)) stats.scoringProgress++;
        else if (s === 'AUTO-APPROVED') stats.autoApproved++;
        else if (s === 'LOCKED') stats.locked++;
      }
    };

    if (fetchCoach) coachSnapshot.forEach(doc => processDoc(doc));
    if (fetchPremises) premiseSnapshot.forEach(doc => processDoc(doc));
    if (fetchCTS) ctsSnapshot.forEach(doc => processDoc(doc));

    console.log(`(Stats) Final Count - Total: ${stats.total}`);
    res.status(200).json(stats);

  } catch (e) {
    console.error('(Stats) Error:', e);
    res.status(500).send({ error: e.message });
  }
});

// =======================================================
// == 6. PREMISES CLEANING FORM APIs (*** WORKFLOW UPDATED ***)
// =======================================================

// --- API 6.1: Submit Premises Form (FINAL FIXED: With Entity Name) ---
app.post('/api/premises-forms', verifyToken, async (req, res) => {
  try {
    const { location, contractId, submittedTo, area, ...body } = req.body;

    if (!location || !contractId || !submittedTo) {
      return res.status(400).send({ error: "Mandatory fields missing." });
    }

    const contractDoc = await db.collection('contracts').doc(contractId).get();
    if (!contractDoc.exists) {
      return res.status(404).send({ error: "Selected Contract not found." });
    }

    const contractData = contractDoc.data();

    const status = (contractData.status || "").toLowerCase();
    if (status !== 'active') {
      return res.status(400).send({ error: `Selected Contract is ${contractData.status} (Not Active).` });
    }

    if (contractData.endDate) {
      const today = new Date();
      const endDate = new Date(contractData.endDate);
      endDate.setHours(23, 59, 59, 999);

      if (today > endDate) {
        return res.status(400).send({ error: "Selected Contract has Expired." });
      }
    }

    const cats = contractData.workCategories;
    let hasPremiseAccess = false;

    if (Array.isArray(cats)) {
      hasPremiseAccess = cats.some(c => c.toLowerCase().includes('premise'));
    } else if (typeof cats === 'string') {
      hasPremiseAccess = cats.toLowerCase().includes('premise');
    }

    if (!hasPremiseAccess) {
      return res.status(403).send({ error: "This Contract does not allow Premises Cleaning work." });
    }
    const areaMap = { 'GICC': 23530, 'OWS': 8630, 'NWS': 10130 };
    const cleanLoc = location.trim();
    const calculatedArea = areaMap[cleanLoc] || area || 0;

    let supervisorName = null;
    if (submittedTo.railwayEmployeeId) {
      const u = await db.collection('users').doc(submittedTo.railwayEmployeeId).get();
      if (u.exists) supervisorName = u.data().fullName;
    }

    const entityName = req.user.entityName || contractData.agencyName || contractData.entityName || "Unknown Agency";

    const newFormId = await generateFormId('premises', req.user.division);

    await db.collection('premisesForms').doc(newFormId).set({
      uid: newFormId,
      formId: newFormId,
      location,
      area: calculatedArea,
      contractId,
      ...body,
      submittedTo: { ...submittedTo, railwayEmployeeName: supervisorName },
      status: 'SUBMITTED',

      submittedById: req.user.uid,
      submittedByName: req.user.fullName,
      submittedByZone: req.user.zone,
      submittedByDivision: req.user.division,
      submittedByDepot: req.user.depot,

      submittedByEntityId: req.user.entityId,
      submittedByEntityName: entityName,

      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.status(201).send({ message: 'Premises form submitted.', uid: newFormId });

  } catch (e) {
    console.error('(PremisesForm) Error:', e);
    res.status(500).send({ error: e.message });
  }
});
// --- API 6.3: Get Premises Forms (FINAL FIXED: No Master Overlap) ---
app.get('/api/premises-forms', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone, division, entityId } = req.user;
    const { status, type } = req.query;

    let query = db.collection('premisesForms');
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master') || userRole === 'company master';
    const isAdmin = userRole.includes('admin');

    console.log(`(PremisesForm) Request by: ${userType} | Role: ${role} | Zone: ${zone || 'NA'}`);

    // =========================================================
    // 1. SECURITY & HIERARCHY FILTERS (Strict Logic)
    // =========================================================
    if (userType === 'railway') {
      if (isMaster) {
        if (zone) {
          query = query.where('submittedByZone', '==', zone);
        }
      } else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      } else {
        query = query.where('submittedTo.railwayEmployeeId', '==', uid);
      }
    }
    else if (userType === 'contractor') {
      if (!entityId) return res.status(403).send({ error: "Company ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (!isMaster) {
        if (isAdmin) {
          query = query.where('submittedByDivision', '==', division);
        } else {
          query = query.where('submittedById', '==', uid);
        }
      }
    }

    // =========================================================
    // 2. STATUS & HISTORY FILTER (Fixed Logic to prevent 2-2 forms)
    // =========================================================
    if (status) {
      query = query.where('status', '==', status);
    }
    else if (type === 'history') {
      query = query.where('status', 'in', ['SCORED', 'LOCKED', 'AUTO-APPROVED', 'REJECTED_BY_RAILWAY']);
    }
    else {
      query = query.where('status', 'in', ['SUBMITTED', 'RE-SUBMITTED', 'APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS']);
    }

    const snapshot = await query.orderBy('createdAt', 'desc').get();
    const list = [];
    snapshot.forEach(d => list.push(d.data()));

    console.log(`(PremisesForm) Found ${list.length} forms for ${role}`);
    res.status(200).json({ count: list.length, forms: list });

  } catch (e) {
    if (e.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index Missing. Check Firebase Console.' });
    }
    res.status(500).send({ error: e.message });
  }
});
// --- NAYI API 6.4: (Railway) Approves Manpower (Unlocks Scoring) ---
app.post('/api/premises-forms/:formId/approve-manpower', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const railwaySupervisor = req.user;

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to approve this form." });
    }

    if (formData.status === 'SUBMITTED' || formData.status === 'RE-SUBMITTED') {
      await formDocRef.update({
        status: 'APPROVED_BY_RAILWAY',
        manpowerApprovedAt: admin.firestore.FieldValue.serverTimestamp()
      });
      console.log(`(PremisesForm) Form ${formId} manpower approved.`);
      return res.status(200).send({ message: 'Form approved for scoring.' });
    }

    res.status(400).send({ message: `Form status is already ${formData.status}.` });

  } catch (error) {
    console.error('(PremisesForm) Error approving manpower:', error);
    res.status(500).send({ error: 'Failed to approve manpower', details: error.message });
  }
});
const calculateSectionStats = (items) => {
  if (!items || !Array.isArray(items) || items.length === 0) {
    return { processed: [], avg: 0, avgPct: 0 };
  }

  const MAX_SCORE = 10;

  let totalAvg = 0;

  const processed = items.map(item => {
    let s1 = Number(item.score1) || 0;
    let s2 = Number(item.score2) || 0;

    if (s1 > MAX_SCORE) s1 = MAX_SCORE;
    if (s2 > MAX_SCORE) s2 = MAX_SCORE;

    const itemAvg = (s1 + s2) / 2;

    const itemPct = (itemAvg / MAX_SCORE) * 100;

    totalAvg += itemAvg;

    return {
      ...item,
      score1: s1,
      score2: s2,
      avg: parseFloat(itemAvg.toFixed(2)),
      avgPercentage: parseFloat(itemPct.toFixed(2)) + '%'
    };
  });

  const sectionAvg = totalAvg / items.length;

  const sectionAvgPct = (sectionAvg / MAX_SCORE) * 100;

  return {
    processed,
    avg: parseFloat(sectionAvg.toFixed(2)),
    avgPct: parseFloat(sectionAvgPct.toFixed(2))
  };
};
// --- 6.5: (Railway) Saves Scoring Draft (UPDATED FOR PREMISES) ---
app.put('/api/premises-forms/:formId/save-scoring-draft', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const {
      housekeepingItems,
      pitLineItems,
      disposalItems,
      railwayRemarks
    } = req.body;
    const railwaySupervisor = req.user;

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to rate this form." });
    }

    if (formData.status !== 'APPROVED_BY_RAILWAY' && formData.status !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot save draft. Status is ${formData.status}. Please approve manpower first.` });
    }

    const housekeeping = calculateSectionStats(housekeepingItems);
    const pitLine = calculateSectionStats(pitLineItems);
    const disposal = calculateSectionStats(disposalItems);

    const overallAverage = (housekeeping.avg + pitLine.avg + disposal.avg) / 3;
    const overallAveragePct = (housekeeping.avgPct + pitLine.avgPct + disposal.avgPct) / 3;

    await formDocRef.update({
      status: 'SCORING_IN_PROGRESS',
      ratingDetails: {
        housekeepingItems: housekeeping.processed,
        pitLineItems: pitLine.processed,
        disposalItems: disposal.processed,

        summary: {
          housekeepingAveragePct: housekeeping.avgPct + '%',
          pitLineAveragePct: pitLine.avgPct + '%',
          garbageDisposalAveragePct: disposal.avgPct + '%',
          overallAveragePct: parseFloat(overallAveragePct.toFixed(2)) + '%'
        }
      },
      railwayRemarks: railwayRemarks || null,
      scoringLastSavedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(PremisesForm) Draft saved. Overall Pct: ${overallAveragePct.toFixed(2)}%`);
    res.status(200).send({ message: 'Scoring draft saved successfully.' });

  } catch (error) {
    console.error('(PremisesForm) Error saving draft:', error);
    res.status(500).send({ error: 'Failed to save draft', details: error.message });
  }
});

// --- API 6.6: (Railway) Submits Final Scoring (UPDATED: With Scores & %) ---
app.post('/api/premises-forms/:formId/submit-scoring', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const {
      housekeepingItems,
      pitLineItems,
      disposalItems,
      railwayRemarks,
      railwaySignatureName,
      railwaySignatureDate
    } = req.body;
    const railwaySupervisor = req.user;
    if (!housekeepingItems || !pitLineItems || !disposalItems || !railwaySignatureName) {
      return res.status(400).send({ error: "All sections and signature are required." });
    }

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "Unauthorized." });
    }

    if (formData.status !== 'APPROVED_BY_RAILWAY' && formData.status !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot score. Status: ${formData.status}` });
    }

    const housekeeping = calculateSectionStats(housekeepingItems);
    const pitLine = calculateSectionStats(pitLineItems);
    const disposal = calculateSectionStats(disposalItems);

    const overallScore = (housekeeping.avg + pitLine.avg + disposal.avg) / 3;

    const overallAveragePct = (housekeeping.avgPct + pitLine.avgPct + disposal.avgPct) / 3;

    await formDocRef.update({
      status: 'SCORED',
      scoringInProgress: false,
      ratingDetails: {
        housekeepingItems: housekeeping.processed,
        pitLineItems: pitLine.processed,
        disposalItems: disposal.processed,
        summary: {
          housekeepingAveragePct: housekeeping.avgPct.toFixed(2) + '%',
          pitLineAveragePct: pitLine.avgPct.toFixed(2) + '%',
          garbageDisposalAveragePct: disposal.avgPct.toFixed(2) + '%',
          overallAveragePct: overallAveragePct.toFixed(2) + '%',

          housekeepingScore: housekeeping.avg.toFixed(2),
          pitLineScore: pitLine.avg.toFixed(2),
          garbageDisposalScore: disposal.avg.toFixed(2),
          overallScore: overallScore.toFixed(2)
        }
      },
      railwaySignature: { name: railwaySignatureName, date: railwaySignatureDate || new Date().toISOString() },
      railwayRemarks: railwayRemarks || null,
      ratedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.status(200).send({ message: 'Form successfully scored.' });
  } catch (error) {
    console.error('(PremisesForm) Error submitting rating:', error);
    res.status(500).send({ error: 'Failed to submit rating', details: error.message });
  }
});
// --- API 6.7: (Contractor) Approves/Accepts the Rating (Fixed for Double Click) ---
app.post('/api/premises-forms/:formId/accept-rating', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to accept this rating." });
    }

    const currentStatus = doc.data().status;

    if (currentStatus === 'LOCKED' || currentStatus === 'AUTO-APPROVED') {
      console.log(`(PremisesForm) Form ${formId} was already LOCKED. Ignoring duplicate request.`);
      return res.status(200).send({ message: 'Rating accepted. Form is now locked.' });
    }

    if (currentStatus !== 'SCORED') {
      return res.status(400).send({ error: `Cannot accept rating. Form status is ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'LOCKED',
      completedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(PremisesForm) Form ${formId} rating accepted and LOCKED.`);
    res.status(200).send({ message: 'Rating accepted. Form is now locked.' });

  } catch (error) {
    console.error('(PremisesForm) Error accepting rating:', error);
    res.status(500).send({ error: 'Failed to accept rating', details: error.message });
  }
});

// --- API 6.8: (Contractor) Re-submits Premises Form (UPDATED) ---
app.put('/api/premises-forms/:formId/resubmit', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const {
      location, manpower, signature,
      contractorRemarks,
      resubmitSign
    } = req.body;

    if (!resubmitSign) {
      return res.status(400).send({ error: "Resubmit Signature is required to submit the form again." });
    }

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to resubmit this form." });
    }

    if (doc.data().status !== 'REJECTED_BY_RAILWAY') {
      return res.status(400).send({ error: `Cannot resubmit form. Status is ${doc.data().status}` });
    }

    const updateData = {
      location: location || doc.data().location,
      manpower: manpower || doc.data().manpower,
      signature: signature || doc.data().signature,
      resubmitSignature: resubmitSign,
      status: 'RE-SUBMITTED',

      resubmittedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    if (contractorRemarks) {
      updateData.contractorRemarks = contractorRemarks;
    }

    await formDocRef.update(updateData);

    console.log(`(PremisesForm) Form ${formId} has been RESUBMITTED. Remarks preserved.`);
    res.status(200).send({ message: 'Form has been re-submitted to Railway Supervisor.' });

  } catch (error) {
    console.error('(PremisesForm) Error resubmitting form:', error);
    res.status(500).send({ error: 'Failed to resubmit form', details: error.message });
  }
});
// --- API 6.9: (Railway) Rejects a Form (PREMISES) ---
app.post('/api/premises-forms/:formId/reject', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const { rejectionComments } = req.body;
    const railwaySupervisor = req.user;

    if (!rejectionComments) {
      return res.status(400).send({ error: "Rejection comments are required." });
    }

    const formDocRef = db.collection('premisesForms').doc(formId);
    const doc = await formDocRef.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found.' });

    if (doc.data().submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to reject this form." });
    }

    const currentStatus = doc.data().status;
    if (currentStatus !== 'SUBMITTED' && currentStatus !== 'RE-SUBMITTED' && currentStatus !== 'APPROVED_BY_RAILWAY' && currentStatus !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot reject a form with status: ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'REJECTED_BY_RAILWAY',
      rejectionComments: rejectionComments,
      rejectedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(PremisesForm) Form ${formId} has been REJECTED.`);
    res.status(200).send({ message: 'Form successfully rejected.' });

  } catch (error) {
    console.error('(PremisesForm) Error rejecting form:', error);
    res.status(500).send({ error: 'Failed to reject form', details: error.message });
  }
});

app.get('/api/premises-forms/submitted', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, entityId: tokenEntityId, division: tokenDivision } = req.user;
    const { entityId, division, contractId, startDate, endDate } = req.query;

    let targetEntityId = entityId;

    if (userType === 'contractor') {
      targetEntityId = tokenEntityId;
    }
    if (!targetEntityId && userType !== 'railway') {
      return res.status(400).send({ error: "Entity ID is required." });
    }

    let query = db.collection('premisesForms');
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    query = query.where('submittedByEntityId', '==', targetEntityId);

    if (userType === 'contractor') {
      if (isMaster) {
        if (division) query = query.where('submittedByDivision', '==', division);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', tokenDivision);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }
    else if (userType === 'railway' && division) {
      query = query.where('submittedByDivision', '==', division);
    }

    if (contractId) query = query.where('contractId', '==', contractId);

    const snapshot = await query.orderBy('createdAt', 'desc').get();
    let formList = [];

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;
      if (start || end) {
        if (!data.formDateTime) {
          include = false;
        } else {
          const fDate = new Date(data.formDateTime);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }
      if (include) formList.push(data);
    });

    res.status(200).json({ count: formList.length, forms: formList });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index Missing', details: error.message });
    }
    console.error('(PremisesForm) Error:', error);
    res.status(500).send({ error: error.message });
  }
});
app.get('/api/premises-forms/pending-scoring', verifyToken, async (req, res) => {
  try {
    const railwaySupervisorId = req.user.uid;

    let query = db.collection('premisesForms');

    query = query.where('submittedTo.railwayEmployeeId', '==', railwaySupervisorId);
    query = query.where('status', 'in', ['APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS']);

    query = query.orderBy('createdAt', 'desc');
    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, forms: [] });
    }

    const formList = [];
    snapshot.forEach(doc => {
      formList.push(doc.data());
    });

    res.status(200).json({
      count: formList.length,
      forms: formList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ // <-- Typo Fixed
        error: 'Query requires an index.',
        details: `Firebase needs a composite index for this query. ${error.message}`
      });
    }
    console.error('(PremisesForm) Error fetching pending-scoring forms:', error);
    res.status(500).send({ error: 'Failed to fetch forms', details: error.message });
  }
});

// =======================================================
// == 9. REPORT DATA API (PREMISES - FINAL: Contract Check + All Logic)
// =======================================================

app.get('/api/reports/premises-data', verifyToken, async (req, res) => {
  try {
    const {
      zone, division, depot,
      startDate, endDate,
      areaType,
      contractorId,
      contractId,
      supervisorId
    } = req.query;

    console.log(`(Report) Fetching Premises Data. Contract: ${contractId}, Area: ${areaType || 'All'}`);

    if (contractId) {
      const contractDoc = await db.collection('contracts').doc(contractId).get();

      if (!contractDoc.exists) {
        return res.status(404).json({
          error: "Contract Not Found",
          details: "The provided Contract ID does not exist."
        });
      }

      const cData = contractDoc.data();

      const isStatusExpired = ['Expired', 'expired', 'Inactive', 'inactive'].includes(cData.status);

      let isDateExpired = false;
      if (cData.endDate) {
        const today = new Date();
        const end = new Date(cData.endDate);
        end.setHours(23, 59, 59, 999);
        if (today > end) isDateExpired = true;
      }

      if (isStatusExpired || isDateExpired) {
        return res.status(400).json({
          error: "Contract Expired",
          details: `This contract expired on ${cData.endDate || 'Unknown Date'}. You cannot view reports for expired contracts.`
        });
      }
    }
    const areaMap = {
      'GICC': 23530,
      'OWS': 8630,
      'NWS': 10130
    };

    let query = db.collection('premisesForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    if (zone) query = query.where('submittedByZone', '==', zone);
    if (division) query = query.where('submittedByDivision', '==', division);
    if (depot) query = query.where('submittedByDepot', '==', depot);

    if (contractorId) query = query.where('submittedByEntityId', '==', contractorId);

    if (contractId) query = query.where('contractId', '==', contractId);

    if (supervisorId) query = query.where('submittedById', '==', supervisorId);


    const snapshot = await query.get();
    let reportData = [];
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    const selectedAreas = areaType ? areaType.split(',').map(item => item.trim()) : [];

    snapshot.forEach(doc => {
      const data = doc.data();
      const formDate = data.formDateTime ? new Date(data.formDateTime) : new Date();

      let includeRecord = true;

      if (start && formDate < start) includeRecord = false;
      if (end && formDate > end) includeRecord = false;

      if (includeRecord && selectedAreas.length > 0) {
        if (!selectedAreas.includes(data.location)) {
          includeRecord = false;
        }
      }

      if (includeRecord) {
        const summary = data.ratingDetails?.summary || {};
        const totalAreaSqMtr = data.area || areaMap[data.location] || 0;

        const overallPct = parseFloat(summary.overallAveragePct) || 0;

        let pen8190 = 'NA', pen7180 = 'NA', pen70 = 'NA';
        if (overallPct >= 81 && overallPct <= 90) pen8190 = 'Yes';
        else if (overallPct >= 71 && overallPct <= 80) pen7180 = 'Yes';
        else if (overallPct <= 70) pen70 = 'Yes';

        reportData.push({
          date: formDate.toLocaleDateString('en-IN'),
          premiseName: data.location || 'Unknown',
          location: data.location || 'Unknown',

          areaCategory: totalAreaSqMtr,

          totalArea: totalAreaSqMtr,
          areaAttended: totalAreaSqMtr,
          areaNotAttended: 0,
          ratingInPct: overallPct.toFixed(2) + '%',

          housekeepingScore: summary.housekeepingAveragePct || '0%',
          pitLineScore: summary.pitLineAveragePct || '0%',
          garbageScore: summary.garbageDisposalAveragePct || '0%',
          overallScore: summary.overallAveragePct || '0%',

          above90: overallPct > 90 ? 'Yes' : 'NA',
          penalty81to90: pen8190,
          penalty71to80: pen7180,
          penaltyBelow70: pen70
        });
      }
    });

    res.status(200).json({
      count: reportData.length,
      data: reportData
    });

  } catch (error) {
    console.error('(Report) Error fetching premises data:', error);
    res.status(500).send({ error: 'Failed to fetch report data', details: error.message });
  }
});

// =======================================================
// == 13. COACH REPORT API (FINAL: Contract Expiry Check + All Filters)
// =======================================================

app.get('/api/reports/coach-data', verifyToken, async (req, res) => {
  try {
    const {
      zone, division, depot,
      startDate, endDate,
      contractorId,
      contractId,
      trainNo,
      coachNo,
      supervisorId,
      areaType
    } = req.query;

    console.log(`(Report) Fetching Coach Data. Contract: ${contractId}, Train: ${trainNo || 'All'}`);

    if (contractId) {
      const contractDoc = await db.collection('contracts').doc(contractId).get();

      if (!contractDoc.exists) {
        return res.status(404).json({
          error: "Contract Not Found",
          details: "The provided Contract ID does not exist."
        });
      }

      const cData = contractDoc.data();

      const isStatusExpired = ['Expired', 'expired', 'Inactive', 'inactive'].includes(cData.status);

      let isDateExpired = false;
      if (cData.endDate) {
        const today = new Date();
        const end = new Date(cData.endDate);
        end.setHours(23, 59, 59, 999);
        if (today > end) isDateExpired = true;
      }

      if (isStatusExpired || isDateExpired) {
        return res.status(400).json({
          error: "Contract Expired",
          details: `This contract expired on ${cData.endDate || 'Unknown Date'}. You cannot view reports for expired contracts.`
        });
      }
    }

    let query = db.collection('coachForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    if (zone) query = query.where('submittedByZone', '==', zone);
    if (division) query = query.where('submittedByDivision', '==', division);
    if (depot) query = query.where('submittedByDepot', '==', depot);

    if (contractorId) query = query.where('submittedByEntityId', '==', contractorId);

    if (contractId) query = query.where('contractId', '==', contractId);

    if (supervisorId) query = query.where('submittedById', '==', supervisorId);

    const snapshot = await query.get();

    let reportData = [];

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);
    const selectedAreas = areaType ? areaType.split(',').map(item => item.trim()) : [];

    snapshot.forEach(doc => {
      const data = doc.data();
      const formDate = data.formDateTime ? new Date(data.formDateTime) : new Date();

      let includeRecord = true;

      if (start && formDate < start) includeRecord = false;
      if (end && formDate > end) includeRecord = false;

      if (selectedAreas.length > 0) {
        if (!selectedAreas.includes(data.location)) {
          includeRecord = false;
        }
      }

      if (includeRecord && trainNo) {
        const dbTrainNo = String(data.trainNumber || "");
        if (!dbTrainNo.includes(trainNo)) {
          includeRecord = false;
        }
      }

      if (includeRecord && coachNo) {
        const coaches = data.ratingDetails?.coachEvaluationTable || [];
        const found = coaches.some(c => c.coachNumber && c.coachNumber.includes(coachNo));
        if (!found) includeRecord = false;
      }

      if (includeRecord) {
        const evalTable = data.ratingDetails?.coachEvaluationTable || [];
        const totalCoaches = evalTable.length;
        const workType = (data.ratingDetails?.workType || "N/A").toLowerCase();
        const acwpStatus = (data.ratingDetails?.acwpStatus || "Without ACWP").toLowerCase();
        const machinesList = data.machinesUsed || [];
        const isMachineUsed = machinesList.length > 0;

        let counts = {
          internal: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          rbpcMachine: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          rbpcManual: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          extAcwp: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          extManual: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          intensive: { A: 0, B: 0, C: 0, D: 0, NA: 0 },
          toiletries: { Yes: 0, No: 0, NA: 0 },
          watering: { Yes: 0, No: 0, NA: 0 },
          doors: { Yes: 0, No: 0, NA: 0 }
        };

        const inc = (obj, key) => { if (key && obj[key] !== undefined) obj[key]++; };

        evalTable.forEach(coach => {
          inc(counts.internal, coach.internalCleaning);
          inc(counts.intensive, coach.intensiveCleaning);
          inc(counts.toiletries, coach.toiletries);
          inc(counts.watering, coach.watering);
          inc(counts.doors, coach.doorsLocking);

          if (workType.includes("rbpc")) {
            const grade = coach.externalCleaning;
            if (isMachineUsed) inc(counts.rbpcMachine, grade);
            else inc(counts.rbpcManual, grade);
          } else {
            const grade = coach.externalCleaning;
            if (acwpStatus.includes("with acwp") || acwpStatus === "functional") inc(counts.extAcwp, grade);
            else inc(counts.extManual, grade);
          }
        });

        let requiredManpower = 0;
        const isPrimSec = workType.includes("primary") || workType.includes("secondary");
        const isRBPC = workType.includes("rbpc");
        const hasACWP = acwpStatus.includes("with acwp") || acwpStatus === "functional";

        if (isPrimSec) {
          requiredManpower += (totalCoaches * 0.45);
          requiredManpower += hasACWP ? (totalCoaches * 0.05) : (totalCoaches * 0.2);
        }
        if (isRBPC) requiredManpower += (totalCoaches * 0.3);

        const calculatedReq = Math.ceil(requiredManpower);
        const actualProvided = data.manpower ? data.manpower.length : 0;
        const shortage = Math.max(0, calculatedReq - actualProvided);

        let actualWithACWP = 'N/A';
        let actualWithoutACWP = 'N/A';

        if (hasACWP) {
          actualWithACWP = actualProvided;
        } else {
          actualWithoutACWP = actualProvided;
        }

        const fmt = (val) => (val === 0 ? '' : val);

        reportData.push({
          date: formDate.toLocaleDateString('en-IN'),
          trainName: data.trainName || 'N/A',
          trainNo: data.trainNumber || 'N/A',
          workType: data.ratingDetails?.workType,
          acwpStatus: data.ratingDetails?.acwpStatus,

          int_A: fmt(counts.internal.A), int_B: fmt(counts.internal.B), int_C: fmt(counts.internal.C), int_D: fmt(counts.internal.D), int_NA: fmt(counts.internal.NA),

          rbpc_mach_A: fmt(counts.rbpcMachine.A), rbpc_mach_B: fmt(counts.rbpcMachine.B), rbpc_mach_C: fmt(counts.rbpcMachine.C), rbpc_mach_D: fmt(counts.rbpcMachine.D), rbpc_mach_NA: fmt(counts.rbpcMachine.NA),
          rbpc_man_A: fmt(counts.rbpcManual.A), rbpc_man_B: fmt(counts.rbpcManual.B), rbpc_man_C: fmt(counts.rbpcManual.C), rbpc_man_D: fmt(counts.rbpcManual.D), rbpc_man_NA: fmt(counts.rbpcManual.NA),

          ext_acwp_A: fmt(counts.extAcwp.A), ext_acwp_B: fmt(counts.extAcwp.B), ext_acwp_C: fmt(counts.extAcwp.C), ext_acwp_D: fmt(counts.extAcwp.D), ext_acwp_NA: fmt(counts.extAcwp.NA),
          ext_man_A: fmt(counts.extManual.A), ext_man_B: fmt(counts.extManual.B), ext_man_C: fmt(counts.extManual.C), ext_man_D: fmt(counts.extManual.D), ext_man_NA: fmt(counts.extManual.NA),

          intense_A: fmt(counts.intensive.A), intense_B: fmt(counts.intensive.B), intense_C: fmt(counts.intensive.C), intense_D: fmt(counts.intensive.D), intense_NA: fmt(counts.intensive.NA),

          toil_Yes: fmt(counts.toiletries.Yes), toil_No: fmt(counts.toiletries.No), toil_NA: fmt(counts.toiletries.NA),
          water_Yes: fmt(counts.watering.Yes), water_No: fmt(counts.watering.No), water_NA: fmt(counts.watering.NA),
          door_Yes: fmt(counts.doors.Yes), door_No: fmt(counts.doors.No), door_NA: fmt(counts.doors.NA),

          actualWithACWP: actualWithACWP,
          actualWithoutACWP: actualWithoutACWP,
          manpowerShortage: shortage > 0 ? shortage : 'Nil',
          machineShortage: 'N/A'
        });
      }
    });

    res.status(200).json({
      count: reportData.length,
      data: reportData
    });

  } catch (error) {
    console.error('(Report) Error fetching coach report data:', error);
    res.status(500).send({ error: 'Failed to fetch report data', details: error.message });
  }
});
// --- API: Get Entities based on Zone & Division (Using CONTRACTS Collection) ---
app.get('/api/filter/entities', verifyToken, async (req, res) => {
  try {
    const { zone, division } = req.query;

    if (!zone || !division) {
      return res.status(400).send({ error: "Zone and Division are required." });
    }

    console.log(`(Filter) Searching Contracts in Zone: "${zone}", Division: "${division}"`);

    const contractsRef = db.collection('contracts');
    const snapshot = await contractsRef
      .where('zone', '==', zone)
      .where('division', '==', division)
      .where('status', '==', 'Active')
      .get();

    if (snapshot.empty) {
      console.log('(Filter) No active contracts found.');
      return res.status(200).json({ count: 0, entities: [] });
    }

    const entityIds = new Set();
    snapshot.forEach(doc => {
      const data = doc.data();
      if (data.entityId) {
        entityIds.add(data.entityId);
      }
    });

    if (entityIds.size === 0) {
      return res.status(200).json({ count: 0, entities: [] });
    }

    const entityList = [];
    const promises = Array.from(entityIds).map(async (id) => {
      const doc = await db.collection('entities').doc(id).get();
      if (doc.exists) {
        const data = doc.data();
        entityList.push({
          uid: data.uid,
          companyName: data.companyName,
          registrationType: data.registrationType
        });
      }
    });

    await Promise.all(promises);

    console.log(`(Filter) Found ${entityList.length} unique entities.`);

    res.status(200).json({
      count: entityList.length,
      entities: entityList
    });

  } catch (error) {
    console.error('(Filter) Error fetching entities:', error);

    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: `Create Index: contracts -> zone + division + status`
      });
    }
    res.status(500).send({ error: 'Failed to fetch entities', details: error.message });
  }
});
// --- API: Get Supervisors based on Entity & Division (For Report Filter Step 2) ---
app.get('/api/filter/supervisors', verifyToken, async (req, res) => {
  try {
    const { entityId, division, zone } = req.query;

    if (!entityId || !division) {
      return res.status(400).send({ error: "Entity ID and Division are required." });
    }

    console.log(`(Filter) Searching Supervisors for Entity: ${entityId} in ${division}`);

    const usersRef = db.collection('users');
    let query = usersRef
      .where('entityId', '==', entityId)
      .where('division', '==', division)
      .where('userType', '==', 'contractor')
      .where('status', '==', 'APPROVED');
    if (zone) {
      query = query.where('zone', '==', zone);
    }

    const snapshot = await query.get();

    if (snapshot.empty) {
      console.log('(Filter) No supervisors found.');
      return res.status(200).json({ count: 0, supervisors: [] });
    }

    const supervisorList = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      supervisorList.push({
        uid: data.uid,
        fullName: data.fullName,
        mobile: data.mobile,
        depot: data.depot
      });
    });

    console.log(`(Filter) Found ${supervisorList.length} supervisors.`);

    res.status(200).json({
      count: supervisorList.length,
      supervisors: supervisorList
    });

  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: `Firebase Index Required: users -> entityId + division + userType + status`
      });
    }
    console.error('(Filter) Error fetching supervisors:', error);
    res.status(500).send({ error: 'Failed to fetch supervisors', details: error.message });
  }
});
// =======================================================
// == 7. CRON JOBS (Auto-Approve Forms & Expire Contracts)
// =======================================================

const checkAndApprove = async (collectionName) => {
  try {
    const thirtyMinAgo = new Date(Date.now() - 30 * 60 * 1000);

    const snapshot = await db.collection(collectionName)
      .where('status', '==', 'SCORED')
      .get();

    if (snapshot.empty) return;

    const batch = db.batch();
    let approvedCount = 0;

    snapshot.forEach(doc => {
      const data = doc.data();
      let shouldApprove = false;

      if (data.ratedAt) {
        let ratedDate;
        if (typeof data.ratedAt.toDate === 'function') {
          ratedDate = data.ratedAt.toDate();
        } else {
          ratedDate = new Date(data.ratedAt);
        }

        if (!isNaN(ratedDate.getTime()) && ratedDate < thirtyMinAgo) {
          shouldApprove = true;
        }
      }

      if (shouldApprove) {
        batch.update(doc.ref, {
          status: 'AUTO-APPROVED',
          completedAt: admin.firestore.FieldValue.serverTimestamp(),
          autoApprovedAt: admin.firestore.FieldValue.serverTimestamp()
        });
        approvedCount++;
      }
    });

    if (approvedCount > 0) {
      await batch.commit();
      console.log(`[Cron] Auto-approved ${approvedCount} forms in ${collectionName}`);
    }

  } catch (e) {
    if (e.code !== 'FAILED_PRECONDITION') console.error(`[Cron Error] Form Approval:`, e.message);
  }
};

const checkContractExpiry = async () => {
  try {
    console.log(' Checking for expired contracts...');
    const now = new Date();

    const snapshot = await db.collection('contracts')
      .where('status', 'in', ['Active', 'active'])
      .get();

    if (snapshot.empty) {
      console.log(' No active contracts found to check.');
      return;
    }

    const batch = db.batch();
    let expiredCount = 0;

    snapshot.forEach(doc => {
      const data = doc.data();

      if (data.endDate) {
        const endDate = new Date(data.endDate);

        endDate.setHours(23, 59, 59, 999);

        if (now > endDate) {
          const docRef = db.collection('contracts').doc(doc.id);

          batch.update(docRef, {
            status: 'Expired',
            updatedAt: new Date()
          });
          expiredCount++;
        }
      }
    });

    if (expiredCount > 0) {
      await batch.commit();
      console.log(` Successfully Expired ${expiredCount} contracts.`);
    } else {
      console.log('All active contracts are still valid.');
    }

  } catch (e) {
    console.error(' [Cron Error] Contract Expiry:', e.message);
  }
};


cron.schedule('0 0 * * *', () => {
  console.log(' Running Midnight Cron Job...');
  checkContractExpiry();
});

cron.schedule('55 23 * * *', async () => {
  console.log('--- Starting Automated Daily Reports ---');
  try {
    await getDailyReportData();
    console.log('--- Finished Automated Daily Reports Successfully ---');
  } catch (error) {
    console.error('--- Automated Daily Reports Failed ---', error);
  }
});

// Evidence archive cron: runs daily at 2 AM
cron.schedule('0 2 * * *', async () => {
  console.log('--- Running Evidence Archive Cron ---');
  try {
    const result = await evidence.archiveOldRecords(db, bucket, 7);
    console.log(`Archived ${result.archived} records older than ${result.daysOld} days`);
    const longResult = await evidence.moveToLongTerm(db, 30);
    console.log(`Moved ${longResult.moved} records to long-term storage`);
  } catch (error) {
    console.error('Evidence archive cron failed:', error);
  }
});

setInterval(() => {
  checkAndApprove('coachForms');
  checkAndApprove('premisesForms');
  checkAndApprove('ctsForms');
  checkContractExpiry();
}, 600000);

// ─── Every 15 minutes: Update task statuses based on time ───
cron.schedule('*/15 * * * *', async () => {
  try {
    const now = new Date();
    const today = now.toISOString().split('T')[0];

    // Get active journey IDs
    const activeRunsSnap = await db.collection('RunInstance')
      .where('status', '==', 'Active')
      .select('runInstanceId')
      .get();
    const activeRunIds = new Set();
    activeRunsSnap.forEach(doc => activeRunIds.add(doc.id));

    // Update task_headers: mark OVERDUE if scheduled time has passed (only for active journeys)
    const headerSnapshot = await db.collection('task_headers')
      .where('status', 'in', ['PLANNED'])
      .where('scheduledDate', '==', today)
      .get();

    const headerBatch = db.batch();
    let updateCount = 0;
    headerSnapshot.forEach(doc => {
      const data = doc.data();
      // Skip headers for non-active journeys
      if (!activeRunIds.has(data.runInstanceId)) return;

      const [h, m] = (data.scheduledTime || '00:00').split(':').map(Number);
      const scheduledDate = new Date(`${today}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:00`);
      if (now > scheduledDate) {
        headerBatch.update(doc.ref, {
          status: 'OVERDUE',
          updatedAt: now.toISOString()
        });
        updateCount++;
      }
    });
    if (updateCount > 0) {
      await headerBatch.commit();
      console.log(`[Cron] Updated ${updateCount} task headers to OVERDUE`);
    }

    // Auto-escalate SLA complaints
    const complaintSnapshot = await db.collection('obhs_complaints')
      .where('status', '==', 'OPEN')
      .get();

    const complaintBatch = db.batch();
    let escalatedCount = 0;
    complaintSnapshot.forEach(doc => {
      const data = doc.data();
      const createdAt = new Date(data.createdAt);
      const elapsedMinutes = (now - createdAt) / (1000 * 60);
      const slaMap = {
        'Cleaning': 30,
        'Garbage': 20,
        'Water': 30,
        'Petty Repair': 60,
        'Toilet': 30,
        'Electrical': 60
      };
      const slaMinutes = slaMap[data.category] || 60;

      if (elapsedMinutes > slaMinutes && !data.slaEscalated) {
        complaintBatch.update(doc.ref, {
          slaEscalated: true,
          slaEscalatedAt: now.toISOString(),
          slaBreach: true,
          escalationLevel: 1,
          status: 'ESCALATED',
          updatedAt: now.toISOString()
        });
        escalatedCount++;
      }
    });
    if (escalatedCount > 0) {
      await complaintBatch.commit();
      console.log(`[Cron] Auto-escalated ${escalatedCount} SLA-breached complaints`);
    }
  } catch (error) {
    console.error('[Cron] Task status update error:', error);
  }
});

// ─── Every 5 minutes: Check task status transitions near scheduled times ───
cron.schedule('*/5 * * * *', async () => {
  try {
    const now = new Date();
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    // Only process tasks for ACTIVE journeys
    const activeRunsSnap = await db.collection('RunInstance')
      .where('status', '==', 'Active')
      .select('runInstanceId')
      .get();
    const activeRunIds = new Set();
    activeRunsSnap.forEach(doc => activeRunIds.add(doc.id));

    // Update PLANNED task_details to OPEN within 15min of scheduled time
    const detailSnapshot = await db.collection('task_details')
      .where('status', '==', 'PLANNED')
      .get();

    const detailBatch = db.batch();
    let openCount = 0;
    let dueSoonCount = 0;
    let skippedCount = 0;

    detailSnapshot.forEach(doc => {
      const data = doc.data();

      // Skip tasks belonging to non-active journeys
      if (!activeRunIds.has(data.runInstanceId)) {
        skippedCount++;
        return;
      }

      const [h, m] = (data.scheduledTime || '00:00').split(':').map(Number);
      const taskMinutes = h * 60 + m;
      const diff = taskMinutes - nowMinutes;

      if (diff <= 0 && diff > -120) {
        detailBatch.update(doc.ref, {
          status: 'OPEN',
          updatedAt: now.toISOString()
        });
        openCount++;
      } else if (diff > 0 && diff <= 15) {
        detailBatch.update(doc.ref, {
          status: 'DUE_SOON',
          updatedAt: now.toISOString()
        });
        dueSoonCount++;
      }
    });

    if (openCount + dueSoonCount > 0) {
      await detailBatch.commit();
      console.log(`[Cron] Task status updated: ${openCount} OPEN, ${dueSoonCount} DUE_SOON (${skippedCount} skipped - not active)`);
    }
  } catch (error) {
    console.error('[Cron] Task detail status error:', error);
  }
});

// ─── Every 5 minutes: Auto-escalate overdue tasks by time threshold ───
// Thresholds: 30min→CS, 45min→CTS, 60min→CA, 90min→CM
cron.schedule('*/5 * * * *', async () => {
  try {
    const now = new Date();
    const overdueTasksSnap = await db.collection('task_details')
      .where('status', '==', 'OVERDUE')
      .get();

    if (overdueTasksSnap.empty) return;

    const escalationChain = [
      { minutes: 90, escalateTo: 'CM', label: 'CM Alert' },
      { minutes: 60, escalateTo: 'CA', label: 'CA Alert' },
      { minutes: 45, escalateTo: 'CTS', label: 'CTS Alert' },
      { minutes: 30, escalateTo: 'CS', label: 'CS Alert' }
    ];

    const batch = db.batch();
    let escalationCount = 0;

    for (const doc of overdueTasksSnap.docs) {
      const task = doc.data();
      const updatedAt = task.updatedAt ? new Date(task.updatedAt) : null;
      if (!updatedAt) continue;

      const elapsedMinutes = (now - updatedAt) / (1000 * 60);
      const lastEscalation = task.lastEscalationAt || 0;

      for (const level of escalationChain) {
        if (elapsedMinutes >= level.minutes && lastEscalation < level.minutes) {
          const escRef = db.collection('escalations').doc();
          batch.set(escRef, {
            escalationId: escRef.id,
            sourceEntity: 'task_detail',
            sourceId: task.detailId || doc.id,
            reason: `${level.label}: Task overdue for ${Math.round(elapsedMinutes)} minutes`,
            details: `${task.taskType} on Coach ${task.coachNo} overdue since ${updatedAt.toISOString()}`,
            escalatedBy: 'system',
            escalatedByName: 'System',
            escalatedByRole: 'system',
            escalatedToRole: level.escalateTo,
            status: 'OPEN',
            createdAt: now.toISOString(),
            updatedAt: now.toISOString()
          });
          batch.update(doc.ref, {
            lastEscalationAt: level.minutes,
            escalationLevel: level.label,
            updatedAt: now.toISOString()
          });
          escalationCount++;
          break; // Only apply highest threshold met
        }
      }
    }

    if (escalationCount > 0) {
      await batch.commit();
      console.log(`[Cron] Auto-escalated ${escalationCount} overdue tasks`);
    }
  } catch (error) {
    console.error('[Cron] Escalation error:', error);
  }
});
// =======================================================
// == 16. COACH CLEANING STATS API (FINAL: GRADES + OPS BREAKDOWN)
// =======================================================

app.get('/api/reports/train-performance', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone, division, depot, entityId } = req.user;
    const { trainNo, startDate, endDate } = req.query;

    console.log(`(TrainReport) Request by: ${role} (${userType}) - Train: ${trainNo || 'All'}`);

    let query = db.collection('coachForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    if (userType === 'railway') {
      if (isMaster) {
        query = query.where('submittedByZone', '==', zone);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedTo.railwayEmployeeId', '==', uid);
      }
    }
    else if (userType === 'contractor') {
      if (!entityId) {
        return res.status(403).send({ error: "Contractor Entity ID missing in token." });
      }

      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) {
        query = query.where('submittedByZone', '==', zone);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }
    else {
      return res.status(403).send({ error: "Unknown User Type" });
    }

    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, trains: [] });
    }

    const trainStats = {};
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      const formDate = new Date(data.formDateTime);


      let include = true;
      if (start && formDate < start) include = false;
      if (end && formDate > end) include = false;

      const tName = data.trainName || "Unknown Train";
      if (trainNo && !tName.includes(trainNo)) include = false;

      if (include) {
        if (!trainStats[tName]) {
          trainStats[tName] = {
            trainName: tName,
            totalForms: 0,
            totalScore: 0,
            scoresList: []
          };
        }

        const summary = data.ratingDetails?.summary || {};
        const scoreStr = summary.overallAveragePct || "0";
        const score = parseFloat(scoreStr);

        trainStats[tName].totalForms++;
        trainStats[tName].totalScore += score;

        trainStats[tName].scoresList.push({
          date: formDate.toLocaleDateString('en-IN'),
          score: scoreStr,
          contractor: data.submittedByEntityName,
          supervisor: data.submittedByName,
          location: data.submittedByDepot
        });
      }
    });
    const finalResult = Object.values(trainStats).map(train => ({
      trainName: train.trainName,
      formsCount: train.totalForms,
      averageScore: (train.totalScore / train.totalForms).toFixed(2) + '%',
      details: train.scoresList
    }));

    res.status(200).json({
      count: finalResult.length,
      data: finalResult
    });

  } catch (error) {
    console.error('(TrainReport) Error:', error);
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: 'Check Firebase Console. Need Composite Index for Role+Zone/Div logic.'
      });
    }
    res.status(500).send({ error: error.message });
  }
});


// =======================================================
// == API: Dashboard User Stats (Includes Logged-In User)
// =======================================================
app.get('/api/dashboard/user-stats', verifyToken, async (req, res) => {
  try {
    const { uid: loggedInUid, role, zone: userZone, division: userDiv } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    const { selectedZone, selectedDivision, selectedDepot } = req.query;

    const userSnap = await db.collection('users').get();
    const stats = {
      totalRegistered: 0,
      approvedUsers: 0,
      pendingApproval: 0,
      draftUsers: 0,
      rejectedUsers: 0,
      railwayStaff: 0,
      contractorStaff: 0
    };

    userSnap.docs.forEach(doc => {
      const d = doc.data();
      const docId = doc.id;

      let isVisible = false;

      if (docId === loggedInUid || d.uid === loggedInUid) {
        isVisible = true;
      }
      else if (userRole.includes('company master') || userRole.includes('super admin')) {
        let match = true;
        if (selectedZone && d.zone !== selectedZone) match = false;
        if (selectedDivision && d.division !== selectedDivision) match = false;
        if (selectedDepot && d.depot !== selectedDepot) match = false;
        isVisible = match;
      }
      else if (userRole.includes('master')) {
        const belongsToMyZone = d.zone && userZone && d.zone.toString().trim() === userZone.toString().trim();
        if (belongsToMyZone) {
          let match = true;
          if (selectedDivision && d.division !== selectedDivision) match = false;
          if (selectedDepot && d.depot !== selectedDepot) match = false;
          isVisible = match;
        }
      }
      else if (userRole.includes('admin') || userRole.includes('supervisor')) {
        const belongsToMyDiv = d.division && userDiv && d.division.toString().trim() === userDiv.toString().trim();
        if (belongsToMyDiv) {
          let match = true;
          if (selectedDepot && d.depot !== selectedDepot) match = false;
          isVisible = match;
        }
      }

      if (isVisible) {
        const s = (d.status || '').toString().trim().toUpperCase();
        stats.totalRegistered++;

        if (s === 'APPROVED') stats.approvedUsers++;
        else if (s === 'PENDING') stats.pendingApproval++;
        else if (s === 'REJECTED') stats.rejectedUsers++;
        else if (s === 'DRAFT' || s === 'DRAFT_USER') stats.draftUsers++;
        if (d.userType === 'railway') stats.railwayStaff++;
        if (d.userType === 'contractor') stats.contractorStaff++;
      }
    });
    res.status(200).json({
      success: true,
      data: stats
    });

  } catch (error) {
    console.error("Dashboard Stats Error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});
// ==========================================
// DASHBOARD SYSTEM OVERVIEW STATS (TOKEN BASED)
// ==========================================
app.get("/api/stats/system-overview", async (req, res) => {
  try {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ success: false, error: "No token provided" });
    }

    const token = authHeader.split(' ')[1];

    let decoded;
    try {
      decoded = jwt.verify(token, process.env.JWT_SECRET);
    } catch (err) {
      return res.status(401).json({ success: false, error: "Invalid or expired token" });
    }

    const userId = decoded.uid;
    const role = decoded.role;
    const userDivision = decoded.division;
    const userEntityId = decoded.entityId;

    console.log(`--- Dashboard Stats Request ---`);
    console.log(`User: ${decoded.fullName} | Role: ${role} | Div: ${userDivision}`);

    const [
      divisionsSnap,
      depotsSnap,
      usersSnap,
      companiesSnap,
      contractsSnap,
      formsSnap
    ] = await Promise.all([
      db.collection("divisions").get(),
      db.collection("depots").get(),
      db.collection("users").get(),
      db.collection("companies").get(),
      db.collection("contracts").get(),
      db.collection("forms_processed").get()
    ]);

    const allUsers = usersSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    const allContracts = contractsSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    const allDepots = depotsSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    const allForms = formsSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    let stats = {
      divisions: 0,
      depots: 0,
      railwayEmployees: 0,
      contractorEmployees: 0,
      registeredEntities: 0,
      activeContracts: 0,
      totalFormProcessed: allForms.length || 0
    };


    const normalizedRole = role ? role.toLowerCase() : "";

    if (normalizedRole === 'admin') {
      stats.divisions = divisionsSnap.size;
      stats.depots = depotsSnap.size;
      stats.registeredEntities = companiesSnap.size;
      stats.railwayEmployees = allUsers.filter(u => u.userType === 'railway' || u.role === 'Railway Supervisor').length;
      stats.contractorEmployees = allUsers.filter(u => u.userType === 'contractor').length;
      stats.activeContracts = allContracts.filter(c => c.status === 'Active' || c.status === 'active').length;
    }

    else if (normalizedRole.includes('railway') || normalizedRole.includes('supervisor')) {
      const assignedDiv = userDivision;

      stats.divisions = assignedDiv ? 1 : 0;
      stats.depots = allDepots.filter(d => d.division === assignedDiv).length;
      stats.railwayEmployees = allUsers.filter(u => u.division === assignedDiv && (u.userType === 'railway' || u.role === 'Railway Supervisor')).length;
      stats.contractorEmployees = allUsers.filter(u => u.division === assignedDiv && u.userType === 'contractor').length;
      stats.activeContracts = allContracts.filter(c => c.division === assignedDiv && (c.status === 'Active' || c.status === 'active')).length;
      stats.totalFormProcessed = allForms.filter(f => f.division === assignedDiv).length;
      stats.registeredEntities = [...new Set(allContracts.filter(c => c.division === assignedDiv).map(c => c.entityId))].length;
    }

    else if (normalizedRole.includes('company')) {
      const entityId = userEntityId;

      stats.registeredEntities = 1;
      stats.contractorEmployees = allUsers.filter(u => u.entityId === entityId).length;
      stats.activeContracts = allContracts.filter(c => c.entityId === entityId && (c.status === 'Active' || c.status === 'active')).length;
      stats.totalFormProcessed = allForms.filter(f => f.entityId === entityId).length;

      const myContracts = allContracts.filter(c => c.entityId === entityId);
      stats.divisions = [...new Set(myContracts.map(c => c.division))].length;
      stats.depots = [...new Set(myContracts.map(c => c.depot))].length;
    }

    console.log("Response Stats:", stats);
    res.json({ success: true, data: stats });

  } catch (error) {
    console.error("Dashboard Stats Error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});
// =======================================================
// == 17. DEDICATED TRAIN STATS API FOR DASHBOARD
// =======================================================
app.get('/api/dashboard/train-stats', verifyToken, async (req, res) => {
  try {
    const { role, zone: userZone, division: userDiv } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    const { selectedZone, selectedDivision, selectedDepot } = req.query;

    const trainSnap = await db.collection('trains').get();

    const stats = {
      totalTrains: 0,
      activeTrains: 0,
      inactiveTrains: 0,
      draftTrains: 0
    };

    trainSnap.docs.forEach(doc => {
      const d = doc.data();
      let isVisible = false;

      if (userRole.includes('company master') || userRole.includes('super admin')) {
        let match = true;
        if (selectedZone && d.zone !== selectedZone) match = false;
        if (selectedDivision && d.division !== selectedDivision) match = false;
        if (selectedDepot && d.depot !== selectedDepot) match = false;
        isVisible = match;
      }

      else if (userRole.includes('master')) {
        const isMyZone = d.zone && userZone && d.zone.toString().trim() === userZone.toString().trim();
        if (isMyZone) {
          let match = true;
          if (selectedDivision && d.division !== selectedDivision) match = false;
          if (selectedDepot && d.depot !== selectedDepot) match = false;
          isVisible = match;
        }
      }
      else if (userRole.includes('admin') || userRole.includes('supervisor')) {
        const isMyDiv = d.division && userDiv && d.division.toString().trim() === userDiv.toString().trim();
        if (isMyDiv) {
          let match = true;
          if (selectedDepot && d.depot !== selectedDepot) match = false;
          isVisible = match;
        }
      }

      if (isVisible) {
        const s = (d.status || '').toString().trim().toUpperCase();

        stats.totalTrains++;

        if (s === 'ACTIVE') {
          stats.activeTrains++;
        } else if (s === 'INACTIVE') {
          stats.inactiveTrains++;
        } else if (s === 'DRAFT') {
          stats.draftTrains++;
        }
      }
    });

    res.status(200).json({
      success: true,
      role: userRole,
      appliedFilters: {
        zone: selectedZone || (userRole.includes('master') ? userZone : "All"),
        division: selectedDivision || (userRole.includes('admin') ? userDiv : "All"),
        depot: selectedDepot || "All"
      },
      data: stats
    });

  } catch (error) {
    console.error('Train Stats API Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});


app.get('/api/dashboard/stats', verifyToken, async (req, res) => {
  try {
    const { role, userType, zone: userZone, division: userDiv } = req.user;
    const userRole = (role || '').trim().toLowerCase();

    const [userSnap, entitySnap, trainSnap, contractSnap] = await Promise.all([
      db.collection('users').get(),
      db.collection('entities').get(),
      db.collection('trains').get(),
      db.collection('contracts').get()
    ]);

    const stats = {
      user: { total: 0, approved: 0, pending: 0, draft: 0, railway: 0, contractor: 0 },
      entity: { total: 0, approved: 0, pending: 0, draft: 0 },
      train: { total: 0, active: 0, inactive: 0, draft: 0 }
    };

    userSnap.docs.forEach(doc => {
      const d = doc.data();

      let isVisible = false;
      if (userRole.includes('master')) {
        if (d.zone === userZone) isVisible = true;
      } else if (userRole.includes('admin') || userRole.includes('supervisor')) {
        if (d.division === userDiv) isVisible = true;
      } else if (userRole.includes('super admin')) {
        isVisible = true;
      }

      if (isVisible) {
        const status = d.status;

        stats.user.total++;

        if (status === 'APPROVED') stats.user.approved++;
        else if (status === 'PENDING') stats.user.pending++;
        else if (status === 'DRAFT') stats.user.draft++;

        if (d.userType === 'railway') stats.user.railway++;
        if (d.userType === 'contractor') stats.user.contractor++;
      }
    });

    entitySnap.docs.forEach(doc => {
      const d = doc.data();
      stats.entity.total++;
      if (d.status === 'APPROVED') stats.entity.approved++;
      if (d.status === 'PENDING') stats.entity.pending++;
    });

    trainSnap.docs.forEach(doc => {
      const d = doc.data();
      let isTrainVisible = false;
      if (userRole.includes('master') && d.zone === userZone) isTrainVisible = true;
      else if ((userRole.includes('admin') || userRole.includes('supervisor')) && d.division === userDiv) isTrainVisible = true;

      if (isTrainVisible) {
        stats.train.total++;
        if (d.status === 'ACTIVE' || d.status === 'active') stats.train.active++;
        else stats.train.inactive++;
      }
    });

    res.status(200).json({
      systemOverview: {
        railwayEmployees: stats.user.railway,
        contractorEmployees: stats.user.contractor,
        totalRegisteredEntities: stats.entity.total,
        activeContracts: contractSnap.docs.filter(c => c.data().status === 'ACTIVE').length,
        totalFormsProcessed: 0
      },
      userOverview: stats.user,
      entityOverview: stats.entity,
      trainOverview: stats.train
    });

  } catch (error) {
    console.error('Sync Error:', error);
    res.status(500).send({ error: error.message });
  }
});

// =======================================================
// == 14. TRAIN WISE PERFORMANCE REPORT API (Updated)
// =======================================================

app.get('/api/reports/train-performance', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone, division, depot, entityId } = req.user;
    const { trainNo, startDate, endDate } = req.query;

    console.log(`(TrainReport) Request by: ${role} (${userType}) - Train: ${trainNo || 'All'}`);

    let query = db.collection('coachForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master') || userRole === 'company master';
    const isAdmin = userRole.includes('admin') || userRole === 'company admin';

    if (userType === 'railway') {
      if (isMaster) {
        if (zone) query = query.where('submittedByZone', '==', zone);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedTo.railwayEmployeeId', '==', uid);
      }
    }
    else if (userType === 'contractor') {
      if (!entityId) {
        return res.status(403).send({ error: "Contractor Entity ID missing in token." });
      }
      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) {
        console.log('Contractor Master: Accessing All Divisions for their Company');
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
        console.log(`Contractor Admin: Accessing Division ${division}`);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }
    else {
      return res.status(403).send({ error: "Unknown User Type" });
    }


    const snapshot = await query.get();

    if (snapshot.empty) {
      return res.status(200).json({ count: 0, trains: [] });
    }

    const trainStats = {};
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      const formDate = new Date(data.formDateTime);

      let include = true;
      if (start && formDate < start) include = false;
      if (end && formDate > end) include = false;

      const tName = data.trainName || "Unknown Train";
      if (trainNo && !tName.includes(trainNo)) include = false;

      if (include) {
        if (!trainStats[tName]) {
          trainStats[tName] = {
            trainName: tName,
            totalForms: 0,
            totalScore: 0,
            scoresList: []
          };
        }

        const summary = data.ratingDetails?.summary || {};
        const scoreStr = summary.overallAveragePct || "0";
        const score = parseFloat(scoreStr);

        trainStats[tName].totalForms++;
        trainStats[tName].totalScore += score;

        trainStats[tName].scoresList.push({
          date: formDate.toLocaleDateString('en-IN'),
          score: scoreStr,
          contractor: data.submittedByEntityName,
          supervisor: data.submittedByName,
          location: data.submittedByDepot
        });
      }
    });

    const finalResult = Object.values(trainStats).map(train => ({
      trainName: train.trainName,
      formsCount: train.totalForms,
      averageScore: (train.totalScore / train.totalForms).toFixed(2) + '%',
      details: train.scoresList
    }));

    res.status(200).json({
      count: finalResult.length,
      data: finalResult
    });

  } catch (error) {
    console.error('(TrainReport) Error:', error);
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({
        error: 'Query requires an index.',
        details: 'Check Firebase Console. Need Composite Index for Role+Entity+Division logic.'
      });
    }
    res.status(500).send({ error: error.message });
  }
});


// =======================================================
// == 16. COACH CLEANING STATS API (FINAL: GRADES + OPS BREAKDOWN)
// =======================================================

app.get('/api/reports/coach-stats', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, entityId, division } = req.user;
    const { startDate, endDate, contractId } = req.query;

    console.log(`(CoachStats) Fetching Detailed Breakdown...`);

    let query = db.collection('coachForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    if (userType === 'contractor') {
      if (!entityId) return res.status(403).json({ error: "Entity ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) { /* Master sees all */ }
      else if (isAdmin) query = query.where('submittedByDivision', '==', division);
      else query = query.where('submittedById', '==', uid);
    } else {
      if (isAdmin) query = query.where('submittedByDivision', '==', division);
    }

    if (contractId) query = query.where('contractId', '==', contractId);

    const snapshot = await query.get();

    let stats = {
      totalTrains: 0,
      totalCoaches: 0,
      totalManpower: 0,
      totalPenalty: 0,

      grades: { A: 0, B: 0, C: 0, D: 0 },

      operations: {
        toiletries: { Yes: 0, No: 0, NA: 0 },
        doors: { Yes: 0, No: 0, NA: 0 },
        watering: { Yes: 0, No: 0, NA: 0 }
      },

      resources: {
        manpowerDeployed: 0,
        machinesUsed: 0
      }
    };

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;

      if (start || end) {
        if (!data.formDateTime) include = false;
        else {
          const fDate = new Date(data.formDateTime);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }

      if (include) {
        stats.totalTrains++;

        const mpCount = data.manpower ? data.manpower.length : 0;
        stats.totalManpower += mpCount;
        stats.resources.manpowerDeployed += mpCount;

        const machines = data.machinesUsed || [];
        stats.resources.machinesUsed += machines.length;

        if (data.penaltyAmount) {
          stats.totalPenalty += parseFloat(data.penaltyAmount);
        }
        const summary = data.ratingDetails?.summary || {};
        stats.totalCoaches += (summary.totalCoaches || 0);

        const categories = ['internal', 'external', 'intensive'];
        categories.forEach(cat => {
          if (summary[cat]) {
            stats.grades.A += (summary[cat].A || 0);
            stats.grades.B += (summary[cat].B || 0);
            stats.grades.C += (summary[cat].C || 0);
            stats.grades.D += (summary[cat].D || 0);
          }
        });

        const addOpStats = (target, source) => {
          if (!source) return;
          target.Yes += (source.Yes || 0) + (source.Ok || 0) + (source.Done || 0);
          target.No += (source.No || 0);
          target.NA += (source.NA || 0);
        };

        addOpStats(stats.operations.doors, summary.doorsLocking);
        addOpStats(stats.operations.toiletries, summary.toiletries);
        addOpStats(stats.operations.watering, summary.watering);
      }
    });
    const totalGrades = stats.grades.A + stats.grades.B + stats.grades.C + stats.grades.D;
    const safeTotal = totalGrades || 1;

    const responseData = {
      cards: {
        totalTrains: stats.totalTrains,
        totalCoaches: stats.totalCoaches,
        totalManpower: stats.totalManpower,
        totalPenalty: stats.totalPenalty
      },
      gradeDistribution: {
        A: { count: stats.grades.A, pct: ((stats.grades.A / safeTotal) * 100).toFixed(1) },
        B: { count: stats.grades.B, pct: ((stats.grades.B / safeTotal) * 100).toFixed(1) },
        C: { count: stats.grades.C, pct: ((stats.grades.C / safeTotal) * 100).toFixed(1) },
        D: { count: stats.grades.D, pct: ((stats.grades.D / safeTotal) * 100).toFixed(1) }
      },
      operations: {
        toiletries: stats.operations.toiletries,
        doors: stats.operations.doors,
        watering: stats.operations.watering
      },
      resources: stats.resources
    };

    console.log(`(CoachStats) Success. Operations breakdown included.`);
    res.status(200).json(responseData);

  } catch (error) {
    console.error('(CoachStats) Error:', error);
    res.status(500).send({ error: error.message });
  }
});

// =======================================================
// == 17. PREMISES CLEANING STATS API (FINAL: CARDS + QUALITY + MANPOWER)
// =======================================================

app.get('/api/reports/premises-stats', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, entityId, division } = req.user;
    const { startDate, endDate, contractId } = req.query;

    console.log(`(PremisesStats) Fetching Aggregated Counts...`);
    let query = db.collection('premisesForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    if (userType === 'contractor') {
      if (!entityId) return res.status(403).json({ error: "Entity ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) { /* Master sees all */ }
      else if (isAdmin) query = query.where('submittedByDivision', '==', division);
      else query = query.where('submittedById', '==', uid);
    } else {
      if (isAdmin) query = query.where('submittedByDivision', '==', division);
    }

    if (contractId) query = query.where('contractId', '==', contractId);

    const snapshot = await query.get();

    const areaMap = {
      'GICC': 23530,
      'OWS': 8630,
      'NWS': 10130
    };

    let stats = {
      totalForms: 0,
      totalAreaCleaned: 0,
      totalManpower: 0,

      quality: {
        above90: 0,
        range81to90: 0,
        range71to80: 0,
        below70: 0
      },

      uniqueDates: new Set()
    };
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;

      if (start || end) {
        if (!data.formDateTime) include = false;
        else {
          const fDate = new Date(data.formDateTime);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }

      if (include) {
        stats.totalForms++;

        const area = parseFloat(data.area) || areaMap[data.location] || 0;
        stats.totalAreaCleaned += area;

        const mpCount = data.manpower ? data.manpower.length : 0;
        stats.totalManpower += mpCount;

        if (data.formDateTime) {
          const dateStr = new Date(data.formDateTime).toDateString();
          stats.uniqueDates.add(dateStr);
        }

        const pct = parseFloat(data.ratingDetails?.summary?.overallAveragePct || 0);

        if (pct > 90) stats.quality.above90++;
        else if (pct >= 81) stats.quality.range81to90++;
        else if (pct >= 71) stats.quality.range71to80++;
        else stats.quality.below70++;
      }
    });

    const totalFormsSafe = stats.totalForms || 1;
    const totalDays = stats.uniqueDates.size || 1;

    const dailyAvgManpower = Math.round(stats.totalManpower / totalDays);

    const responseData = {
      cards: {
        totalPremisesCleaned: stats.totalForms,
        totalAreaCleaned: stats.totalAreaCleaned.toFixed(2),
        totalAreaUncleaned: 0,
        manpowerDeployed: stats.totalManpower
      },
      qualityPerformanceDistribution: {
        above90: {
          count: stats.quality.above90,
          pct: ((stats.quality.above90 / totalFormsSafe) * 100).toFixed(1)
        },
        range81to90: {
          count: stats.quality.range81to90,
          pct: ((stats.quality.range81to90 / totalFormsSafe) * 100).toFixed(1)
        },
        range71to80: {
          count: stats.quality.range71to80,
          pct: ((stats.quality.range71to80 / totalFormsSafe) * 100).toFixed(1)
        },
        below70: {
          count: stats.quality.below70,
          pct: ((stats.quality.below70 / totalFormsSafe) * 100).toFixed(1)
        }
      },
      manpowerAllocation: {
        dailyAverage: dailyAvgManpower,
        peakHours: 0
      }
    };

    console.log(`(PremisesStats) Success. Processed ${stats.totalForms} forms.`);
    res.status(200).json(responseData);

  } catch (error) {
    console.error('(PremisesStats) Error:', error);
    res.status(500).send({ error: error.message });
  }
});

// =======================================================
// == API 5.3: (Contractor) Submits a NEW CTS Form (Strict Logic)
// =======================================================
app.post('/api/cts-forms', verifyToken, async (req, res) => {
  try {
    const contractorSupervisor = req.user;

    const {
      trainId, formDateTime, platform, actArrival, actDeparture,
      workStart, workEnd, allowedWindow, lateYN, coachesInRake, coachesAttended,
      attendanceStaff, garbageDisposed, nominatedLocation, occupiedToilets, notes,
      submittedTo, signature, contractId
    } = req.body;

    if (!trainId || !contractId || !formDateTime) {
      return res.status(400).send({ error: "Missing mandatory fields (Train or Contract)." });
    }
    const [contractDoc, trainDoc, userDoc] = await Promise.all([
      db.collection('contracts').doc(contractId).get(),
      db.collection('trains').doc(trainId).get(),
      submittedTo?.railwayEmployeeId ? db.collection('users').doc(submittedTo.railwayEmployeeId).get() : Promise.resolve(null)
    ]);

    if (!contractDoc.exists) return res.status(404).send({ error: "Contract not found." });
    if (!trainDoc.exists) return res.status(404).send({ error: "Train not found." });

    const contractData = contractDoc.data();
    const trainData = trainDoc.data();
    const railwayEmployeeName = userDoc?.exists ? (userDoc.data().fullName || userDoc.data().name || "Unknown") : "N/A";

    const stationName = contractorSupervisor.division || "Unknown Station";
    const agreementNo = contractData.contractNumber || "N/A";

    const agreementDate = contractData.startDate || "N/A";
    const contractorName = contractData.agencyName || contractData.entityName || "N/A";
    const jobDate = new Date().toISOString().split('T')[0];

    const applicableFor = trainData.TrainApplicableFor || [];
    if (!applicableFor.includes('CTS')) {
      return res.status(400).send({ error: "This Train is not apllicable for CTS service." });
    }

    const newFormId = await generateFormId('cts', contractorSupervisor.division);

    const ctsFormData = {
      uid: newFormId,
      formId: newFormId,
      contractId: contractId,
      station: stationName,
      agreementNo: agreementNo,
      agreementDate: agreementDate,
      contractorName: contractorName,
      jobDate: jobDate,
      trainId,
      trainNumber: trainData.trainNo || "",
      trainName: trainData.trainName || "",
      formDateTime,
      actArrival,
      actDeparture,
      workStart,
      workEnd,
      platform: platform || null,
      allowedWindow: allowedWindow || null,
      lateYN: lateYN || "No",
      coachesInRake: coachesInRake || 0,
      coachesAttended: coachesAttended || 0,
      attendanceStaff: attendanceStaff || [],
      garbageDisposed: garbageDisposed || false,
      nominatedLocation: nominatedLocation || null,
      occupiedToilets: occupiedToilets || 0,
      notes: notes || "",

      submittedTo: {
        railwayEmployeeId: submittedTo?.railwayEmployeeId || null,
        railwayEmployeeName: railwayEmployeeName,
        division: submittedTo?.division || contractorSupervisor.division,
        depot: submittedTo?.depot || contractorSupervisor.depot
      },
      signature: {
        name: signature?.name || null,
        date: signature?.date || new Date().toISOString()
      },
      status: 'SUBMITTED',
      submittedById: contractorSupervisor.uid,
      submittedByName: contractorSupervisor.fullName || contractorSupervisor.name,
      submittedByZone: contractorSupervisor.zone || null,
      submittedByDivision: contractorSupervisor.division || null,
      submittedByDepot: contractorSupervisor.depot || null,
      submittedByEntityId: contractorSupervisor.entityId || null,
      submittedByEntityName: contractorSupervisor.entityName || contractorName,

      createdAt: admin.firestore.FieldValue.serverTimestamp()
    };

    await db.collection('ctsForms').doc(newFormId).set(ctsFormData);

    res.status(201).send({
      message: 'CTS Form submitted successfully.',
      uid: newFormId,
      railwayEmployeeName: railwayEmployeeName
    });

  } catch (error) {
    console.error('(CTS-Form) Error:', error);
    res.status(500).send({ error: 'Internal Server Error', details: error.message });
  }
});

// =======================================================
// == API 5.4: Get CTS Forms (STRICT COACH LOGIC)
// =======================================================
app.get('/api/cts-forms', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, zone, division, entityId } = req.user;
    const { status, type } = req.query;

    let query = db.collection('ctsForms');
    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master') || userRole === 'company master';
    const isAdmin = userRole.includes('admin');

    if (userType === 'railway') {
      if (isMaster) {
        if (zone) {
          query = query.where('submittedByZone', '==', zone);
          console.log(`(CTS-Form) Railway Master Locked to Zone: ${zone}`);
        } else {
          console.log('(CTS-Form) Railway Master (No Zone in Token) - Showing All');
        }
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedTo.railwayEmployeeId', '==', uid);
      }
    }
    else if (userType === 'contractor') {
      if (!entityId) return res.status(403).send({ error: "Company ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (isMaster) {
        console.log(`(CTS-Form) Contractor Master: Viewing All Company Data`);
      }
      else if (isAdmin) {
        query = query.where('submittedByDivision', '==', division);
      }
      else {
        query = query.where('submittedById', '==', uid);
      }
    }

    if (status) {
      query = query.where('status', '==', status);
    }
    else if (type === 'history') {
      query = query.where('status', 'in', ['SCORED', 'LOCKED', 'AUTO-APPROVED', 'REJECTED_BY_RAILWAY']);
    }
    else {
      query = query.where('status', 'in', ['SUBMITTED', 'RE-SUBMITTED', 'APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS']);
      console.log(`(CTS-Form) Active View for ${role}`);
    }

    const snapshot = await query.get();

    const list = [];
    snapshot.forEach(doc => {
      list.push({ id: doc.id, ...doc.data() });
    });

    // Sort in memory by createdAt descending to avoid requiring composite indexes
    list.sort((a, b) => {
      const dateA = a.createdAt ? (a.createdAt.toDate ? a.createdAt.toDate() : new Date(a.createdAt)) : new Date(0);
      const dateB = b.createdAt ? (b.createdAt.toDate ? b.createdAt.toDate() : new Date(b.createdAt)) : new Date(0);
      return dateB - dateA;
    });

    console.log(`(CTS-Form) Total Found: ${list.length}`);
    res.status(200).json({ count: list.length, forms: list });

  } catch (e) {
    console.error('(CTS-Get) Error:', e);
    if (e.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index Missing. Check Firebase Console.' });
    }
    res.status(500).send({ error: e.message });
  }
});


// =======================================================
// == API 5.5: (Railway) Approves CTS Manpower (Unlocks Scoring)
// =======================================================
app.post('/api/cts-forms/:formId/approve-manpower', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const railwaySupervisor = req.user;
    const formDocRef = db.collection('ctsForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) {
      return res.status(404).send({ error: 'CTS Form not found.' });
    }

    const formData = doc.data();
    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to approve this form." });
    }
    if (formData.status === 'SUBMITTED' || formData.status === 'RE-SUBMITTED') {
      await formDocRef.update({
        status: 'APPROVED_BY_RAILWAY',
        manpowerApprovedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      console.log(`(CTS-Form) Form ${formId} manpower approved by ${railwaySupervisor.fullName}.`);
      return res.status(200).send({ message: 'CTS Form approved for scoring.' });
    }
    res.status(400).send({ message: `Form status is already ${formData.status}.` });

  } catch (error) {
    console.error('(CTS-Form) Error approving manpower:', error);
    res.status(500).send({ error: 'Failed to approve manpower', details: error.message });
  }
});

// =======================================================
// == API 5.6: (Railway) Rejects CTS Form
// =======================================================
app.post('/api/cts-forms/:formId/reject', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const { rejectionComments } = req.body;
    const railwaySupervisor = req.user;

    if (!rejectionComments) {
      return res.status(400).send({ error: "Rejection comments are required." });
    }

    const formDocRef = db.collection('ctsForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: 'CTS Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to reject this form." });
    }

    const currentStatus = formData.status;
    const allowedStatuses = ['SUBMITTED', 'RE-SUBMITTED', 'APPROVED_BY_RAILWAY', 'SCORING_IN_PROGRESS'];

    if (!allowedStatuses.includes(currentStatus)) {
      return res.status(400).send({ error: `Cannot reject a form with status: ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'REJECTED_BY_RAILWAY',
      rejectionComments: rejectionComments,
      rejectedAt: admin.firestore.FieldValue.serverTimestamp(),
      rejectedByName: railwaySupervisor.fullName || railwaySupervisor.name
    });

    console.log(`(CTS-Form) Form ${formId} has been REJECTED by ${railwaySupervisor.fullName}.`);
    res.status(200).send({ message: 'Form successfully rejected.' });

  } catch (error) {
    console.error('(CTS-Form) Error rejecting form:', error);
    res.status(500).send({ error: 'Failed to reject form', details: error.message });
  }
});

// =======================================================
// == API 5.7: (Railway) Submits Final CTS Scoring (Chemical Details Update)
// =======================================================
app.post('/api/cts-forms/:formId/submit-scoring', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const {
      inspectionHeader,
      coachEvaluationTable,
      machinesUsed,
      chemicals,
      railwaySignatureName,
      railwaySignatureDate
    } = req.body;

    const railwaySupervisor = req.user;

    if (!inspectionHeader || !coachEvaluationTable || !railwaySignatureName || !chemicals) {
      return res.status(400).send({ error: "Inspection Header, Evaluation Table, Chemicals, and Signature are required." });
    }

    const formDocRef = db.collection('ctsForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: 'CTS Form not found.' });

    const formData = doc.data();

    if (formData.submittedTo.railwayEmployeeId !== railwaySupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to score this form." });
    }

    if (formData.status !== 'APPROVED_BY_RAILWAY' && formData.status !== 'SCORING_IN_PROGRESS') {
      return res.status(400).send({ error: `Cannot score form. Status is ${formData.status}.` });
    }

    let totalAllCoachesScore = 0;
    const processedEvaluation = coachEvaluationTable.map(coach => {
      const coachTotal = (Number(coach.jetCleaningScore) || 0) +
        (Number(coach.basinCleaningScore) || 0) +
        (Number(coach.disposalScore) || 0);

      totalAllCoachesScore += coachTotal;

      let coachGrade = 'D';
      if (coachTotal >= 8) coachGrade = 'A';
      else if (coachTotal >= 6) coachGrade = 'B';
      else if (coachTotal >= 4) coachGrade = 'C';

      return { ...coach, totalScore: coachTotal, grade: coachGrade };
    });

    const averageScore = processedEvaluation.length > 0 ? (totalAllCoachesScore / processedEvaluation.length).toFixed(2) : 0;

    let overallGrade = 'Fail';
    if (averageScore >= 8) overallGrade = 'A';
    else if (averageScore >= 6) overallGrade = 'B';
    else if (averageScore >= 4) overallGrade = 'C';
    else overallGrade = 'D';

    await formDocRef.update({
      status: 'SCORED',
      scoringInProgress: false,
      ratingDetails: {
        inspectionHeader: inspectionHeader,
        coachEvaluationTable: processedEvaluation,
        machinesUsed: Array.isArray(machinesUsed) ? machinesUsed : [],
        chemicals: Array.isArray(chemicals) ? chemicals : [],
        totalPenalty: 0,
        summary: {
          averageScore: Number(averageScore),
          overallGrade: overallGrade
        }
      },
      railwaySignature: {
        name: railwaySignatureName || railwaySupervisor.fullName,
        date: railwaySignatureDate || new Date().toISOString()
      },
      ratedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(CTS-Form) Form ${formId} has been SCORED with Avg: ${averageScore}.`);
    res.status(200).send({
      message: 'CTS Scoring submitted successfully with chemical details.',
      averageScore,
      overallGrade
    });

  } catch (error) {
    console.error('(CTS-Scoring) Error:', error);
    res.status(500).send({ error: 'Failed to submit scoring', details: error.message });
  }
});

// =======================================================
// == API 5.8: (Contractor) Approves/Accepts the Rating (CTS LOCKED)
// =======================================================
app.post('/api/cts-forms/:formId/accept-rating', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const formDocRef = db.collection('ctsForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: 'CTS Form not found.' });

    const formData = doc.data();

    if (formData.submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to accept this rating." });
    }

    const currentStatus = formData.status;

    if (currentStatus === 'LOCKED' || currentStatus === 'AUTO-APPROVED') {
      return res.status(400).send({ error: "This form is already locked." });
    }

    if (currentStatus !== 'SCORED') {
      return res.status(400).send({ error: `Cannot accept rating. Form status is ${currentStatus}` });
    }

    await formDocRef.update({
      status: 'LOCKED',
      contractorAcceptedAt: admin.firestore.FieldValue.serverTimestamp(),
      completedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`(CTS-Form) Form ${formId} rating accepted and LOCKED by contractor.`);
    res.status(200).send({ message: 'Rating accepted. CTS Form is now locked.' });

  } catch (error) {
    console.error('(CTS-Form) Error accepting rating:', error);
    res.status(500).send({ error: 'Failed to accept rating', details: error.message });
  }
});

// =======================================================
// == API 5.9: (Contractor) Re-submits the CTS Form
// =======================================================
app.put('/api/cts-forms/:formId/resubmit', verifyToken, async (req, res) => {
  try {
    const { formId } = req.params;
    const contractorSupervisor = req.user;

    const {
      trainId, formDateTime, platform, actArrival, actDeparture,
      workStart, workEnd, allowedWindow, lateYN, coachesInRake, coachesAttended,
      attendanceStaff, garbageDisposed, nominatedLocation, occupiedToilets, notes,
      signature, contractorRemarks, resubmitSign
    } = req.body;

    if (!resubmitSign) {
      return res.status(400).send({ error: "Resubmit Signature is required to submit the form again." });
    }

    const formDocRef = db.collection('ctsForms').doc(formId);
    const doc = await formDocRef.get();

    if (!doc.exists) return res.status(404).send({ error: 'CTS Form not found.' });

    const existingData = doc.data();

    if (existingData.submittedById !== contractorSupervisor.uid) {
      return res.status(403).send({ error: "You are not authorized to resubmit this form." });
    }

    if (existingData.status !== 'REJECTED_BY_RAILWAY') {
      return res.status(400).send({ error: `Cannot resubmit form. Current status is ${existingData.status}` });
    }

    let trainName = existingData.trainName;
    let trainNumber = existingData.trainNumber;

    if (trainId && trainId !== existingData.trainId) {
      const trainDoc = await db.collection('trains').doc(trainId).get();
      if (trainDoc.exists) {
        const tData = trainDoc.data();
        trainName = tData.trainName || trainName;
        trainNumber = tData.trainNo || tData.trainNumber || trainNumber;
      }
    }

    const updateData = {
      trainId: trainId || existingData.trainId,
      trainName: trainName,
      trainNumber: trainNumber,
      formDateTime: formDateTime || existingData.formDateTime,
      actArrival: actArrival || existingData.actArrival,
      actDeparture: actDeparture || existingData.actDeparture,
      workStart: workStart || existingData.workStart,
      workEnd: workEnd || existingData.workEnd,
      platform: platform || existingData.platform,
      allowedWindow: allowedWindow || existingData.allowedWindow,
      lateYN: lateYN || existingData.lateYN,
      coachesInRake: coachesInRake || existingData.coachesInRake,
      coachesAttended: coachesAttended || existingData.coachesAttended,
      attendanceStaff: attendanceStaff || existingData.attendanceStaff,
      garbageDisposed: garbageDisposed !== undefined ? garbageDisposed : existingData.garbageDisposed,
      nominatedLocation: nominatedLocation || existingData.nominatedLocation,
      occupiedToilets: occupiedToilets || existingData.occupiedToilets,
      notes: notes || existingData.notes,
      signature: signature || existingData.signature,
      resubmitSignature: resubmitSign,
      status: 'RE-SUBMITTED',
      resubmittedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    if (contractorRemarks) {
      updateData.contractorRemarks = contractorRemarks;
    }

    await formDocRef.update(updateData);

    console.log(`(CTS-Form) Form ${formId} re-submitted for Train: ${trainName}`);
    res.status(200).send({
      message: 'CTS Form has been re-submitted successfully.',
      uid: formId
    });

  } catch (error) {
    console.error('(CTS-Form) Error resubmitting form:', error);
    res.status(500).send({ error: 'Failed to resubmit form', details: error.message });
  }
});

// =======================================================
// == API 5.16: CTS DETAILED STATS (UI-BASED BREAKDOWN)
// =======================================================
app.get('/api/reports/cts-stats', verifyToken, async (req, res) => {
  try {
    const { uid, role, userType, entityId, division } = req.user;
    const { startDate, endDate, contractId } = req.query;

    console.log(`(CTS-Stats) Fetching UI Breakdown for ${role}...`);

    let query = db.collection('ctsForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    const userRole = (role || '').toLowerCase();
    const isMaster = userRole.includes('master');
    const isAdmin = userRole.includes('admin');

    if (userType === 'contractor') {
      if (!entityId) return res.status(403).json({ error: "Entity ID missing." });
      query = query.where('submittedByEntityId', '==', entityId);

      if (!isMaster) {
        if (isAdmin) query = query.where('submittedByDivision', '==', division);
        else query = query.where('submittedById', '==', uid);
      }
    } else {
      if (isAdmin) query = query.where('submittedByDivision', '==', division);
      else if (!isMaster) query = query.where('submittedTo.railwayEmployeeId', '==', uid);
    }

    if (contractId) query = query.where('contractId', '==', contractId);

    const snapshot = await query.get();

    let stats = {
      totalTrains: 0,
      totalCoachesCleaned: 0,
      totalUnattendedCoaches: 0,
      totalSampledCoaches: 0,
      sumSamplingPercentage: 0,
      sumWindowTime: 0,
      windowTimeCount: 0,

      grades: { A: 0, B: 0, C: 0, D: 0 },

      ops: {
        toiletOccupied: 0,
        toiletUnattended: 0,
        totalScoreSum: 0,
        scoredFormsCount: 0
      },

      resources: {
        manpowerTotal: 0,
        chemicalQuantity: 0,
        machinesUsed: 0
      }
    };

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      let include = true;

      if (start || end) {
        if (!data.formDateTime) include = false;
        else {
          const fDate = new Date(data.formDateTime);
          if (start && fDate < start) include = false;
          if (end && fDate > end) include = false;
        }
      }

      if (include) {
        stats.totalTrains++;

        const attended = Number(data.coachesAttended) || 0;
        const inRake = Number(data.coachesInRake) || 0;
        stats.totalCoachesCleaned += attended;
        stats.totalUnattendedCoaches += (inRake - attended);

        if (data.allowedWindow) {
          stats.sumWindowTime += Number(data.allowedWindow);
          stats.windowTimeCount++;
        }

        stats.resources.manpowerTotal += (data.attendanceStaff ? data.attendanceStaff.length : 0);

        const rating = data.ratingDetails || {};
        const chemicalsArray = rating.chemicals || [];
        chemicalsArray.forEach(c => {
          stats.resources.chemicalQuantity += (parseFloat(c.quantity) || 0);
        });

        const summary = rating.summary || {};

        if (summary.overallGrade) {
          const g = summary.overallGrade;
          if (g.includes('(A)')) stats.grades.A++;
          else if (g.includes('(B)')) stats.grades.B++;
          else if (g.includes('(C)')) stats.grades.C++;
          else stats.grades.D++;

          stats.ops.totalScoreSum += (Number(summary.averageScore) || 0);
          stats.ops.scoredFormsCount++;
        }

        const evalTable = rating.coachEvaluationTable || [];
        stats.totalSampledCoaches += evalTable.length;

        if (rating.inspectionHeader) {
          stats.sumSamplingPercentage += (Number(rating.inspectionHeader.samplingPercentage) || 0);
        }

        const occupied = Number(data.occupiedToilets) || 0;
        stats.ops.toiletOccupied += occupied;
        stats.ops.toiletUnattended += occupied;
        if (Array.isArray(rating.machinesUsed)) {
          stats.resources.machinesUsed += rating.machinesUsed.length;
        }
      }
    });

    const totalGrades = stats.grades.A + stats.grades.B + stats.grades.C + stats.grades.D || 1;
    const avgScore = stats.ops.scoredFormsCount > 0 ? (stats.ops.totalScoreSum / stats.ops.scoredFormsCount).toFixed(2) : 0;
    const avgWindow = stats.windowTimeCount > 0 ? (stats.sumWindowTime / stats.windowTimeCount).toFixed(0) : 0;
    const avgSamplingPct = stats.totalTrains > 0 ? (stats.sumSamplingPercentage / stats.totalTrains).toFixed(1) : 0;

    const responseData = {
      statistics: {
        totalTrains: stats.totalTrains,
        totalCoachesCleaned: stats.totalCoachesCleaned,
        totalUnattendedCoaches: stats.totalUnattendedCoaches,
        sampledCoaches: stats.totalSampledCoaches,
        samplingPercentage: `${avgSamplingPct}%`,
        averageWindowTime: `${avgWindow} Min`
      },
      gradeDistribution: {
        A: { count: stats.grades.A, pct: ((stats.grades.A / totalGrades) * 100).toFixed(1) },
        B: { count: stats.grades.B, pct: ((stats.grades.B / totalGrades) * 100).toFixed(1) },
        C: { count: stats.grades.C, pct: ((stats.grades.C / totalGrades) * 100).toFixed(1) },
        D: { count: stats.grades.D, pct: ((stats.grades.D / totalGrades) * 100).toFixed(1) }
      },
      cleaningOperations: {
        toiletOccupied: stats.ops.toiletOccupied,
        averageScore: avgScore,
        toiletUnattended: stats.ops.toiletUnattended
      },
      resourcesUsed: {
        manpowerTotal: stats.resources.manpowerTotal,
        chemicalQuantity: stats.resources.chemicalQuantity.toFixed(1),
        machineryUsed: stats.resources.machinesUsed
      }
    };

    console.log(`(CTS-Stats) Success. Processed ${stats.totalTrains} trains.`);
    res.status(200).json(responseData);

  } catch (error) {
    console.error('(CTS-Stats) Error:', error);
    res.status(500).send({ error: error.message });
  }
});

// =======================================================
// == API: CTS REPORT DATA (FULL MAPPING & TRAIN LOOKUP)
// =======================================================
app.get('/api/reports/cts-data', verifyToken, async (req, res) => {
  try {
    const {
      zone, division, depot,
      startDate, endDate,
      contractorId,
      contractId,
      trainNo,
      supervisorId
    } = req.query;

    console.log(`(CTS-Report) Fetching Detailed Data for Excel...`);

    if (contractId) {
      const contractDoc = await db.collection('contracts').doc(contractId).get();
      if (contractDoc.exists) {
        const cData = contractDoc.data();
        const isStatusExpired = ['Expired', 'expired', 'Inactive', 'inactive'].includes(cData.status);
        let isDateExpired = false;
        if (cData.endDate) {
          const today = new Date();
          const end = new Date(cData.endDate);
          end.setHours(23, 59, 59, 999);
          if (today > end) isDateExpired = true;
        }
        if (isStatusExpired || isDateExpired) {
          return res.status(400).json({ error: "Contract Expired" });
        }
      }
    }

    let query = db.collection('ctsForms')
      .where('status', 'in', ['LOCKED', 'AUTO-APPROVED']);

    if (zone) query = query.where('submittedByZone', '==', zone);
    if (division) query = query.where('submittedByDivision', '==', division);
    if (depot) query = query.where('submittedByDepot', '==', depot);
    if (contractorId) query = query.where('submittedByEntityId', '==', contractorId);
    if (contractId) query = query.where('contractId', '==', contractId);
    if (supervisorId) query = query.where('submittedById', '==', supervisorId);

    const snapshot = await query.get();

    const allTrainNumbers = snapshot.docs.map(doc => doc.data().trainNumber).filter(Boolean);
    const uniqueTrainNumbers = [...new Set(allTrainNumbers)];
    let trainNamesMap = {};

    if (uniqueTrainNumbers.length > 0) {
      for (let i = 0; i < uniqueTrainNumbers.length; i += 30) {
        const chunk = uniqueTrainNumbers.slice(i, i + 30);
        const trainsSnap = await db.collection('trains').where('trainNo', 'in', chunk).get();
        trainsSnap.forEach(tDoc => {
          const tData = tDoc.data();
          trainNamesMap[tData.trainNo] = tData.trainName;
        });
      }
    }

    let reportData = [];

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    snapshot.forEach(doc => {
      const data = doc.data();
      const formDate = data.formDateTime ? new Date(data.formDateTime) : null;

      let include = true;
      if (start && formDate < start) include = false;
      if (end && formDate > end) include = false;
      if (include && trainNo && !String(data.trainNumber || "").includes(trainNo)) include = false;

      if (include) {
        const rating = data.ratingDetails || {};
        const summary = rating.summary || {};
        const inspection = rating.inspectionHeader || {};
        const evalTable = rating.coachEvaluationTable || [];

        let dist = {
          jet: { A: 0, B: 0, C: 0, D: 0 },
          basin: { A: 0, B: 0, C: 0, D: 0 },
          disposal: { A: 0, B: 0, C: 0, D: 0 }
        };

        const getGradeFromScore = (score) => {
          if (score >= 3) return 'A';
          if (score === 2) return 'B';
          if (score === 1) return 'C';
          return 'D';
        };

        evalTable.forEach(row => {
          dist.jet[getGradeFromScore(row.jetCleaningScore || 0)]++;
          dist.basin[getGradeFromScore(row.basinCleaningScore || 0)]++;
          dist.disposal[getGradeFromScore(row.disposalScore || 0)]++;
        });

        let totalChemical = 0;
        const chemicalsArray = rating.chemicals || [];
        chemicalsArray.forEach(c => totalChemical += (parseFloat(c.quantity) || 0));

        const formatDateTime = (isoStr) => {
          if (!isoStr) return 'N/A';
          try {
            const d = new Date(isoStr);
            const datePart = d.toLocaleDateString('en-GB');
            const timePart = d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false });
            return `${datePart} ${timePart}`;
          } catch (e) { return 'N/A'; }
        };

        reportData.push({
          date: formDate ? formDate.toLocaleDateString('en-GB') : 'N/A',
          trainNo: data.trainNumber || 'N/A',
          trainName: trainNamesMap[data.trainNumber] || data.submittedTo?.trainName || 'N/A',

          actualArrival: formatDateTime(data.actArrival),
          actualDeparture: formatDateTime(data.actDeparture),
          workStart: formatDateTime(data.workStart),
          workEnd: formatDateTime(data.workEnd),

          contractorSupervisor: data.submittedByName || 'N/A',
          railwaySupervisor: data.submittedTo?.railwayEmployeeName || 'N/A',

          totalCoaches: data.coachesInRake || 0,
          attendedCoaches: data.coachesAttended || 0,
          unattendedCoaches: (Number(data.coachesInRake) || 0) - (Number(data.coachesAttended) || 0),

          late: data.lateYN || 'No',
          garbageDisposed: data.garbageDisposed ? 'Yes' : 'No',
          location: data.nominatedLocation || 'N/A',

          machinesUsedCount: Array.isArray(rating.machinesUsed) ? rating.machinesUsed.length : 0,
          chemicalUsedLiter: totalChemical.toFixed(2),

          jet_A: dist.jet.A || 0,
          jet_B: dist.jet.B || 0,
          jet_C: dist.jet.C || 0,
          jet_D: dist.jet.D || 0,

          basin_A: dist.basin.A || 0,
          basin_B: dist.basin.B || 0,
          basin_C: dist.basin.C || 0,
          basin_D: dist.basin.D || 0,

          disposal_A: dist.disposal.A || 0,
          disposal_B: dist.disposal.B || 0,
          disposal_C: dist.disposal.C || 0,
          disposal_D: dist.disposal.D || 0,

          sampledPercentage: inspection.samplingPercentage ? `${inspection.samplingPercentage}%` : '0%',
          sampledCoaches: evalTable.length,
          overallGrade: summary.overallGrade || 'N/A',
          overallScore: summary.averageScore || 0,

          actualManpower: Array.isArray(data.attendanceStaff) ? data.attendanceStaff.length : 0,
          manpowerShortage: data.manpowerShortage || 'Nil'
        });
      }
    });

    res.status(200).json({
      count: reportData.length,
      data: reportData
    });

  } catch (error) {
    console.error('(CTS-Report) Error:', error);
    res.status(500).send({ error: 'Failed to fetch CTS report data', details: error.message });
  }
});

// ─── HAVERSINE DISTANCE HELPER (GPS Geo-fencing) ──────────────────────
function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = (deg) => (deg * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

//const { RekognitionClient, CompareFacesCommand } = require("@aws-sdk/client-rekognition");
//const axios = require('axios');

// AWS Rekognition Client Initialization
let rekognition;
try {
  rekognition = new RekognitionClient({
    region: process.env.AWS_REGION || "ap-south-1",
    credentials: {
      accessKeyId: process.env.AWS_ACCESS_KEY_ID,
      secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY
    }
  });
  console.log('✅ AWS Rekognition initialized');
} catch (error) {
  console.error('❌ Failed to initialize AWS Rekognition:', error.message);
}

/**
 * Core AWS Helper: Fetches remote images as binary buffers and evaluates face structures.
 * Fully optimized for handling Firebase Storage tokens and special character URL contexts.
 */

async function compareFaces(image1Url, image2Url) {
  try {
    console.log(`[AWS Debug] Source: [${image1Url}]`);
    console.log(`[AWS Debug] Target: [${image2Url}]`);

    if (!image1Url || !image2Url || image1Url.includes('undefined') || image2Url.includes('undefined')) {
      console.error(`[AWS Parameter Blocked] Target URLs are invalid or undefined. Source: [${image1Url}], Target: [${image2Url}]`);
      return { matched: false, similarity: 0, reason: "One or both image URLs are missing or invalid." };
    }

    console.log(`[AWS Biometric] Parsing paths from URLs...`);

    const getStoragePath = (url) => {
      try {
        const decodedUrl = decodeURIComponent(url);
        const parts = decodedUrl.split('/o/');
        if (parts.length > 1) {
          return parts[1].split('?')[0];
        }
        console.error(`[AWS Path Block] URL split failed — no '/o/' found after decode: ${decodedUrl.substring(0, 100)}`);
        return null;
      } catch (e) {
        console.error(`[AWS Path Block] decodeURIComponent threw for: ${url.substring(0, 100)}`, e.message);
        return null;
      }
    };

    const path1 = getStoragePath(image1Url);
    const path2 = getStoragePath(image2Url);

    if (!path1 || !path2) {
      console.error(`[AWS Path Block] Could not extract valid cloud paths.`);
      return { matched: false, similarity: 0, reason: "Failed to process image location parameters." };
    }

    console.log(`[AWS Biometric] Downloading buffers directly: Path1 [${path1}], Path2 [${path2}]`);

    const [fileBuffer1] = await bucket.file(path1).download();
    const [fileBuffer2] = await bucket.file(path2).download();

    if (!fileBuffer1 || fileBuffer1.length === 0) {
      console.error(`[AWS Buffer Block] First image buffer empty/corrupt. Path: ${path1}`);
      return { matched: false, similarity: 0, reason: "Baseline snapshot is empty or corrupt." };
    }
    if (!fileBuffer2 || fileBuffer2.length === 0) {
      console.error(`[AWS Buffer Block] Second image buffer empty/corrupt. Path: ${path2}`);
      return { matched: false, similarity: 0, reason: "Current snapshot is empty or corrupt." };
    }

    console.log(`[AWS Biometric] Both buffers loaded successfully (${fileBuffer1.length} bytes / ${fileBuffer2.length} bytes). Sending to Rekognition...`);

    try {
      const command = new CompareFacesCommand({
        SourceImage: { Bytes: fileBuffer1 },
        TargetImage: { Bytes: fileBuffer2 },
        SimilarityThreshold: 75
      });

      const data = await rekognition.send(command);

      if (data.FaceMatches && data.FaceMatches.length > 0) {
        const matchScore = data.FaceMatches[0].Similarity;
        console.log(`[AWS Rekognition Success] Face structure verified. Match score: ${matchScore}%`);
        return { matched: true, similarity: matchScore, reason: "Face matched successfully." };
      }

      if (data.UnmatchedFaces && data.UnmatchedFaces.length > 0) {
        console.warn(`[AWS Rekognition Audit] Faces detected but all below 75% threshold. Unmatched: ${data.UnmatchedFaces.length}`);
      } else {
        console.warn(`[AWS Rekognition Audit] No faces found in comparison. Possible: bad lighting, angle, or obstruction.`);
      }
      return { matched: false, similarity: 0, reason: "Face structure does not match the baseline profile selfie." };

    } catch (awsException) {
      let errorAlertMessage = "Biometric analysis failed due to poor image properties.";

      if (awsException.message.includes("contains no faces") || awsException.name === "InvalidParameterException") {
        errorAlertMessage = "Face not detected! Please ensure the photo is clear, well-lit, and contains a visible human face.";
      } else if (awsException.name === "ImageTooLargeException") {
        errorAlertMessage = "The uploaded file size exceeds the allowed processing dimensions.";
      }

      console.error(`[AWS Handled Trace]: ${errorAlertMessage} (${awsException.message})`);
      return { matched: false, similarity: 0, reason: errorAlertMessage };
    }

  } catch (err) {
    console.error("AWS Rekognition core wrapper failure:", err.message);
    return { matched: false, similarity: 0, reason: "Internal system error during face comparison process." };
  }
}

async function verifyFaceLiveness(imageUrl, expectedChallenge) {
  try {
    if (!imageUrl || imageUrl.includes('undefined')) {
      return { matched: false, reason: "Invalid image URL." };
    }

    const getStoragePath = (url) => {
      try {
        const decodedUrl = decodeURIComponent(url);
        const parts = decodedUrl.split('/o/');
        if (parts.length > 1) {
          return parts[1].split('?')[0];
        }
        return null;
      } catch (e) {
        return null;
      }
    };

    const path = getStoragePath(imageUrl);
    if (!path) return { matched: false, reason: "Failed to extract cloud path." };

    const [fileBuffer] = await bucket.file(path).download();
    if (!fileBuffer || fileBuffer.length === 0) return { matched: false, reason: "Empty image buffer." };

    if (expectedChallenge === 'THUMBS_UP' || expectedChallenge === 'FIST') {
      const command = new DetectLabelsCommand({
        Image: { Bytes: fileBuffer },
        MaxLabels: 20,
      });

      const data = await rekognition.send(command);
      const labels = data.Labels || [];
      
      let matched = false;
      const detectedLabels = labels.map(l => l.Name.toLowerCase());
      console.log(`[Liveness Debug] Detected labels:`, detectedLabels);

      if (expectedChallenge === 'THUMBS_UP') {
        matched = labels.some(l => {
          const name = l.Name.toLowerCase();
          return (name.includes('thumb') || name.includes('hand') || name.includes('finger') || name.includes('gesture')) && l.Confidence > 60;
        });
      } else if (expectedChallenge === 'FIST') {
        matched = labels.some(l => {
          const name = l.Name.toLowerCase();
          return (name.includes('fist') || name.includes('hand') || name.includes('clenched')) && l.Confidence > 60;
        });
      }

      if (matched) {
        return { matched: true, reason: `Successfully verified gesture challenge: ${expectedChallenge}` };
      } else {
        return { matched: false, reason: `Failed to detect the expected gesture: ${expectedChallenge}` };
      }
    } else {
      const command = new DetectFacesCommand({
        Image: { Bytes: fileBuffer },
        Attributes: ['ALL']
      });

      const data = await rekognition.send(command);
      if (!data.FaceDetails || data.FaceDetails.length === 0) {
        return { matched: false, reason: "No face detected in the image." };
      }

      const face = data.FaceDetails[0];
      let matched = false;

      if (expectedChallenge === 'SMILE') {
        matched = face.Smile && face.Smile.Value === true && face.Smile.Confidence > 70;
      } else if (expectedChallenge === 'EYES_OPEN') {
        matched = face.EyesOpen && face.EyesOpen.Value === true && face.EyesOpen.Confidence > 70;
      } else if (expectedChallenge === 'BLINK' || expectedChallenge === 'EYES_CLOSED') {
        matched = face.EyesOpen && face.EyesOpen.Value === false && face.EyesOpen.Confidence > 70;
      } else {
        matched = true;
      }

      if (matched) {
        return { matched: true, reason: `Successfully verified face challenge: ${expectedChallenge}` };
      } else {
        return { matched: false, reason: `Failed to verify expected face challenge: ${expectedChallenge}` };
      }
    }
  } catch (err) {
    console.error("AWS Rekognition liveness failure:", err.message);
    return { matched: false, reason: "Internal system error during liveness verification." };
  }
}

// =======================================================
// == API 4.7.1: Worker Attendance Issue Reporting
// =======================================================
app.post('/api/obhs/attendance/report-issue', verifyToken, async (req, res) => {
  try {
    const {
      runInstanceId,
      issueType,
      remark,
      photoUrl,
      latitude,
      longitude,
      attendanceType
    } = req.body;

    if (!runInstanceId || !issueType) {
      return res.status(400).send({ error: "runInstanceId and issueType are required." });
    }

    const { uid: userId, fullName: userName, role } = req.user;

    const exceptionId = `EXC-${Date.now()}`;
    const exceptionData = {
      uid: exceptionId,
      workerId: userId,
      workerName: userName,
      workerRole: role,
      runInstanceId,
      issueType,
      remark: remark || '',
      photoUrl: photoUrl || null,
      latitude: latitude || null,
      longitude: longitude || null,
      attendanceType: attendanceType || 'start',
      status: 'PENDING', // PENDING, APPROVED, REJECTED, EXCUSED
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await db.collection('attendance_exceptions').doc(exceptionId).set(exceptionData);

    // Also mirror to Firestore collection 'attendance_issues' for real-time tracking if needed
    console.log(`(Attendance) Exception Reported: ${issueType} by ${userName}`);

    res.status(201).send({
      success: true,
      message: "Attendance issue reported successfully. CTS will review it.",
      exceptionId
    });

  } catch (error) {
    console.error('(Attendance Issue) Error:', error);
    res.status(500).send({ error: 'Failed to report issue', details: error.message });
  }
});

// =======================================================
// == API 4.7.2: Get Attendance Exceptions (CTS/Admin View)
// =======================================================
app.get('/api/obhs/attendance/exceptions', verifyToken, async (req, res) => {
  try {
    const { status, runInstanceId } = req.query;
    let query = db.collection('attendance_exceptions');

    if (status) query = query.where('status', '==', status);
    if (runInstanceId) query = query.where('runInstanceId', '==', runInstanceId);

    const snapshot = await query.orderBy('createdAt', 'desc').get();
    const exceptions = [];
    snapshot.forEach(doc => exceptions.push(doc.data()));

    res.status(200).json({ count: exceptions.length, exceptions });
  } catch (error) {
    console.error('(Attendance Exceptions) Error:', error);
    res.status(500).json({ error: 'Failed to fetch exceptions', details: error.message });
  }
});

// =======================================================
// == API 4.7.3: Resolve Attendance Exception (CTS Action)
// =======================================================
app.post('/api/obhs/attendance/exceptions/action', verifyToken, async (req, res) => {
  try {
    const { exceptionId, action, adminRemark } = req.body; // action: 'APPROVED', 'REJECTED', 'EXCUSED'

    if (!exceptionId || !action) {
      return res.status(400).send({ error: "exceptionId and action are required." });
    }

    const { uid: adminId, fullName: adminName } = req.user;

    const excRef = db.collection('attendance_exceptions').doc(exceptionId);
    const doc = await excRef.get();

    if (!doc.exists) return res.status(404).send({ error: "Exception record not found." });

    const updateData = {
      status: action,
      adminRemark: adminRemark || '',
      resolvedBy: adminId,
      resolvedByName: adminName,
      resolvedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await excRef.update(updateData);

    // If APPROVED or EXCUSED, we might want to automatically mark the worker as PRESENT in the main attendance table
    if (action === 'APPROVED' || action === 'EXCUSED' || action === 'MARK_PRESENT') {
      const exception = doc.data();
      // Logic to update/create attendance record could go here
    }

    res.status(200).send({
      success: true,
      message: `Exception ${action} successfully.`
    });

  } catch (error) {
    console.error('(Resolve Exception) Error:', error);
    res.status(500).send({ error: 'Failed to resolve exception', details: error.message });
  }
});

app.post('/api/obhs/attendance', verifyToken, async (req, res) => {
  try {
    const allowedFields = [
      'runInstanceId',
      'attendanceType',
      'imageUrl',
      'latitude',
      'longitude',
      'deviceTimestamp',
      'mobileNumber',
      'deviceId'
    ];

    const bodyKeys = Object.keys(req.body);
    for (const key of bodyKeys) {
      if (!allowedFields.includes(key)) {
        return res.status(400).send({
          error: `Invalid field name.`,
          details: `The field '${key}' is not allowed.`
        });
      }
    }

    const { runInstanceId, attendanceType, imageUrl, latitude, longitude, deviceTimestamp, mobileNumber, deviceId } = req.body;

    if (!runInstanceId || !attendanceType || !imageUrl || !deviceTimestamp) {
      return res.status(400).send({
        error: "Missing fields.",
        details: "runInstanceId, attendanceType, imageUrl, and deviceTimestamp are required."
      });
    }

    const validTypes = ['start', 'mid', 'end'];
    if (!validTypes.includes(attendanceType)) {
      return res.status(400).send({
        error: "Invalid attendanceType.",
        details: "Allowed values are: 'start', 'mid', 'end'"
      });
    }

    const { uid: workerId, fullName: workerName } = req.user;
    const finalWorkerName = workerName || 'Unknown Worker';

    // ─── ATTENDANCE WINDOW VALIDATION (First-Worker Anchor Logic) ──────────
    let isLateAttendance = false;
    let lateByMinutes = 0;
    let firstAttendanceTime = null;

    try {
      if (runInstanceId && attendanceType === 'start') {
        const runSnap = await db.collection('RunInstance').doc(runInstanceId).get();
        if (runSnap.exists) {
          const runData = runSnap.data();

          firstAttendanceTime = runData.first_attendance_time;
          const currentTimestamp = new Date(deviceTimestamp || new Date().toISOString());

          if (!firstAttendanceTime) {
            // This is the first worker! Anchor the attendance window.
            firstAttendanceTime = currentTimestamp.toISOString();
            await db.collection('RunInstance').doc(runInstanceId).update({
              first_attendance_time: firstAttendanceTime,
              updatedAt: new Date().toISOString()
            });
            console.log(`(Attendance Anchor) RunInstance ${runInstanceId} anchored at ${firstAttendanceTime}`);
          } else {
            // Compare current timestamp against anchor
            const anchorTime = new Date(firstAttendanceTime);
            const diffMs = currentTimestamp.getTime() - anchorTime.getTime();
            const diffMins = Math.floor(diffMs / 60000);

            // 15 minutes window from first attendance
            if (diffMins > 15) {
              isLateAttendance = true;
              lateByMinutes = diffMins - 15; // Late by amount exceeding 15 min window
              console.log(`(Attendance Engine) Worker ${workerId} is late by ${lateByMinutes} mins. Anchor: ${firstAttendanceTime}`);
            }
          }
        }
      }
    } catch (winErr) {
      console.error('(Attendance Timing Engine) Error:', winErr);
    }

    // ─── GPS GEO-FENCING & SHIFT-TIMING VALIDATION ─────────────────────
    let resolvedStationLat = null;
    let resolvedStationLng = null;
    let resolvedStationName = null;
    let resolvedStationId = null;

    if (runInstanceId) {
      const runInstanceDoc = await db.collection('RunInstance').doc(runInstanceId).get();
      if (runInstanceDoc.exists) {
        const runData = runInstanceDoc.data();
        const parentTrainId = runData.parentTrainId;

        // Resolve station from train's origin
        let stationName = '';
        if (parentTrainId) {
          const trainDoc = await db.collection('trains').doc(parentTrainId).get();
          if (trainDoc.exists) {
            stationName = trainDoc.data().origin || '';
          }
        }

        if (stationName) {
          const stationSnap = await db.collection('stations')
            .where('stationName', '==', stationName)
            .where('active', '==', true)
            .limit(1)
            .get();

          if (!stationSnap.empty) {
            const st = stationSnap.docs[0].data();
            resolvedStationLat = parseFloat(st.latitude);
            resolvedStationLng = parseFloat(st.longitude);
            resolvedStationName = st.stationName;
            resolvedStationId = st.uid;
          }
        }
      }
    }

    // GPS Geo-fence check
    if (latitude && longitude && resolvedStationLat && resolvedStationLng) {
      const workerLat = parseFloat(latitude);
      const workerLng = parseFloat(longitude);
      const distance = haversineDistance(workerLat, workerLng, resolvedStationLat, resolvedStationLng);
      const geofenceRadius = 500; // metres

      if (distance > geofenceRadius) {
        return res.status(403).send({
          error: "Location Out of Bounds",
          details: `You are ${Math.round(distance)}m away from ${resolvedStationName}. Please mark attendance within ${geofenceRadius}m of the station.`
        });
      }
    }

    // Shift-timing validation
    if (resolvedStationId && workerId) {
      const scheduleSnap = await db.collection('stationSchedules')
        .where('stationId', '==', resolvedStationId)
        .where('entityId', '==', workerId)
        .where('active', '==', true)
        .limit(1)
        .get();

      if (!scheduleSnap.empty) {
        const schedule = scheduleSnap.docs[0].data();
        const startTime = schedule.startTime;
        const endTime = schedule.endTime;

        if (startTime && endTime) {
          const [startH, startM] = startTime.split(':').map(Number);
          const [endH, endM] = endTime.split(':').map(Number);
          const now = new Date();
          const currentMinutes = now.getHours() * 60 + now.getMinutes();
          const startMinutes = startH * 60 + startM;
          const endMinutes = endH * 60 + endM;

          if (currentMinutes < startMinutes || currentMinutes > endMinutes) {
            return res.status(403).send({
              error: "Outside Shift Hours",
              details: `Attendance can only be marked between ${startTime} and ${endTime}. Current time is outside your shift window.`
            });
          }
        }
      }
    }
    // ─── END VALIDATION ─────────────────────────────────────────────────

    const attendanceDocId = `${runInstanceId}_${workerId}`;
    const attendanceRef = db.collection('obhs_attendance').doc(attendanceDocId);
    const attendanceDoc = await attendanceRef.get();

    const attendanceEntry = {
      photoUrl: imageUrl,
      deviceTimestamp: deviceTimestamp,
      serverTimestamp: new Date().toISOString(),
      location: (latitude && longitude) ? { latitude, longitude } : null,
      mobileNumber: mobileNumber || null,
      deviceId: deviceId || null,
      isLate: isLateAttendance,
      lateByMinutes: lateByMinutes,
      firstAttendanceReference: firstAttendanceTime,
      status: isLateAttendance ? 'LATE' : 'PRESENT'
    };

    // --- CASE 1: INITIAL SUBSCRIPTION LOOP ('start' execution) ---
    if (!attendanceDoc.exists) {
      if (attendanceType !== 'start') {
        return res.status(400).send({
          error: "Workflow Violation",
          details: "You must submit 'start' attendance first."
        });
      }

      const newRecord = {
        uid: attendanceDocId,
        runInstanceId,
        workerId,
        workerName: finalWorkerName,
        mobileNumber: mobileNumber || null,
        deviceId: deviceId || null,
        isStartMarked: true,
        isMidMarked: false,
        isEndMarked: false,
        attendanceStatus: isLateAttendance ? 'LATE' : 'PRESENT',
        lateByMinutes: lateByMinutes,
        firstAttendanceReference: firstAttendanceTime,
        identityAuditStatus: "PENDING_VERIFICATION",
        startAttendance: attendanceEntry,
        midAttendance: null,
        endAttendance: null,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };

      await attendanceRef.set(newRecord);

      // Check if all workers for this journey have marked attendance → transition to READY
      try {
        if (runInstanceId && attendanceType === 'start') {
          const runSnap2 = await db.collection('RunInstance').doc(runInstanceId).get();
          if (runSnap2.exists && runSnap2.data().status === 'ALLOCATED') {
            const runData2 = runSnap2.data();
            const totalWorkers = (runData2.coaches || []).filter(c => c.workerId).length;
            const attendSnap = await db.collection('obhs_attendance')
              .where('runInstanceId', '==', runInstanceId)
              .where('isStartMarked', '==', true)
              .get();
            if (attendSnap.size >= totalWorkers && totalWorkers > 0) {
              await db.collection('RunInstance').doc(runInstanceId).update({
                status: 'READY',
                updatedAt: new Date().toISOString()
              });
              console.log(`(Attendance) Journey ${runInstanceId} → READY (all ${totalWorkers} workers reported)`);
            }
          }
        }
      } catch (readyErr) {
        console.error('(Attendance->READY) Error:', readyErr.message);
      }
    }
    // --- CASE 2: INCREMENTAL VERIFICATION LOOPS ('mid' or 'end') ---
    else {
      const currentData = attendanceDoc.data();
      const updateData = { updatedAt: new Date().toISOString() };

      if (attendanceType === 'start') {
        return res.status(400).send({ error: "Start attendance already submitted for this trip." });
      }

      // Extract baseline primary reference image
      const baseStartPhoto = currentData.startAttendance?.photoUrl;
      if (!baseStartPhoto) {
        return res.status(400).send({
          error: "Workflow Corrupted",
          details: "Baseline profile image missing from DB. Try restarting the trip shift lifecycle."
        });
      }

      // --- AWS REKOGNITION BIOMETRIC EVALUATION ENGINE ---
      const faceVerification = await compareFaces(baseStartPhoto, imageUrl);

      console.log(`[Attendance FaceCheck] ${attendanceType} for worker ${workerId}: matched=${faceVerification.matched}, similarity=${faceVerification.similarity}%, reason=${faceVerification.reason}`);

      if (!faceVerification.matched) {
        // Log exact similarity for debugging
        const auditNote = `Face mismatch on ${attendanceType} at ${new Date().toISOString()}: similarity=${faceVerification.similarity}%, reason=${faceVerification.reason}. Base: ${baseStartPhoto.substring(0, 80)}... Current: ${imageUrl.substring(0, 80)}...`;

        await attendanceRef.update({
          identityAuditStatus: "MISMATCH_ALERT",
          lastMismatchReason: faceVerification.reason,
          lastMismatchSimilarity: faceVerification.similarity,
          lastMismatchAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });

        return res.status(401).send({
          error: "Identity Verification Failed",
          details: `Face match score (${faceVerification.similarity}%) below threshold. ${faceVerification.reason}. Please retry with a clear, well-lit selfie facing the camera.`,
          mismatchSimilarity: faceVerification.similarity,
          mismatchReason: faceVerification.reason
        });
      }

      // Processing: MID Evaluation Target Assignment
      if (attendanceType === 'mid') {
        if (currentData.midAttendance) {
          return res.status(400).send({ error: "Mid attendance already submitted for this trip." });
        }
        updateData.midAttendance = attendanceEntry;
        updateData.isMidMarked = true;
        updateData.identityAuditStatus = "MID_VERIFIED";
      }

      // Processing: END Evaluation Target Assignment
      if (attendanceType === 'end') {
        if (currentData.endAttendance) {
          return res.status(400).send({ error: "End attendance already submitted for this trip." });
        }
        updateData.endAttendance = attendanceEntry;
        updateData.isEndMarked = true;
        updateData.identityAuditStatus = "VERIFIED_SUCCESS";
      }

      updateData.attendanceStatus = isLateAttendance ? 'LATE' : 'PRESENT';
      updateData.timingSnapshot = {
        isLate: isLateAttendance,
        lateByMinutes: lateByMinutes,
        firstAttendanceReference: firstAttendanceTime,
        recordedAt: new Date().toISOString()
      };
      await attendanceRef.update(updateData);
    }

    console.log(`(OBHS Attendance Matrix) Verified matching for ${attendanceType} on worker ${finalWorkerName}`);

    // Audit log for attendance
    await logAudit({
      action: `ATTENDANCE_${attendanceType.toUpperCase()}`,
      entityType: 'obhs_attendance',
      entityId: attendanceDocId,
      actorId: workerId,
      actorRole: req.user.role,
      actorName: finalWorkerName,
      newValue: { attendanceType, isLate: isLateAttendance, mobileNumber, deviceId },
      details: `${attendanceType} attendance marked - ${isLateAttendance ? 'LATE' : 'ON TIME'}`
    });

    // Update coach attendance status in RunInstance
    try {
      if (runInstanceId) {
        const updateKey = attendanceType === 'start' ? 'startAttendance' : (attendanceType === 'mid' ? 'midAttendance' : 'endAttendance');
        await db.collection('RunInstance').doc(runInstanceId).update({
          [`coachAttendance.${workerId}.${updateKey}`]: new Date().toISOString(),
          [`coachAttendance.${workerId}.status`]: isLateAttendance ? 'Late' : 'OnTime'
        });
      }
    } catch (updErr) {
      console.error('(Attendance Coach Update) Error:', updErr);
    }

    res.status(200).send({
      success: true,
      message: `${attendanceType.toUpperCase()} attendance face matched and processed successfully.`,
      uid: attendanceDocId,
      isLate: isLateAttendance,
      isNonCompliant: isLateAttendance
    });

  } catch (error) {
    console.error('(OBHS Attendance Engine) Thread Exception:', error);
    res.status(500).send({ error: 'Failed to process attendance lifecycle', details: error.message });
  }
});

// =======================================================
// == API 4.7.2: Standalone Face Comparison / Verification APIs
// =======================================================
app.post('/api/obhs/attendance/verify-face', verifyToken, async (req, res) => {
  try {
    const { image1Url, image2Url } = req.body;
    if (!image1Url || !image2Url) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both image1Url and image2Url are required in the request body."
      });
    }

    const faceVerification = await compareFaces(image1Url, image2Url);

    if (!faceVerification.matched) {
      return res.status(400).send({
        success: false,
        error: "Identity Verification Failed",
        details: faceVerification.reason
      });
    }

    res.status(200).send({
      success: true,
      message: "Face verified successfully",
      similarity: faceVerification.similarity
    });
  } catch (error) {
    console.error('(Face Verification Endpoint) Error:', error);
    res.status(500).send({ error: 'Failed to verify face', details: error.message });
  }
});

app.post('/api/verifyFace', verifyToken, async (req, res) => {
  try {
    const { image1Url, image2Url } = req.body;
    if (!image1Url || !image2Url) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both image1Url and image2Url are required."
      });
    }
    const faceVerification = await compareFaces(image1Url, image2Url);
    res.status(faceVerification.matched ? 200 : 400).send({
      success: faceVerification.matched,
      similarity: faceVerification.similarity,
      reason: faceVerification.reason
    });
  } catch (error) {
    res.status(500).send({ error: 'Face verification failed', details: error.message });
  }
});

app.post('/api/compareFace', verifyToken, async (req, res) => {
  try {
    const { image1Url, image2Url } = req.body;
    if (!image1Url || !image2Url) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both image1Url and image2Url are required."
      });
    }
    const faceVerification = await compareFaces(image1Url, image2Url);
    res.status(faceVerification.matched ? 200 : 400).send({
      success: faceVerification.matched,
      similarity: faceVerification.similarity,
      reason: faceVerification.reason
    });
  } catch (error) {
    res.status(500).send({ error: 'Face comparison failed', details: error.message });
  }
});

app.post('/api/verifyLiveness', verifyToken, async (req, res) => {
  try {
    const { imageUrl, expectedChallenge } = req.body;
    if (!imageUrl || !expectedChallenge) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both imageUrl and expectedChallenge are required."
      });
    }
    const result = await verifyFaceLiveness(imageUrl, expectedChallenge);
    res.status(result.matched ? 200 : 400).send(result);
  } catch (error) {
    console.error('(Liveness Verification Endpoint) Error:', error);
    res.status(500).send({ error: 'Liveness verification failed', details: error.message });
  }
});

app.post('/api/verifyFaceLiveness', verifyToken, async (req, res) => {
  try {
    const { imageUrl, expectedChallenge } = req.body;
    if (!imageUrl || !expectedChallenge) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both imageUrl and expectedChallenge are required."
      });
    }
    const result = await verifyFaceLiveness(imageUrl, expectedChallenge);
    res.status(result.matched ? 200 : 400).send(result);
  } catch (error) {
    console.error('(Face Liveness Verification Endpoint) Error:', error);
    res.status(500).send({ error: 'Face liveness verification failed', details: error.message });
  }
});

app.post('/api/obhs/attendance/verify-liveness', verifyToken, async (req, res) => {
  try {
    const { imageUrl, expectedChallenge } = req.body;
    if (!imageUrl || !expectedChallenge) {
      return res.status(400).send({
        error: "Missing parameters",
        details: "Both imageUrl and expectedChallenge are required."
      });
    }
    const result = await verifyFaceLiveness(imageUrl, expectedChallenge);
    res.status(result.matched ? 200 : 400).send(result);
  } catch (error) {
    console.error('(Attendance Liveness Verification Endpoint) Error:', error);
    res.status(500).send({ error: 'Liveness verification failed', details: error.message });
  }
});

// =======================================================
// == API 4.7.1: Get Attendance Status for App Launch Check
// =======================================================
app.get('/api/obhs/attendance/status', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    const { uid: workerId } = req.user;

    if (!runInstanceId) {
      return res.status(400).send({ error: "Missing query parameter: runInstanceId is required." });
    }

    const attendanceDocId = `${runInstanceId}_${workerId}`;
    const attendanceDoc = await db.collection('obhs_attendance').doc(attendanceDocId).get();

    if (!attendanceDoc.exists) {
      return res.status(200).json({
        success: true,
        isStartMarked: false,
        isMidMarked: false,
        isEndMarked: false,
        identityAuditStatus: "NOT_STARTED",
        message: "No attendance records found for this trip run instance."
      });
    }

    const data = attendanceDoc.data();

    return res.status(200).json({
      success: true,
      isStartMarked: data.isStartMarked || false,
      isMidMarked: data.isMidMarked || false,
      isEndMarked: data.isEndMarked || false,
      identityAuditStatus: data.identityAuditStatus || "PENDING_VERIFICATION",
      message: "Attendance state maps fetched successfully."
    });

  } catch (error) {
    console.error('(OBHS Status Get) Error:', error);
    return res.status(500).send({ error: 'Failed to fetch attendance state maps', details: error.message });
  }
});

// GET /api/obhs/attendance/list — List all attendance records (RBAC scoped)
app.get('/api/obhs/attendance/list', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    const { uid: callerId, role, trainId: userTrainId, trainIds: userTrainIds, division: userDiv } = req.user;
    const roleUpper = (role || '').toUpperCase();

    let query = db.collection('obhs_attendance');
    if (runInstanceId) query = query.where('runInstanceId', '==', runInstanceId);

    const snapshot = await query.get();
    let records = [];
    snapshot.forEach(doc => records.push(doc.data()));

    // ─── RBAC POST-FILTER ───────────────────────────────────────────────
    // Super-admins / Masters → see all
    // Railway Supervisor → see their assigned trains only (resolved via runInstanceId cross-reference below)
    // Contractor Supervisor (CTS) → see only their single assigned train
    // Worker → see only their own record
    if (roleUpper === 'WORKER' || roleUpper === 'RAILWAY WORKER') {
      records = records.filter(r => r.workerId === callerId);
    } else if (roleUpper === 'CTS' || roleUpper === 'CONTRACTOR SUPERVISOR') {
      // Scope to their single assigned train by cross-referencing runInstanceIds
      if (userTrainId) {
        const runSnaps = await db.collection('RunInstance')
          .where('parentTrainId', '==', userTrainId)
          .get();
        const allowedRunIds = new Set();
        runSnaps.forEach(d => allowedRunIds.add(d.id));
        records = records.filter(r => allowedRunIds.has(r.runInstanceId));
      }
    } else if (roleUpper === 'RAILWAY SUPERVISOR') {
      const assignedTrains = userTrainIds || (userTrainId ? [userTrainId] : []);
      if (assignedTrains.length > 0) {
        const runSnaps = await db.collection('RunInstance')
          .where('parentTrainId', 'in', assignedTrains.slice(0, 10)) // Firestore 'in' limit
          .get();
        const allowedRunIds = new Set();
        runSnaps.forEach(d => allowedRunIds.add(d.id));
        records = records.filter(r => allowedRunIds.has(r.runInstanceId));
      }
    } else if (roleUpper.includes('ADMIN') && !roleUpper.includes('SUPER') && !roleUpper.includes('MASTER')) {
      // Division Admin: scope to their division's runs
      if (userDiv) {
        const runSnaps = await db.collection('RunInstance')
          .where('division', '==', userDiv)
          .get();
        const allowedRunIds = new Set();
        runSnaps.forEach(d => allowedRunIds.add(d.id));
        records = records.filter(r => allowedRunIds.has(r.runInstanceId));
      }
    }
    // ─── END RBAC ────────────────────────────────────────────────────────

    // Sort by updatedAt desc
    records.sort((a, b) => ((b.updatedAt || '') > (a.updatedAt || '') ? 1 : -1));
    res.status(200).json({ count: records.length, records });
  } catch (error) {
    console.error('(OBHS Attendance List) Error:', error);
    res.status(500).json({ error: 'Failed to fetch attendance records', details: error.message });
  }
});


// =======================================================
// == API 4.8: Submit Completed OBHS Task (Single Submission)
// =======================================================
app.post('/api/obhs/tasks/submit', verifyToken, async (req, res) => {
  try {
    const allowedFields = [
      'taskId',
      'runInstanceId',
      'taskType',
      'coachNo',
      'frequencyIndex',
      'beforePhoto',
      'afterPhoto',
      'comment',
      'deviceTimestamp',
      'gpsLatitude',
      'gpsLongitude'
    ];

    const bodyKeys = Object.keys(req.body);
    for (const key of bodyKeys) {
      if (!allowedFields.includes(key)) {
        return res.status(400).send({ error: `Invalid field: '${key}'` });
      }
    }

    const {
      taskId: incomingTaskId,
      runInstanceId,
      taskType,
      coachNo,
      frequencyIndex,
      beforePhoto,
      afterPhoto,
      gpsLatitude,
      gpsLongitude,
      comment,
      deviceTimestamp
    } = req.body;

    // Validate mandatory fields
    if (!runInstanceId || !taskType || !coachNo || !beforePhoto || !afterPhoto || !deviceTimestamp) {
      return res.status(400).send({
        error: "Missing mandatory fields.",
        details: "runInstanceId, taskType, coachNo, beforePhoto, afterPhoto, and deviceTimestamp are required."
      });
    }

    // GPS is optional — worker may be in low-signal zone; default to null
    const finalGpsLat = (gpsLatitude && gpsLatitude !== '') ? gpsLatitude : null;
    const finalGpsLng = (gpsLongitude && gpsLongitude !== '') ? gpsLongitude : null;

    // Token se User Details nikalna
    const { uid: userId, fullName: userName, role } = req.user;

    // Use incoming taskId if provided (from task board), else generate one
    const taskId = incomingTaskId || `${runInstanceId}_${coachNo}_${taskType.replace(/\s+/g, '')}_${Date.now()}`;

    const taskRef = db.collection('obhs_tasks').doc(taskId);

    // Check if task already exists to prevent duplicate submissions
    const existingTask = await taskRef.get();
    if (existingTask.exists) {
      return res.status(400).send({ error: "This specific task instance has already been submitted." });
    }

    const taskData = {
      uid: taskId,
      runInstanceId,
      coachNo,
      taskType,
      frequencyIndex: frequencyIndex || null,

      // User mapping
      submittedBy: {
        id: userId,
        name: userName || 'Unknown User',
        role: role
      },

      // Evidence
      beforePhoto,
      afterPhoto,
      gpsLatitude: finalGpsLat,
      gpsLongitude: finalGpsLng,
      comment: comment || "",
      status: 'Completed',

      // Timing metadata
      deviceTimestamp,
      serverCreatedAt: new Date().toISOString()
    };

    await taskRef.set(taskData);

    // ─── Sync: Update corresponding task_detail to COMPLETED ───
    const taskTypeMap = {
      'Toilet Cleaning': 'toilet_cleaning',
      'Coach Cleaning': 'coach_cleaning',
      'Toilet Cleaning & Disinfection': 'toilet_cleaning',
      'Compartment Cleaning': 'coach_cleaning'
    };
    const mappedType = taskTypeMap[taskType] || taskType.toLowerCase().replace(/\s+/g, '_');
    const detailId = `${runInstanceId}_${mappedType}_${frequencyIndex ? frequencyIndex.replace(':', '') : ''}_${coachNo}`.replace(/\s+/g, '');

    try {
      const detailRef = db.collection('task_details').doc(detailId);
      const detailDoc = await detailRef.get();
      if (detailDoc.exists) {
        await detailRef.update({
          status: 'COMPLETED',
          beforePhoto,
          afterPhoto,
          gpsLatitude,
          gpsLongitude,
          completionTime: new Date().toISOString(),
          remarks: comment || null,
          deviceId: req.body.deviceId || null,
          mobileNumber: req.body.mobileNumber || null,
          updatedAt: new Date().toISOString()
        });
        // Try to auto-close parent task_header
        await autoCloseParentTask(detailRef, detailDoc);
      }
    } catch (syncErr) {
      console.error('(TaskSubmit) Detail sync warning:', syncErr.message);
      // Non-blocking: submission already succeeded
    }

    // ─── Audit log ───
    await logAudit({
      action: 'TASK_COMPLETED',
      entityType: 'task_details',
      entityId: detailId,
      actorId: userId,
      actorRole: role,
      actorName: userName,
      newValue: { status: 'Completed', coachNo, taskType, gpsLatitude, gpsLongitude },
      oldValue: { status: 'OPEN' },
      details: `${taskType} completed for Coach ${coachNo} with GPS evidence`
    });

    console.log(`(OBHS Task) ${taskType} for Coach ${coachNo} successfully submitted by ${userName}`);
    res.status(201).send({
      success: true,
      message: "Task data submitted and recorded successfully.",
      taskId: taskId
    });

  } catch (error) {
    console.error('(OBHS Task Submit) Error:', error);
    res.status(500).send({ error: 'Failed to record task submission', details: error.message });
  }
});

// =======================================================
// == API 4.9: Get OBHS Task Board (Due, Overdue, All Tabs Fully Corrected)
// =======================================================
app.get('/api/obhs/tasks/board', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;

    if (!runInstanceId) {
      return res.status(400).send({ error: "Missing query parameter: runInstanceId is required." });
    }

    // 1. Fetch all completed tasks for this specific run instance from DB
    const completedSnapshot = await db.collection('obhs_tasks')
      .where('runInstanceId', '==', runInstanceId)
      .get();

    const completedTaskIds = new Set();
    const completedTasksMap = {};

    completedSnapshot.forEach(doc => {
      const data = doc.data();
      completedTaskIds.add(data.uid);
      completedTasksMap[data.uid] = data;
    });

    // 2. Fetch train run details to get trip start time metadata
    // Note: Collection name updated to 'RunInstance' to match your exact DB structure
    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    let tripStartDate = new Date();

    if (runDoc.exists) {
      const runData = runDoc.data();
      const parentTrainId = runData.parentTrainId;
      const departureDate = runData.departureDate; // Format: "2026-04-30"

      if (parentTrainId && departureDate) {
        // Fetch the master train details to get journeyStartTime
        const trainDoc = await db.collection('trains').doc(parentTrainId).get();
        if (trainDoc.exists && trainDoc.data().journeyStartTime) {
          const journeyStartTime = trainDoc.data().journeyStartTime; // Format: "18:46:00"
          // Combine departureDate and journeyStartTime into a valid ISO string
          tripStartDate = new Date(`${departureDate}T${journeyStartTime}.000Z`);
        } else if (runData.createdAt) {
          tripStartDate = new Date(runData.createdAt);
        }
      } else if (runData.createdAt) {
        tripStartDate = new Date(runData.createdAt);
      }
    }

    // 3. Build the Dynamic Task Matrix from the worker's assigned coaches
    const coachesArr = runDoc.exists ? (runDoc.data().coaches || []) : [];
    const myCoaches = coachesArr.filter(c => c.workerId === req.user.uid);
    const coachTypes = myCoaches.length > 0
      ? myCoaches.map(c => c.coachType)
      : ['S2', 'S1']; // fallback if no coach mapping found

    const taskDefinitions = [];
    const toiletFreqs = [
      { freq: 'Hour 1', startRel: 0, duration: 1 },
      { freq: 'Hour 2', startRel: 1, duration: 1 },
      { freq: 'Hour 3', startRel: 2, duration: 1 },
      { freq: 'Hour 5', startRel: 4, duration: 2 },
      { freq: 'Hour 7', startRel: 6, duration: 2 },
      { freq: 'Hour 9', startRel: 8, duration: 2 },
    ];
    const coachCleaningFreqs = [
      { freq: 'Train Start', startRel: 0, duration: 3 },
      { freq: 'Mid Journey', startRel: 6, duration: 4 },
      { freq: 'Train End', startRel: 12, duration: 3 },
    ];
    const linenFreqs = [
      { freq: 'Initial Distribution', startRel: 0, duration: 4 },
      { freq: 'Night Return Check', startRel: 10, duration: 4 },
    ];

    for (const coach of coachTypes) {
      for (const f of toiletFreqs) {
        taskDefinitions.push({ type: 'Toilet Cleaning', coach, ...f });
      }
      for (const f of coachCleaningFreqs) {
        taskDefinitions.push({ type: 'Coach Cleaning', coach, ...f });
      }
      // Only first coach gets linen distribution
      if (coach === coachTypes[0]) {
        for (const f of linenFreqs) {
          taskDefinitions.push({ type: 'Linen Distribution', coach, ...f });
        }
      }
    }

    const now = new Date();

    const dueList = [];
    const overdueList = [];
    const allList = [];

    // 4. Compute target deadlines and evaluate state flags
    taskDefinitions.forEach(task => {
      const startTime = new Date(tripStartDate.getTime() + task.startRel * 60 * 60 * 1000);
      const endTime = new Date(startTime.getTime() + task.duration * 60 * 60 * 1000);

      const suffix = task.freq ? `_${task.freq.replace(/\s+/g, '')}` : '';
      const generatedTaskId = `${runInstanceId}_${task.coach}_${task.type.replace(/\s+/g, '')}${suffix}`;

      // Base Display Item for UI render pipeline mapping
      const taskItem = {
        uid: generatedTaskId,
        taskType: task.type,
        coachNo: task.coach,
        frequencyIndex: task.freq,
        scheduledStartTime: startTime.toISOString(),
        scheduledEndTime: endTime.toISOString(),
        status: 'Pending',
        beforePhoto: null,
        afterPhoto: null,
        comment: "",
        requiresPhoto: true,
        requiresComment: true
      };

      if (completedTaskIds.has(generatedTaskId)) {
        // Condition A: Completed Status (Sirf master 'All' list mein update hoga, baaki do tabs mein nahi dikhega)
        const savedData = completedTasksMap[generatedTaskId];
        taskItem.status = 'Completed';
        taskItem.beforePhoto = savedData.beforePhoto;
        taskItem.afterPhoto = savedData.afterPhoto;
        taskItem.comment = savedData.comment;
        taskItem.submittedBy = savedData.submittedBy;
        taskItem.serverCreatedAt = savedData.serverCreatedAt;
      } else if (now > endTime) {
        // Condition B: Time has passed and not done yet -> Overdue Tab
        taskItem.status = 'Overdue';
        overdueList.push(taskItem);
      } else {
        // Condition C: Jo task active hai ya future ka (Upcoming) hai -> Due Tab
        // Flutter UI consistency ke liye status string ko "Due" rakhein
        taskItem.status = 'Due';
        dueList.push(taskItem);
      }

      // Add everything to the consolidated Master "All" List response
      allList.push(taskItem);
    });

    // 5. Send structural lists perfectly mapped for Flutter 3-Options UI
    return res.status(200).json({
      success: true,
      meta: {
        runInstanceId,
        serverTimeEvaluated: now.toISOString(),
        counts: {
          due: dueList.length,
          overdue: overdueList.length,
          all: allList.length
        }
      },
      categories: {
        due: dueList,       // Goes straight to "Due" option
        overdue: overdueList, // Goes straight to "Overdue" option
        all: allList         // Goes straight to "All" option (Includes Completed cards)
      }
    });

  } catch (error) {
    console.error('(OBHS Task Board Engine) Error:', error);
    return res.status(500).send({
      error: 'Failed to compile categorized task matrix metrics.',
      details: error.message
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// == API 4.x: Bridge — Route Complaints / Feedback into OBHS Tasks Queue
// ═══════════════════════════════════════════════════════════════════════════

// ─── 4.x.1: Route a Complaint to the OBHS Task Queue ─────────────────────
app.post('/api/obhs/tasks/route-from-complaint/:complaintId', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.params;
    if (!complaintId) {
      return res.status(400).send({ error: "complaintId is required." });
    }

    const userRole = (req.user.role || '').toLowerCase();
    if (!['admin', 'supervisor', 'company master'].includes(userRole)) {
      return res.status(403).send({ error: "Only Admins, Supervisors, or Company Masters can route complaints to tasks." });
    }

    const complaintRef = db.collection('obhs_complaints').doc(complaintId);
    const complaintDoc = await complaintRef.get();

    if (!complaintDoc.exists) {
      return res.status(404).send({ error: `Complaint ${complaintId} not found.` });
    }

    const complaint = complaintDoc.data();

    if (complaint.status === 'ROUTED_TO_TASK') {
      return res.status(400).send({ error: "This complaint has already been routed to a task." });
    }

    // Build the task document
    const taskType = complaint.category || 'Cleaning Issue';
    const taskId = `${complaint.runInstanceId}_ROUTED_${complaint.coachNo}_${taskType.replace(/\s+/g, '')}_${Date.now()}`;

    const taskRef = db.collection('obhs_tasks').doc(taskId);
    const taskData = {
      uid: taskId,
      runInstanceId: complaint.runInstanceId,
      coachNo: complaint.coachNo,
      taskType: taskType,
      frequencyIndex: 'Routed', // Marks it as system-routed
      source: 'complaint',
      sourceId: complaintId,
      submittedBy: {
        id: req.user.uid,
        name: req.user.fullName || 'System',
        role: req.user.role || 'Admin'
      },
      beforePhoto: complaint.photoUrl || null,
      afterPhoto: null,
      comment: complaint.description || '',
      status: 'Pending',
      deviceTimestamp: new Date().toISOString(),
      serverCreatedAt: new Date().toISOString(),
      routedFrom: {
        type: 'complaint',
        id: complaintId,
        at: new Date().toISOString()
      }
    };

    await taskRef.set(taskData);

    // Mark the complaint as routed
    await complaintRef.update({
      status: 'ROUTED_TO_TASK',
      routedToTaskId: taskId,
      routedAt: new Date().toISOString()
    });

    console.log(`(Bridge) Complaint ${complaintId} routed to task ${taskId}`);
    res.status(201).send({
      success: true,
      message: 'Complaint routed to task queue successfully.',
      taskId: taskId,
      task: taskData
    });

  } catch (error) {
    console.error('(Bridge) Error routing complaint to task:', error);
    res.status(500).send({ error: 'Failed to route complaint to task', details: error.message });
  }
});

// ─── 4.x.2: Route Feedback to the OBHS Task Queue ─────────────────────────
app.post('/api/obhs/tasks/route-from-feedback/:feedbackId', verifyToken, async (req, res) => {
  try {
    const { feedbackId } = req.params;
    const { category } = req.body; // Optional override for task type

    if (!feedbackId) {
      return res.status(400).send({ error: "feedbackId is required." });
    }

    const userRole = (req.user.role || '').toLowerCase();
    if (!['admin', 'supervisor', 'company master'].includes(userRole)) {
      return res.status(403).send({ error: "Only Admins, Supervisors, or Company Masters can route feedback to tasks." });
    }

    const feedbackRef = db.collection('obhs_feedbacks').doc(feedbackId);
    const feedbackDoc = await feedbackRef.get();

    if (!feedbackDoc.exists) {
      return res.status(404).send({ error: `Feedback ${feedbackId} not found.` });
    }

    const feedback = feedbackDoc.data();

    if (feedback.routedToTaskId) {
      return res.status(400).send({ error: "This feedback has already been routed to a task." });
    }

    // Determine task type from ratings or category
    let taskType = category || 'Cleaning Required';
    if (!category && feedback.ratings) {
      const r = feedback.ratings;
      const lowestParam = Object.entries(r).reduce((a, b) => (a[1] < b[1] ? a : b));
      const paramLabels = {
        cleanliness: 'Coach Cleaning',
        toiletHygiene: 'Toilet Cleaning',
        linenQuality: 'Linen Replacement',
        security: 'Security Check',
        staffBehaviour: 'Staff Behaviour Review'
      };
      taskType = paramLabels[lowestParam[0]] || 'Cleaning Required';
    }

    const runInstanceId = feedback.runInstanceId || 'UNKNOWN_RUN';
    const taskId = `${runInstanceId}_ROUTED_${feedback.coachNo || 'GEN'}_${taskType.replace(/\s+/g, '')}_${Date.now()}`;

    const taskRef = db.collection('obhs_tasks').doc(taskId);
    const taskData = {
      uid: taskId,
      runInstanceId: runInstanceId,
      coachNo: feedback.coachNo || 'General',
      taskType: taskType,
      frequencyIndex: 'Routed',
      source: 'feedback',
      sourceId: feedbackId,
      submittedBy: {
        id: req.user.uid,
        name: req.user.fullName || 'System',
        role: req.user.role || 'Admin'
      },
      beforePhoto: feedback.photoUrl || null,
      afterPhoto: null,
      comment: feedback.remarks || `Routed from feedback (Overall: ${feedback.overallRating || 'N/A'})`,
      status: 'Pending',
      deviceTimestamp: new Date().toISOString(),
      serverCreatedAt: new Date().toISOString(),
      routedFrom: {
        type: 'feedback',
        id: feedbackId,
        at: new Date().toISOString()
      }
    };

    await taskRef.set(taskData);

    // Mark the feedback as routed
    await feedbackRef.update({
      routedToTaskId: taskId,
      routedAt: new Date().toISOString()
    });

    console.log(`(Bridge) Feedback ${feedbackId} routed to task ${taskId}`);
    res.status(201).send({
      success: true,
      message: 'Feedback routed to task queue successfully.',
      taskId: taskId,
      task: taskData
    });

  } catch (error) {
    console.error('(Bridge) Error routing feedback to task:', error);
    res.status(500).send({ error: 'Failed to route feedback to task', details: error.message });
  }
});

// ─── 4.x.3: List All Un-routed Complaints & Feedbacks (for admin dashboard) ─
app.get('/api/obhs/tasks/pending-routing', verifyToken, async (req, res) => {
  try {
    const userRole = (req.user.role || '').toLowerCase();
    if (!['admin', 'supervisor', 'company master'].includes(userRole)) {
      return res.status(403).send({ error: "Only Admins, Supervisors, or Company Masters can view pending routing items." });
    }

    // Fetch open (not yet routed) complaints
    const openComplaintsSnap = await db.collection('obhs_complaints')
      .where('status', '==', 'OPEN')
      .get();

    const unRoutedComplaints = [];
    openComplaintsSnap.forEach(doc => {
      const data = doc.data();
      if (!data.routedToTaskId) {
        unRoutedComplaints.push({
          sourceType: 'complaint',
          sourceId: data.complaintId || doc.id,
          runInstanceId: data.runInstanceId,
          trainNo: data.trainNo,
          coachNo: data.coachNo,
          category: data.category,
          description: data.description,
          photoUrl: data.photoUrl,
          submittedBy: data.submittedBy,
          createdAt: data.createdAt
        });
      }
    });

    // Fetch feedbacks not yet routed
    const allFeedbacksSnap = await db.collection('obhs_feedbacks').get();
    const unRoutedFeedbacks = [];
    allFeedbacksSnap.forEach(doc => {
      const data = doc.data();
      if (!data.routedToTaskId) {
        unRoutedFeedbacks.push({
          sourceType: 'feedback',
          sourceId: data.feedbackId || doc.id,
          feedbackType: data.feedbackType,
          runInstanceId: data.runInstanceId,
          trainNo: data.trainNo,
          coachNo: data.coachNo,
          overallRating: data.overallRating,
          ratings: data.ratings,
          remarks: data.remarks,
          photoUrl: data.photoUrl,
          passengerName: data.passengerName || null,
          inspectorName: data.inspectorName || null,
          targetWorker: data.targetWorker || null,
          createdAt: data.createdAt
        });
      }
    });

    // Sort by createdAt descending (newest first)
    const allItems = [...unRoutedComplaints, ...unRoutedFeedbacks]
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    res.status(200).send({
      success: true,
      counts: {
        complaints: unRoutedComplaints.length,
        feedbacks: unRoutedFeedbacks.length,
        total: allItems.length
      },
      items: allItems
    });

  } catch (error) {
    console.error('(Bridge) Error fetching pending routing items:', error);
    res.status(500).send({ error: 'Failed to fetch pending routing items', details: error.message });
  }
});

// NOTE: "verifyToken" aapke JWT verification middleware ka naam hona chahiye
app.post('/api/obhs/complaints/raise', verifyToken, async (req, res) => {
  try {
    const { runInstanceId: bodyRunInstanceId, coachNo, category, description, photoUrl } = req.body;

    // 1. Input Fields Validation
    if (!coachNo || !category || !description) {
      return res.status(400).json({
        success: false,
        error: "Coach number, category, and description are required fields."
      });
    }

    // 2. Safely Extract Worker Details from req.user (Middleware context)
    // Agar req.user poora object nahi mil raha toh backup handle karein
    const workerUser = req.user || req.userData;
    if (!workerUser || !workerUser.uid) {
      return res.status(401).json({
        success: false,
        error: "Unauthorized. Worker session not found in token context."
      });
    }

    const workerId = workerUser.uid;
    const workerName = workerUser.fullName || "Testing";

    // 3. Resilient RunInstance Logic (Fallback mechanism)
    // Priority: body > token > DB fallback
    let runInstanceId = bodyRunInstanceId || workerUser.activeRunInstanceId;

    // Fallback: Agar body ya token mein runInstanceId nahi mila, 
    // toh live database (RunInstance collection) se fetch karo jahan workerId matched ho.
    if (!runInstanceId) {
      console.log(`(Complaint) runInstanceId missing in request & token for worker ${workerId}. Fetching from DB...`);

      const runInstanceSnapshot = await db.collection('RunInstance')
        .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running'])
        .get();

      for (const doc of runInstanceSnapshot.docs) {
        const runData = doc.data();
        if (runData.coaches && Array.isArray(runData.coaches)) {
          const isWorkerAssigned = runData.coaches.some(coach => coach.workerId === workerId);
          if (isWorkerAssigned) {
            runInstanceId = runData.runInstanceId || doc.id;
            break;
          }
        }
      }
    }

    // Agar fallback check ke baad bhi nahi milta, tabhi bad request throw karein
    if (!runInstanceId) {
      return res.status(400).json({
        success: false,
        error: "No active train journey (RunInstance) currently mapped to this worker."
      });
    }

    // 4. Fetch Train Number & Data
    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) {
      return res.status(404).json({ success: false, error: "Active RunInstance details not found in database." });
    }

    const runData = runDoc.data();
    const trainNo = runData.trainNo || "234556";
    const trainName = runData.trainName || "exp";

    // 5. Save Complaint Document
    const complaintsRef = db.collection('obhs_complaints').doc();

    const complaintData = {
      complaintId: complaintsRef.id,
      runInstanceId: runInstanceId,
      trainNo: trainNo,
      trainName: trainName,
      coachNo: coachNo,
      category: category,
      description: description,
      photoUrl: photoUrl || null,
      status: "OPEN",
      date: new Date().toISOString().split('T')[0],
      createdAt: new Date().toISOString(),
      submittedBy: {
        uid: workerId,
        name: workerName,
        role: workerUser.role || "Railway Worker"
      }
    };

    await complaintsRef.set(complaintData);

    // ─── Auto-generate task from complaint ───
    try {
      const categoryMap = {
        'Toilet': 'TOILET_CLEANING', 'Garbage': 'GARBAGE_COLLECTION',
        'Water': 'WATER_CHECK', 'Cleaning': 'COACH_CLEANING',
        'Behaviour': 'STAFF_BEHAVIOUR', 'Repair': 'PETTY_REPAIR',
        'Electrical': 'ELECTRICAL_ISSUE'
      };
      const taskMasterCode = categoryMap[category] || 'COMPLAINT_GENERIC';
      const complaintTaskId = `${runInstanceId}_COMPLAINT_${complaintsRef.id}_${coachNo}`;

      // Find worker assigned to this coach
      let assignedWorker = null;
      if (runData.coaches) {
        const coach = runData.coaches.find(c => (c.coachPosition === coachNo || c.coachNo === coachNo));
        if (coach && coach.workerId) assignedWorker = coach;
      }

      const complaintTaskRef = db.collection('task_instances').doc(complaintTaskId);
      await complaintTaskRef.set({
        taskId: complaintTaskId, parentTaskId: null,
        runInstanceId, trainNo, coachNo,
        taskMasterCode, taskName: `Complaint: ${category} - ${description?.substring(0, 50)}`,
        category: 'complaint', subCategory: category,
        workerId: assignedWorker?.workerId || null,
        workerName: assignedWorker?.workerName || 'Unassigned',
        workerRole: assignedWorker?.workerRole || 'janitor',
        scheduledTime: new Date().toTimeString().slice(0, 5),
        scheduledDate: new Date().toISOString().split('T')[0],
        status: 'OPEN', priority: 2, isParentTask: false,
        childTaskIds: [], beforePhoto: null, afterPhoto: null,
        gpsLatitude: null, gpsLongitude: null,
        supervisorVerified: false,
        sourceComplaintId: complaintsRef.id,
        createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
      });
      console.log(`(Complaint) Auto-generated task ${complaintTaskId} from complaint ${complaintsRef.id}`);
    } catch (taskGenErr) {
      console.error('(Complaint->Task) Error:', taskGenErr.message);
      // Non-blocking: complaint already saved
    }

    console.log(`(Complaint) Successfully raised ID: ${complaintsRef.id} for Train: ${trainNo}`);

    res.status(201).json({
      success: true,
      message: "Complaint registered successfully",
      complaintId: complaintsRef.id,
      data: complaintData
    });

  } catch (error) {
    console.error('(Complaint) Error while raising complaint:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to raise complaint',
      details: error.message
    });
  }
});


// --- Route: Get Complaints (Unified for Admin, Company Master, Supervisor, and Worker) ---
app.get('/api/obhs/complaints', verifyToken, async (req, res) => {
  try {
    const currentUser = req.user || req.userData;
    if (!currentUser || !currentUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. Session not found." });
    }

    const userRole = currentUser.role || "Railway Worker";
    const userId = currentUser.uid;
    const { status, runInstanceId } = req.query;

    // 1. Base Reference of Firestore Collection
    let complaintsQuery = db.collection('obhs_complaints');

    // 2. ROLE-BASED FILTER (Condition Check - UPDATED)
    // In roles ko global access milega (Sari complaints dikhengi)
    const globalAccessRoles = ['Admin', 'Supervisor', 'Company Master'];

    if (!globalAccessRoles.includes(userRole)) {
      // Agar user upar diye gaye roles mein nahi hai, toh use worker treat karenge aur filter lagayenge
      console.log(`(Fetch Complaints) Filtering for Worker ID: ${userId}`);
      complaintsQuery = complaintsQuery.where('submittedBy.uid', '==', userId);
    } else {
      console.log(`(Fetch Complaints) Global Fetch Authorized for: ${currentUser.fullName} (${userRole})`);
    }

    // 3. STATUS FILTER (Optional Query Parameter Check)
    if (status) {
      const upperStatus = status.toUpperCase();
      complaintsQuery = complaintsQuery.where('status', '==', upperStatus);
    }

    // 4. Execute Query
    const snapshot = await complaintsQuery.get();

    if (snapshot.empty) {
      return res.status(200).json({
        success: true,
        count: 0,
        complaints: []
      });
    }

    let complaintsList = [];
    snapshot.docs.forEach(doc => {
      complaintsList.push(doc.data());
    });

    // 5. Optional runInstanceId filter (client-side to avoid composite indexes)
    if (runInstanceId) {
      complaintsList = complaintsList.filter(c => c.runInstanceId === runInstanceId);
    }

    // 6. Newest complaints first sorting
    complaintsList.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    // 6. Return Clean Response
    res.status(200).json({
      success: true,
      count: complaintsList.length,
      complaints: complaintsList
    });

  } catch (error) {
    console.error('(Fetch Complaints) Error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to fetch complaints list',
      details: error.message
    });
  }
});

// --- Route: Admin Resolves/Closes an Existing Complaint ---
app.patch('/api/obhs/complaints/resolve/:complaintId', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.params;
    const { adminRemarks, resolutionPhotoUrl } = req.body; // Inputs from admin frontend

    if (!complaintId) {
      return res.status(400).json({
        success: false,
        error: "Complaint ID is required in URL parameters."
      });
    }

    const adminUser = req.user || req.userData;
    if (!adminUser || !adminUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. Admin session not found." });
    }

    const userRole = adminUser.role || "Admin";
    const allowedRoles = ['Admin', 'Supervisor', 'Company Master'];

    if (!allowedRoles.includes(userRole)) {
      return res.status(403).json({
        success: false,
        error: `Access Denied. Your role is '${userRole}'. Only authorized system masters or admins can resolve complaints.`
      });
    }

    const complaintRef = db.collection('obhs_complaints').doc(complaintId);
    const complaintDoc = await complaintRef.get();

    if (!complaintDoc.exists) {
      return res.status(404).json({
        success: false,
        error: `No complaint found with ID: ${complaintId}`
      });
    }

    const complaintData = complaintDoc.data();

    if (complaintData.status === 'CLOSED') {
      return res.status(400).json({
        success: false,
        error: "This complaint has already been resolved and CLOSED."
      });
    }

    const updateData = {
      status: "CLOSED",
      updatedAt: new Date().toISOString(),

      // Resolution Auditing Metadata
      resolutionDetails: {
        resolvedAt: new Date().toISOString(),
        adminRemarks: adminRemarks || "Resolved by system administrator.",
        resolutionPhotoUrl: resolutionPhotoUrl || null,
        resolvedBy: {
          uid: adminUser.uid,
          name: adminUser.fullName || "System Admin",
          role: userRole
        }
      }
    };

    // 6. Update Firestore using merge mechanism
    await complaintRef.update(updateData);

    console.log(`(Complaint Resolve) ID ${complaintId} successfully CLOSED by Admin: ${adminUser.fullName}`);

    // 7. Success Response back to Admin Portal
    res.status(200).json({
      success: true,
      message: "Complaint status updated to CLOSED successfully.",
      complaintId: complaintId,
      updatedFields: updateData
    });

  } catch (error) {
    console.error('(Complaint Resolve) Error updating status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to resolve/close complaint',
      details: error.message
    });
  }
});

// --- Route: Submit Passenger Feedback ---
app.post('/api/obhs/feedback/passenger', verifyToken, async (req, res) => {
  try {
    const {
      passengerName,
      pnrNumber,
      mobileNumber,
      coachNo,
      ratings,
      remarks,
      photoUrl,
      runInstanceId: bodyRunInstanceId,
      workerId: bodyWorkerId
    } = req.body;

    if (!passengerName || !pnrNumber || !mobileNumber || !coachNo || !ratings) {
      return res.status(400).json({
        success: false,
        error: "Passenger name, PNR, mobile, coach, and ratings are required fields."
      });
    }

    const { cleanliness, toiletHygiene, linenQuality, security, staffBehaviour } = ratings;
    if (
      cleanliness === undefined ||
      toiletHygiene === undefined ||
      linenQuality === undefined ||
      security === undefined ||
      staffBehaviour === undefined
    ) {
      return res.status(400).json({
        success: false,
        error: "All 5 rating parameters (cleanliness, toiletHygiene, linenQuality, security, staffBehaviour) must be provided."
      });
    }

    const workerUser = req.user || req.userData;
    if (!workerUser || !workerUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. Worker session not found." });
    }

    const workerId = workerUser.uid;
    const workerName = workerUser.fullName || "OBHS Worker";
    let runInstanceId = bodyRunInstanceId || workerUser.activeRunInstanceId;

    if (!runInstanceId) {
      const runInstanceSnapshot = await db.collection('RunInstance')
        .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running'])
        .get();

      for (const doc of runInstanceSnapshot.docs) {
        const runData = doc.data();
        if (runData.coaches && Array.isArray(runData.coaches)) {
          const isWorkerAssigned = runData.coaches.some(coach => coach.workerId === workerId);
          if (isWorkerAssigned) {
            runInstanceId = runData.runInstanceId || doc.id;
            break;
          }
        }
      }
    }

    if (!runInstanceId) {
      return res.status(400).json({
        success: false,
        error: "No active train journey (RunInstance) found for this worker to log feedback."
      });
    }

    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) {
      return res.status(404).json({ success: false, error: "Active RunInstance details not found." });
    }
    const runData = runDoc.data();
    const trainNo = runData.trainNo || "UNKNOWN";
    const trainName = runData.trainName || "";

    const totalStars =
      Number(cleanliness) +
      Number(toiletHygiene) +
      Number(linenQuality) +
      Number(security) +
      Number(staffBehaviour);

    const overallRating = parseFloat((totalStars / 5).toFixed(2)); // e.g., 4.20

    const feedbackRef = db.collection('obhs_feedbacks').doc(); // Auto ID

    const feedbackData = {
      feedbackId: feedbackRef.id,
      feedbackType: "PASSENGER",
      runInstanceId: runInstanceId,
      trainNo: trainNo,
      trainName: trainName,
      coachNo: coachNo,

      passengerName: passengerName.trim(),
      pnrNumber: pnrNumber.trim(),
      mobileNumber: mobileNumber.trim(),
      remarks: remarks || "",
      photoUrl: photoUrl || null,

      ratings: {
        cleanliness: Number(cleanliness),
        toiletHygiene: Number(toiletHygiene),
        linenQuality: Number(linenQuality),
        security: Number(security),
        staffBehaviour: Number(staffBehaviour)
      },
      overallRating: overallRating,

      date: new Date().toISOString().split('T')[0],
      createdAt: new Date().toISOString(),
      collectedBy: {
        uid: workerId,
        name: workerName,
        role: workerUser.role || "Railway Worker"
      }
    };

    await feedbackRef.set(feedbackData);

    console.log(`(Feedback) Successfully saved Passenger Feedback ID: ${feedbackRef.id} (Overall Score: ${overallRating})`);

    res.status(201).json({
      success: true,
      message: "Passenger feedback submitted successfully",
      feedbackId: feedbackRef.id,
      overallRating: overallRating,
      data: feedbackData
    });

  } catch (error) {
    console.error('(Feedback) Error while submitting passenger feedback:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to submit feedback',
      details: error.message
    });
  }
});

// --- Public Route: Submit Passenger Feedback via QR ---
app.post('/api/public/passenger/submit-feedback', async (req, res) => {
  try {
    const {
      passengerName,
      mobileNumber,
      coachNo,
      ratings,
      remarks,
      runInstanceId
    } = req.body;

    if (!runInstanceId || !coachNo || !ratings) {
      return res.status(400).json({
        success: false,
        error: "Run Instance ID, Coach No, and Ratings are required."
      });
    }

    const { cleanliness, toiletHygiene, linenQuality, security, staffBehaviour } = ratings;
    if (
      cleanliness === undefined ||
      toiletHygiene === undefined ||
      linenQuality === undefined ||
      security === undefined ||
      staffBehaviour === undefined
    ) {
      return res.status(400).json({
        success: false,
        error: "All 5 rating parameters must be provided."
      });
    }

    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) {
      return res.status(404).json({ success: false, error: "Journey not found." });
    }
    const runData = runDoc.data();

    const totalStars =
      Number(cleanliness) +
      Number(toiletHygiene) +
      Number(linenQuality) +
      Number(security) +
      Number(staffBehaviour);

    const overallRating = parseFloat((totalStars / 5).toFixed(2));

    const feedbackRef = db.collection('obhs_feedbacks').doc();

    const feedbackData = {
      feedbackId: feedbackRef.id,
      feedbackType: "QR_PASSENGER",
      runInstanceId: runInstanceId,
      trainNo: runData.trainNo || "UNKNOWN",
      trainName: runData.trainName || "",
      coachNo: coachNo,
      passengerName: passengerName || "Anonymous",
      mobileNumber: mobileNumber || "N/A",
      remarks: remarks || "",
      ratings: {
        cleanliness: Number(cleanliness),
        toiletHygiene: Number(toiletHygiene),
        linenQuality: Number(linenQuality),
        security: Number(security),
        staffBehaviour: Number(staffBehaviour)
      },
      overallRating: overallRating,
      source: "QR_CODE",
      createdAt: new Date().toISOString(),
      timestamp: Date.now()
    };

    await feedbackRef.set(feedbackData);

    res.status(201).json({
      success: true,
      message: "Thank you for your valuable feedback!",
      overallRating: overallRating
    });

  } catch (error) {
    console.error('(Public Feedback) Error:', error);
    res.status(500).json({ success: false, error: 'Failed to record feedback' });
  }
});

// --- Route: Submit Official Inspection Feedback ---
// Shared collection 'obhs_feedbacks' mein official evaluation data save karta hai
app.post('/api/obhs/feedback/official', verifyToken, async (req, res) => {
  try {
    const {
      inspectorName,
      isRandomInspection, // Expecting boolean (true/false) or string ("Yes"/"No")
      workerId,           // Unique UID of the ground worker being inspected
      workerName,         // Name of the ground worker being inspected
      coachNo,            // e.g., "S2", "A1"
      raterType,          // 'Official', 'TTE', 'PSME', 'Supervisor/Admin'
      ratings,            // Object containing the 5 parameters
      remarks,            // Optional
      photoUrl            // Optional
    } = req.body;

    // 1. Input Validation for Official Schema
    if (!inspectorName || isRandomInspection === undefined || !workerId || !workerName || !coachNo || !ratings) {
      return res.status(400).json({
        success: false,
        error: "Inspector name, random inspection toggle, worker ID, worker name, coach, and ratings are required fields."
      });
    }

    // 5 Star Rating Parameters Validation
    const { cleanliness, toiletHygiene, linenQuality, security, staffBehaviour } = ratings;
    if (
      cleanliness === undefined ||
      toiletHygiene === undefined ||
      linenQuality === undefined ||
      security === undefined ||
      staffBehaviour === undefined
    ) {
      return res.status(400).json({
        success: false,
        error: "All 5 rating parameters (cleanliness, toiletHygiene, linenQuality, security, staffBehaviour) must be provided."
      });
    }

    // 2. Extract Submitting User Context from Token
    const currentSessionUser = req.user || req.userData;
    if (!currentSessionUser || !currentSessionUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. User session not found." });
    }

    const submittedByUid = currentSessionUser.uid;
    const submittedByName = currentSessionUser.fullName || "Official";
    const userRole = currentSessionUser.role || "";

    // 3. ROLE-BASED ACCESS CONTROL (Security Guard)
    // Sirf Company Master, Admin, ya Supervisor hi official feedback de sakte hain
    const allowedOfficialRoles = ['Company Master', 'Admin', 'Supervisor'];
    if (!allowedOfficialRoles.includes(userRole)) {
      return res.status(403).json({
        success: false,
        error: `Access Denied. Your role is '${userRole}'. Only Company Masters, Admins, or Supervisors can submit official feedback.`
      });
    }

    let runInstanceId = req.body.runInstanceId || currentSessionUser.activeRunInstanceId;

    // 4. SMART RUNINSTANCE FALLBACK (Flexible Status Mapping)
    if (!runInstanceId) {
      const runInstanceSnapshot = await db.collection('RunInstance')
        .where('status', 'in', ['PLANNED', 'ALLOCATED', 'READY', 'Active', 'ACTIVE', 'active', 'Scheduled', 'scheduled', 'Running', 'running', 'Completed'])
        .get();

      for (const doc of runInstanceSnapshot.docs) {
        const runData = doc.data();
        if (runData.coaches && Array.isArray(runData.coaches)) {
          // Explicitly target worker assignment to find the trip context
          const isWorkerAssigned = runData.coaches.some(coach => coach.workerId === workerId.trim());
          if (isWorkerAssigned) {
            runInstanceId = runData.runInstanceId || doc.id;
            break;
          }
        }
      }
    }

    if (!runInstanceId) {
      return res.status(400).json({
        success: false,
        error: `No train journey (RunInstance) found matching the worker ID: ${workerId}. Please verify if the worker was assigned to any journey.`
      });
    }

    // 5. Fetch Train Metadata from RunInstance
    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) {
      return res.status(404).json({ success: false, error: "Active/Historical RunInstance details not found." });
    }
    const runData = runDoc.data();
    const trainNo = runData.trainNo || "UNKNOWN";
    const trainName = runData.trainName || "";

    // 6. Calculate Auto Overall Rating
    const totalStars =
      Number(cleanliness) +
      Number(toiletHygiene) +
      Number(linenQuality) +
      Number(security) +
      Number(staffBehaviour);

    const overallRating = parseFloat((totalStars / 5).toFixed(2));

    // 7. Prepare Official Feedback Document Structure
    const feedbackRef = db.collection('obhs_feedbacks').doc();

    const feedbackData = {
      feedbackId: feedbackRef.id,
      feedbackType: "OFFICIAL",
      runInstanceId: runInstanceId,
      trainNo: trainNo,
      trainName: trainName,
      coachNo: coachNo,
      raterType: raterType || 'Official',

      // Official Specific Fields
      inspectorName: inspectorName.trim(),
      isRandomInspection: isRandomInspection,
      remarks: remarks || "",
      photoUrl: photoUrl || null,

      // Target Worker Mapping Structure (Dashboard Analytics Engine Link)
      targetWorker: {
        uid: workerId.trim(),
        name: workerName.trim()
      },

      // Ratings Framework
      ratings: {
        cleanliness: Number(cleanliness),
        toiletHygiene: Number(toiletHygiene),
        linenQuality: Number(linenQuality),
        security: Number(security),
        staffBehaviour: Number(staffBehaviour)
      },
      overallRating: overallRating,

      // Metadata Metrics
      date: new Date().toISOString().split('T')[0],
      createdAt: new Date().toISOString(),
      collectedBy: {
        uid: submittedByUid,
        name: submittedByName,
        role: userRole
      }
    };

    // 8. Execute Database Save
    await feedbackRef.set(feedbackData);

    console.log(`(Official Feedback) Successfully created by ${submittedByName} (${userRole}) for Worker: ${workerId}`);

    // 9. Send Unified Response
    res.status(201).json({
      success: true,
      message: "Official inspection feedback submitted successfully",
      feedbackId: feedbackRef.id,
      overallRating: overallRating,
      data: feedbackData
    });

  } catch (error) {
    console.error('(Official Feedback) Submission Exception:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to submit official feedback',
      details: error.message
    });
  }
});
// --- Route: Admin Global Dashboard Workers Performance Analytics ---
app.get('/api/admin/analytics/workers-performance', verifyToken, async (req, res) => {
  try {
    // 1. Admin/Master Security Role Verification
    const adminUser = req.user || req.userData;
    if (!adminUser || !adminUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. Session expired." });
    }

    const allowedRoles = ['Admin', 'Railway Supervisor', 'Company Master'];
    if (!allowedRoles.includes(adminUser.role)) {
      return res.status(403).json({ success: false, error: "Access Denied. Admin privilege required." });
    }

    // 2. Fetch only approved Railway Workers from 'users' collection (FIXED: Added role filter)
    const workersSnapshot = await db.collection('users')
      .where('status', '==', 'APPROVED')
      .where('role', '==', 'Railway Worker')
      .get();

    if (workersSnapshot.empty) {
      return res.status(200).json({ success: true, message: "No active workers found.", data: [] });
    }

    // Ek map banate hain fast lookup aur accumulator holding ke liye
    let workerPerformanceMap = {};

    workersSnapshot.docs.forEach(doc => {
      const uData = doc.data();
      if (uData.uid) {
        workerPerformanceMap[uData.uid] = {
          workerId: uData.uid,
          workerName: uData.fullName || "Unknown Worker",
          email: uData.email || "",
          designation: uData.designation || "OBHS Staff",

          // Passenger Summary Metrics
          passengerFeedbackCount: 0,
          passengerSumRating: 0,
          passengerAvgRating: 0.0,

          // Official Summary Metrics
          officialFeedbackCount: 0,
          officialSumRating: 0,
          officialAvgRating: 0.0,

          // Combined Ultimate Core Metric
          combinedOverallRating: 0.0,
          totalFeedbackCount: 0,

          // Granular Parameters Breakdown (Combined)
          parametersBreakdown: {
            cleanlinessSum: 0,
            toiletHygieneSum: 0,
            linenQualitySum: 0,
            securitySum: 0,
            staffBehaviourSum: 0
          }
        };
      }
    });

    // 3. Fetch ALL feedbacks from 'obhs_feedbacks' collection
    const feedbackSnapshot = await db.collection('obhs_feedbacks').get();

    // 4. Map-Reduce Processing Loop (Grouping Data Worker ID wise)
    feedbackSnapshot.docs.forEach(doc => {
      const fData = doc.data();
      if (!fData) return; // Safe check for empty docs

      let assignedWorkerId = null;

      // LOGIC BASED ON FEEDBACK TYPE
      if (fData.feedbackType === 'OFFICIAL') {
        assignedWorkerId = fData.targetWorker?.uid;
      } else {
        assignedWorkerId = fData.collectedBy?.uid;
      }

      // Agar yeh worker humare active map mein exist karta hai, toh metrics update karo
      if (assignedWorkerId && workerPerformanceMap[assignedWorkerId]) {
        let workerNode = workerPerformanceMap[assignedWorkerId];
        const overallScore = Number(fData.overallRating || 0);

        workerNode.totalFeedbackCount += 1;

        // Categorize by Feedback Type
        if (fData.feedbackType === 'OFFICIAL') {
          workerNode.officialFeedbackCount += 1;
          workerNode.officialSumRating += overallScore;
        } else {
          workerNode.passengerFeedbackCount += 1;
          workerNode.passengerSumRating += overallScore;
        }

        // Accumulate parameters for overall breakdown (FIXED: Safe property checks)
        if (fData.ratings) {
          workerNode.parametersBreakdown.cleanlinessSum += Number(fData.ratings.cleanliness || 0);
          workerNode.parametersBreakdown.toiletHygieneSum += Number(fData.ratings.toiletHygiene || 0);
          workerNode.parametersBreakdown.linenQualitySum += Number(fData.ratings.linenQuality || 0);
          workerNode.parametersBreakdown.securitySum += Number(fData.ratings.security || 0);
          workerNode.parametersBreakdown.staffBehaviourSum += Number(fData.ratings.staffBehaviour || 0);
        }
      }
    });

    // 5. Final Mathematical Average Compilation
    let finalPerformanceList = Object.values(workerPerformanceMap).map(worker => {

      // Passenger Average
      if (worker.passengerFeedbackCount > 0) {
        worker.passengerAvgRating = parseFloat((worker.passengerSumRating / worker.passengerFeedbackCount).toFixed(2));
      }

      // Official Average
      if (worker.officialFeedbackCount > 0) {
        worker.officialAvgRating = parseFloat((worker.officialSumRating / worker.officialFeedbackCount).toFixed(2));
      }

      // Combined Overall Average (Passenger + Official standard mean weight)
      if (worker.totalFeedbackCount > 0) {
        const grandSum = worker.passengerSumRating + worker.officialSumRating;
        worker.combinedOverallRating = parseFloat((grandSum / worker.totalFeedbackCount).toFixed(2));

        // Parameter breakdown averages conversion
        worker.parameters = {
          cleanliness: parseFloat((worker.parametersBreakdown.cleanlinessSum / worker.totalFeedbackCount).toFixed(2)),
          toiletHygiene: parseFloat((worker.parametersBreakdown.toiletHygieneSum / worker.totalFeedbackCount).toFixed(2)),
          linenQuality: parseFloat((worker.parametersBreakdown.linenQualitySum / worker.totalFeedbackCount).toFixed(2)),
          security: parseFloat((worker.parametersBreakdown.securitySum / worker.totalFeedbackCount).toFixed(2)),
          staffBehaviour: parseFloat((worker.parametersBreakdown.staffBehaviourSum / worker.totalFeedbackCount).toFixed(2))
        };
      } else {
        worker.parameters = { cleanliness: 0, toiletHygiene: 0, linenQuality: 0, security: 0, staffBehaviour: 0 };
      }

      // Safely delete raw accumulator sums before sending response
      delete worker.passengerSumRating;
      delete worker.officialSumRating;
      delete worker.parametersBreakdown;

      return worker;
    });

    // 6. Sort workers list: Highest Rated workers on top
    finalPerformanceList.sort((a, b) => b.combinedOverallRating - a.combinedOverallRating);

    // 7. Send Response
    res.status(200).json({
      success: true,
      totalWorkersEvaluated: finalPerformanceList.length,
      data: finalPerformanceList
    });

  } catch (error) {
    console.error('(Admin Analytics) Critical Error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to aggregate global workers performance data',
      details: error.message
    });
  }
});

// --- Route: Get Logged-in Worker's Feedback Summary & Analytics ---
app.get('/api/obhs/feedback/worker-summary', verifyToken, async (req, res) => {
  try {
    // 1. Get Logged-in Worker UID from Token Context
    const workerUser = req.user || req.userData;
    if (!workerUser || !workerUser.uid) {
      return res.status(401).json({ success: false, error: "Unauthorized. Worker session not found." });
    }

    const workerId = workerUser.uid;

    // 2. Fetch all feedbacks collected by this specific worker
    const feedbackSnapshot = await db.collection('obhs_feedbacks')
      .where('collectedBy.uid', '==', workerId)
      .get();

    // If no feedbacks found yet
    if (feedbackSnapshot.empty) {
      return res.status(200).json({
        success: true,
        message: "No feedback data recorded yet for this worker.",
        analytics: {
          overallRating: 0.0,
          totalFeedbacksCount: 0,
          parametersBreakdown: {
            cleanliness: 0.0,
            toiletHygiene: 0.0,
            linenQuality: 0.0,
            security: 0.0,
            staffBehaviour: 0.0
          }
        },
        recentFeedbacks: []
      });
    }

    // 3. Variables to accumulate ratings for calculating averages
    let totalFeedbacksCount = 0;
    let sumOverallRating = 0;

    let sumCleanliness = 0;
    let sumToiletHygiene = 0;
    let sumLinenQuality = 0;
    let sumSecurity = 0;
    let sumStaffBehaviour = 0;

    let allFeedbacksList = [];

    // 4. Process all feedback documents
    feedbackSnapshot.docs.forEach(doc => {
      const data = doc.data();
      totalFeedbacksCount++;

      // Sum values for overall metrics
      sumOverallRating += Number(data.overallRating || 0);

      // Sum individual 5 star rating parameters
      if (data.ratings) {
        sumCleanliness += Number(data.ratings.cleanliness || 0);
        sumToiletHygiene += Number(data.ratings.toiletHygiene || 0);
        sumLinenQuality += Number(data.ratings.linenQuality || 0);
        sumSecurity += Number(data.ratings.security || 0);
        sumStaffBehaviour += Number(data.ratings.staffBehaviour || 0);
      }

      // Add relevant clean snippet into list for sorting recent ones later
      allFeedbacksList.push({
        passengerName: data.passengerName || "Anonymous Passenger",
        overallRating: data.overallRating || 0,
        remarks: data.remarks || "",
        createdAt: data.createdAt || new Date().toISOString()
      });
    });

    // 5. Calculate Final Averages (Rounded to 2 decimal points)
    const avgOverall = parseFloat((sumOverallRating / totalFeedbacksCount).toFixed(2));
    const avgCleanliness = parseFloat((sumCleanliness / totalFeedbacksCount).toFixed(2));
    const avgToiletHygiene = parseFloat((sumToiletHygiene / totalFeedbacksCount).toFixed(2));
    const avgLinenQuality = parseFloat((sumLinenQuality / totalFeedbacksCount).toFixed(2));
    const avgSecurity = parseFloat((sumSecurity / totalFeedbacksCount).toFixed(2));
    const avgStaffBehaviour = parseFloat((sumStaffBehaviour / totalFeedbacksCount).toFixed(2));

    // 6. Sort by date and slice the top 5 most recent feedbacks
    // Newest first order execution
    const recent5Feedbacks = allFeedbacksList
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
      .slice(0, 5)
      .map(item => ({
        passengerName: item.passengerName,
        overallRating: item.overallRating,
        remarks: item.remarks
      }));

    // 7. Return Structured Response for Profile/Analytics Screen
    res.status(200).json({
      success: true,
      analytics: {
        overallRating: avgOverall,
        totalFeedbacksCount: totalFeedbacksCount,
        parametersBreakdown: {
          cleanliness: avgCleanliness,
          toiletHygiene: avgToiletHygiene,
          linenQuality: avgLinenQuality,
          security: avgSecurity,
          staffBehaviour: avgStaffBehaviour
        }
      },
      recentFeedbacks: recent5Feedbacks
    });

  } catch (error) {
    console.error('(Worker Summary) Error fetching analytics:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to fetch worker feedback analytics summary',
      details: error.message
    });
  }
});


// =======================================================
// == 10. BILLING MANAGEMENT APIs
// =======================================================

const billingEngine = {
  calculatePerformanceDeductionPct(overallScore, rule) {
    if (overallScore >= 90) return rule.penaltyScore90Plus || 0;
    if (overallScore >= 80) return rule.penaltyScore80To89 || 2;
    if (overallScore >= 70) return rule.penaltyScore70To79 || 5;
    return rule.penaltyScoreBelow70 || 10;
  },

  calculateGrade(score) {
    if (score >= 90) return 'A';
    if (score >= 80) return 'B';
    if (score >= 70) return 'C';
    return 'D';
  },

  generateBill(rule, month, year, overallScore, generatedBy, scoreBreakdown, machineShortageCount, manpowerShortageCount, missedObhsCount, otherPenalties) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const period = `${months[month - 1]} ${year}`;
    const grade = billingEngine.calculateGrade(overallScore);
    const perfDedPct = billingEngine.calculatePerformanceDeductionPct(overallScore, rule);
    const perfDedAmount = rule.contractValue * perfDedPct / 100;
    const machineDed = (machineShortageCount || 0) * (rule.machineShortagePenalty || 1000);
    const manpowerDed = (manpowerShortageCount || 0) * (rule.manpowerShortagePenalty || 500);
    const obhsDed = (missedObhsCount || 0) * (rule.missedObhsComplaintPenalty || 2000);
    const otherDed = otherPenalties || 0;
    const totalDed = perfDedAmount + machineDed + manpowerDed + obhsDed + otherDed;
    const finalPayable = Math.max(0, rule.contractValue - totalDed);

    const deductions = [];
    if (perfDedAmount > 0) deductions.push({ type: 'Performance', description: `Performance deduction at ${perfDedPct}% (Score: ${overallScore}%)`, count: 1, rate: perfDedPct, amount: perfDedAmount });
    if (machineDed > 0) deductions.push({ type: 'Machine Shortage', description: 'Machine shortage penalty', count: machineShortageCount, rate: rule.machineShortagePenalty, amount: machineDed });
    if (manpowerDed > 0) deductions.push({ type: 'Manpower Shortage', description: 'Manpower shortage penalty', count: manpowerShortageCount, rate: rule.manpowerShortagePenalty, amount: manpowerDed });
    if (obhsDed > 0) deductions.push({ type: 'Missed OBHS Complaint', description: 'Missed OBHS complaint penalty', count: missedObhsCount, rate: rule.missedObhsComplaintPenalty, amount: obhsDed });
    if (otherDed > 0) deductions.push({ type: 'Other Penalties', description: 'Other applicable penalties', count: 1, rate: otherDed, amount: otherDed });

    return {
      billingRuleId: rule.uid,
      contractId: rule.contractId,
      contractNumber: rule.contractNumber,
      entityId: rule.entityId,
      entityName: rule.entityName,
      division: rule.division,
      zone: rule.zone,
      period, month, year,
      contractValue: rule.contractValue,
      overallScore, grade,
      performanceDeductionPct: perfDedPct,
      performanceDeductionAmount: perfDedAmount,
      machineShortageCount: machineShortageCount || 0,
      machineDeduction: machineDed,
      manpowerShortageCount: manpowerShortageCount || 0,
      manpowerDeduction: manpowerDed,
      missedObhsCount: missedObhsCount || 0,
      obhsDeduction: obhsDed,
      otherPenalties: otherDed,
      totalDeduction: totalDed,
      finalPayable,
      status: 'PENDING',
      scoreBreakdown: scoreBreakdown || null,
      deductions,
      auditLog: [{
        action: 'GENERATED',
        performedBy: generatedBy,
        performedByName: 'System',
        timestamp: new Date().toISOString(),
        details: `Bill generated for ${period} - Score: ${overallScore}%, Grade: ${grade}`
      }],
      createdAt: new Date().toISOString(),
      generatedBy
    };
  }
};

// 10.1: Save/Update Billing Config
app.post('/api/billing/config', verifyToken, async (req, res) => {
  try {
    const configData = req.body;
    const { uid, fullName } = req.user;

    const existingQuery = await db.collection('billingRules')
      .where('contractId', '==', configData.contractId)
      .limit(1)
      .get();

    let ref;
    if (!existingQuery.empty) {
      ref = db.collection('billingRules').doc(existingQuery.docs[0].id);
      configData.updatedAt = new Date().toISOString();
      configData.updatedBy = uid;
      delete configData.uid;
      await ref.update(configData);
    } else {
      ref = db.collection('billingRules').doc();
      configData.uid = ref.id;
      configData.createdAt = new Date().toISOString();
      configData.createdBy = uid;
      configData.status = 'Active';
      await ref.set(configData);
    }

    console.log(`(Billing) Config saved for contract ${configData.contractId}`);
    res.status(200).send({ message: 'Billing config saved successfully', uid: ref.id });
  } catch (error) {
    console.error('(Billing) Error saving config:', error);
    res.status(500).send({ error: 'Failed to save billing config', details: error.message });
  }
});

// 10.2: Get Billing Config for a Contract
app.get('/api/billing/config/:contractId', verifyToken, async (req, res) => {
  try {
    const { contractId } = req.params;
    const snapshot = await db.collection('billingRules')
      .where('contractId', '==', contractId)
      .limit(1)
      .get();

    if (snapshot.empty) {
      return res.status(404).send({ message: 'No billing config found for this contract' });
    }

    res.status(200).send({ config: snapshot.docs[0].data() });
  } catch (error) {
    console.error('(Billing) Error fetching config:', error);
    res.status(500).send({ error: 'Failed to fetch billing config', details: error.message });
  }
});

// 10.3: Get All Billing Configs
app.get('/api/billing/config', verifyToken, async (req, res) => {
  try {
    const { role, zone, division, entityId, userType } = req.user;
    let query = db.collection('billingRules');
    const userRole = (role || '').toLowerCase();

    if (userType === 'contractor') {
      if (!entityId) return res.status(403).send({ error: 'Entity ID missing.' });
      query = query.where('entityId', '==', entityId);
    } else if (userRole === 'railway master') {
      query = query.where('zone', '==', zone);
    } else if (userRole.includes('admin') || userRole.includes('supervisor')) {
      query = query.where('division', '==', division);
    }

    const snapshot = await query.get();
    const configs = [];
    snapshot.forEach(doc => configs.push(doc.data()));

    res.status(200).json({ count: configs.length, configs });
  } catch (error) {
    console.error('(Billing) Error fetching configs:', error);
    res.status(500).send({ error: 'Failed to fetch billing configs', details: error.message });
  }
});

// 10.4: Generate Bill
app.post('/api/billing/generate', verifyToken, async (req, res) => {
  try {
    const { contractId, month, year, overallScore, scoreBreakdown, machineShortageCount, manpowerShortageCount, missedObhsCount, otherPenalties } = req.body;
    const { uid, fullName } = req.user;

    const ruleSnapshot = await db.collection('billingRules')
      .where('contractId', '==', contractId)
      .limit(1)
      .get();

    if (ruleSnapshot.empty) {
      return res.status(400).send({ error: 'No billing config found for this contract. Please configure billing rules first.' });
    }

    const rule = ruleSnapshot.docs[0].data();
    rule.uid = ruleSnapshot.docs[0].id;

    const billData = billingEngine.generateBill(rule, month, year, overallScore, uid, scoreBreakdown, machineShortageCount, manpowerShortageCount, missedObhsCount, otherPenalties);

    const ref = db.collection('billingReports').doc();
    billData.uid = ref.id;
    await ref.set(billData);

    console.log(`(Billing) Bill ${ref.id} generated for contract ${contractId} - ${billData.period}`);

    // Notification: Bill generated
    notifyContractAdmins(contractId, 'Bill Generated', `Bill for ${billData.period} (₹${(billData.finalPayable || 0).toLocaleString()}) is ready for review.`, 'bill_generated', { billId: ref.id, contractId, period: billData.period, amount: billData.finalPayable });

    res.status(201).send({ message: 'Bill generated successfully', uid: ref.id, report: billData });
  } catch (error) {
    console.error('(Billing) Error generating bill:', error);
    res.status(500).send({ error: 'Failed to generate bill' });
  }
});

// 10.5: Get Billing Reports
app.get('/api/billing/reports', verifyToken, async (req, res) => {
  try {
    const { status, contractId, entityId: filterEntityId, division, zone, month, year } = req.query;
    const { role, zone: userZone, division: userDivision, entityId: userEntityId, userType } = req.user;
    const userRole = (role || '').toLowerCase();

    let query = db.collection('billingReports');

    if (userType === 'contractor') {
      query = query.where('entityId', '==', userEntityId);
    } else if (userRole === 'railway master') {
      query = query.where('zone', '==', userZone);
    } else if (userRole.includes('admin') || userRole.includes('supervisor')) {
      query = query.where('division', '==', userDivision);
    }

    if (status) query = query.where('status', '==', status);
    if (contractId) query = query.where('contractId', '==', contractId);
    if (filterEntityId) query = query.where('entityId', '==', filterEntityId);
    if (division) query = query.where('division', '==', division);
    if (zone) query = query.where('zone', '==', zone);
    if (month) query = query.where('month', '==', parseInt(month));
    if (year) query = query.where('year', '==', parseInt(year));

    const snapshot = await query.get();
    const reports = [];
    snapshot.forEach(doc => reports.push(doc.data()));
    reports.sort((a, b) => ((b.createdAt || '') > (a.createdAt || '') ? 1 : -1));

    res.status(200).json({ count: reports.length, reports });
  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index required. Check Firebase Console.' });
    }
    console.error('(Billing) Error fetching reports:', error);
    res.status(500).send({ error: 'Failed to fetch billing reports', details: error.message });
  }
});

// 10.6: Get Single Billing Report
app.get('/api/billing/reports/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const doc = await db.collection('billingReports').doc(uid).get();
    if (!doc.exists) return res.status(404).send({ error: 'Billing report not found' });

    res.status(200).json({ report: doc.data() });
  } catch (error) {
    console.error('(Billing) Error fetching report:', error);
    res.status(500).send({ error: 'Failed to fetch billing report', details: error.message });
  }
});

// 10.7: Approve Bill
app.post('/api/billing/approve/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { uid: approverId, fullName } = req.user;

    const ref = db.collection('billingReports').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Billing report not found' });

    const report = doc.data();
    if (report.status !== 'PENDING') {
      return res.status(400).send({ error: `Cannot approve bill with status: ${report.status}` });
    }

    const invoiceNumber = `INV-${report.contractNumber}-${report.year}${String(report.month).padStart(2, '0')}-${Date.now().toString().slice(-4)}`;

    await ref.update({
      status: 'APPROVED',
      approvedBy: approverId,
      approvedByName: fullName,
      approvedAt: new Date().toISOString(),
      invoiceNumber,
      invoiceGeneratedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'APPROVED',
        performedBy: approverId,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Bill approved by ${fullName}. Invoice: ${invoiceNumber}`
      })
    });

    console.log(`(Billing) Bill ${uid} approved by ${fullName}. Invoice: ${invoiceNumber}`);

    // Notification: Bill approved
    notifyContractAdmins(report.contractId, 'Bill Approved', `Bill for ${report.period} has been approved. Invoice: ${invoiceNumber}. Payable: ₹${(report.finalPayable || 0).toLocaleString()}`, 'bill_approved', { billId: uid, contractId: report.contractId, invoiceNumber, amount: report.finalPayable, period: report.period });

    res.status(200).send({ message: 'Bill approved successfully', invoiceNumber });
  } catch (error) {
    console.error('(Billing) Error approving bill:', error);
    res.status(500).send({ error: 'Failed to approve bill', details: error.message });
  }
});

// 10.8: Reject Bill
app.post('/api/billing/reject/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { reason } = req.body;
    const { uid: rejectorId, fullName } = req.user;

    const ref = db.collection('billingReports').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Billing report not found' });

    const report = doc.data();
    if (report.status !== 'PENDING') {
      return res.status(400).send({ error: `Cannot reject bill with status: ${report.status}` });
    }

    await ref.update({
      status: 'REJECTED',
      rejectedBy: rejectorId,
      rejectedByName: fullName,
      rejectedAt: new Date().toISOString(),
      rejectionReason: reason || 'No reason provided',
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'REJECTED',
        performedBy: rejectorId,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Bill rejected by ${fullName}. Reason: ${reason || 'Not specified'}`
      })
    });

    console.log(`(Billing) Bill ${uid} rejected by ${fullName}`);

    // Notification: Bill rejected
    notifyContractAdmins(report.contractId, 'Bill Rejected', `Bill for ${report.period} has been rejected. Reason: ${reason || 'Not specified'}`, 'bill_rejected', { billId: uid, contractId: report.contractId, reason: reason || 'Not specified', period: report.period });

    res.status(200).send({ message: 'Bill rejected successfully' });
  } catch (error) {
    console.error('(Billing) Error rejecting bill:', error);
    res.status(500).send({ error: 'Failed to reject bill', details: error.message });
  }
});

// 10.9: Dashboard Summary
app.get('/api/billing/dashboard', verifyToken, async (req, res) => {
  try {
    const { role, zone, division, entityId, userType } = req.user;
    const userRole = (role || '').toLowerCase();

    let query = db.collection('billingReports');
    if (userType === 'contractor') {
      query = query.where('entityId', '==', entityId);
    } else if (userRole === 'railway master') {
      query = query.where('zone', '==', zone);
    } else if (userRole.includes('admin') || userRole.includes('supervisor')) {
      query = query.where('division', '==', division);
    }

    const snapshot = await query.get();
    const summary = { pendingBills: 0, approvedBills: 0, rejectedBills: 0, totalContractValue: 0, totalDeductions: 0, totalPayable: 0, activeContracts: 0 };

    const contractIds = new Set();
    snapshot.forEach(doc => {
      const d = doc.data();
      if (d.status === 'PENDING') summary.pendingBills++;
      else if (d.status === 'APPROVED') summary.approvedBills++;
      else if (d.status === 'REJECTED') summary.rejectedBills++;
      summary.totalContractValue += d.contractValue || 0;
      summary.totalDeductions += d.totalDeduction || 0;
      summary.totalPayable += d.finalPayable || 0;
      if (d.contractId) contractIds.add(d.contractId);
    });
    summary.activeContracts = contractIds.size;

    res.status(200).json(summary);
  } catch (error) {
    console.error('(Billing) Error fetching dashboard:', error);
    res.status(500).send({ error: 'Failed to fetch dashboard', details: error.message });
  }
});

// 10.10: Generate Invoice
app.post('/api/billing/generate-invoice/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('billingReports').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Billing report not found' });

    const report = doc.data();
    if (report.status !== 'APPROVED') {
      return res.status(400).send({ error: 'Invoice can only be generated for approved bills' });
    }

    const invoiceNumber = report.invoiceNumber || `INV-${report.contractNumber}-${report.year}${String(report.month).padStart(2, '0')}-${Date.now().toString().slice(-4)}`;

    await ref.update({
      invoiceNumber,
      invoiceGeneratedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'INVOICE_GENERATED',
        performedBy: req.user.uid,
        performedByName: req.user.fullName,
        timestamp: new Date().toISOString(),
        details: `Invoice ${invoiceNumber} generated`
      })
    });

    // Notification: Invoice generated
    notifyContractAdmins(report.contractId, 'Invoice Generated', `Invoice ${invoiceNumber} for ${report.period} is ready. Amount: ₹${(report.finalPayable || 0).toLocaleString()}`, 'invoice_generated', { billId: uid, contractId: report.contractId, invoiceNumber, amount: report.finalPayable, period: report.period });

    res.status(200).send({ message: 'Invoice generated successfully', invoiceNumber });
  } catch (error) {
    console.error('(Billing) Error generating invoice:', error);
    res.status(500).send({ error: 'Failed to generate invoice', details: error.message });
  }
});

// 10.11: Contractor Billing Dashboard
app.get('/api/billing/contractor', verifyToken, async (req, res) => {
  try {
    const { entityId, uid } = req.user;
    if (!entityId) return res.status(403).send({ error: 'Entity ID missing.' });

    const ruleSnapshot = await db.collection('billingRules')
      .where('entityId', '==', entityId)
      .get();

    const configs = [];
    ruleSnapshot.forEach(doc => configs.push(doc.data()));

    const reportSnapshot = await db.collection('billingReports')
      .where('entityId', '==', entityId)
      .get();

    const reports = [];
    reportSnapshot.forEach(doc => reports.push(doc.data()));
    reports.sort((a, b) => ((b.createdAt || '') > (a.createdAt || '') ? 1 : -1));
    const recentReports = reports.slice(0, 12);

    let pendingAmount = 0;
    let approvedAmount = 0;
    let totalDeductions = 0;
    reports.forEach(r => {
      totalDeductions += r.totalDeduction || 0;
      if (r.status === 'PENDING') pendingAmount += r.finalPayable || 0;
      if (r.status === 'APPROVED') approvedAmount += r.finalPayable || 0;
    });

    res.status(200).json({
      configs: configs.length,
      totalBills: reports.length,
      pendingAmount,
      approvedAmount,
      totalDeductions,
      recentBills: recentReports.slice(0, 5),
      configList: configs
    });
  } catch (error) {
    console.error('(Billing) Error fetching contractor data:', error);
    res.status(500).send({ error: 'Failed to fetch contractor billing data', details: error.message });
  }
});

// 10.12: Supervisor Billing Dashboard
app.get('/api/billing/supervisor', verifyToken, async (req, res) => {
  try {
    const { uid, division, zone } = req.user;

    const contractsSnapshot = await db.collection('contracts')
      .where('division', '==', division)
      .where('status', '==', 'Active')
      .get();

    const contractIds = [];
    contractsSnapshot.forEach(doc => contractIds.push(doc.id));

    const reportSnapshot = await db.collection('billingReports')
      .where('division', '==', division)
      .get();

    const reports = [];
    reportSnapshot.forEach(doc => reports.push(doc.data()));
    reports.sort((a, b) => ((b.createdAt || '') > (a.createdAt || '') ? 1 : -1));
    const recentReports = reports.slice(0, 20);

    let totalPenalties = 0;
    let pendingCount = 0;
    let approvedCount = 0;
    reports.forEach(r => {
      totalPenalties += r.totalDeduction || 0;
      if (r.status === 'PENDING') pendingCount++;
      if (r.status === 'APPROVED') approvedCount++;
    });

    res.status(200).json({
      activeContracts: contractIds.length,
      totalBills: reports.length,
      pendingCount,
      approvedCount,
      totalPenalties,
      recentBills: reports.slice(0, 5)
    });
  } catch (error) {
    console.error('(Billing) Error fetching supervisor data:', error);
    res.status(500).send({ error: 'Failed to fetch supervisor billing data', details: error.message });
  }
});

//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Section 12: Station Management APIs >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

async function generateStationFormId(stationCode) {
  const prefix = 'SF';
  const code = (stationCode || 'STN').substring(0, 4).toUpperCase();
  const date = new Date();
  const seq = date.getFullYear().toString().slice(-2) + String(date.getMonth() + 1).padStart(2, '0');
  const counterRef = db.collection('counters').doc(`stationForm_${code}_${seq}`);
  const counter = await counterRef.get();
  let nextNum = 1;
  if (counter.exists) { nextNum = (counter.data().value || 0) + 1; }
  await counterRef.set({ value: nextNum }, { merge: true });
  return `${prefix}-${code}-${seq}-${String(nextNum).padStart(4, '0')}`;
}

// 12.1: Create / Update Station
app.post('/api/stations/create', verifyToken, async (req, res) => {
  try {
    const { stationCode, stationName, zone, division, category, stationType, active, latitude, longitude, address } = req.body;
    const { uid } = req.user;
    if (!stationCode || !stationName || !zone || !division) {
      return res.status(400).send({ error: 'stationCode, stationName, zone, division are required' });
    }
    const existing = await db.collection('stations').where('stationCode', '==', stationCode).limit(1).get();
    if (!existing.empty) {
      return res.status(409).send({ error: `Station with code ${stationCode} already exists` });
    }
    const ref = db.collection('stations').doc();
    const data = {
      uid: ref.id, stationCode, stationName, zone, division,
      category: category || 'c', stationType: stationType || 'regular',
      active: active !== false, latitude: latitude || 0, longitude: longitude || 0,
      address: address || '', createdBy: uid,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    };
    await ref.set(data);
    res.status(201).send({ message: 'Station created', uid: ref.id, station: data });
  } catch (error) {
    console.error('(Station) Error creating station:', error);
    res.status(500).send({ error: 'Failed to create station', details: error.message });
  }
});

app.put('/api/stations/update/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('stations').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Station not found' });
    const updates = {};
    const allowedStationFields = ['stationName', 'category', 'stationType', 'active', 'latitude', 'longitude', 'address', 'zone', 'division'];
    for (const key of allowedStationFields) {
      if (req.body[key] !== undefined) updates[key] = req.body[key];
    }
    updates.updatedAt = new Date().toISOString();
    await ref.update(updates);
    res.status(200).send({ message: 'Station updated' });
  } catch (error) {
    res.status(500).send({ error: 'Failed to update station', details: error.message });
  }
});

// 12.2: Get Stations List
app.get('/api/stations/list', verifyToken, async (req, res) => {
  try {
    const { zone, division, category, active } = req.query;
    const { role, zone: userZone, division: userDiv } = req.user;
    const userRole = (role || '').toLowerCase();
    // Fetch all stations and filter in-memory to avoid composite index requirements
    const snapshot = await db.collection('stations').get();
    let stations = [];
    snapshot.forEach(doc => stations.push(doc.data()));
    // Apply filters in-memory
    if (!userRole.includes('master')) {
      const divFilter = division || userDiv;
      if (divFilter) stations = stations.filter(s => s.division === divFilter);
    }
    if (zone) stations = stations.filter(s => s.zone === zone);
    if (category) stations = stations.filter(s => s.category === category);
    if (active !== undefined) stations = stations.filter(s => s.active === (active === 'true'));
    stations.sort((a, b) => (a.stationName || '').localeCompare(b.stationName || ''));
    res.status(200).json({ count: stations.length, stations });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch stations', details: error.message });
  }
});

// 12.3: Station Area CRUD
app.post('/api/station-area/create', verifyToken, async (req, res) => {
  try {
    const { stationId, name, order, description } = req.body;
    if (!stationId || !name) return res.status(400).send({ error: 'stationId and name required' });
    const ref = db.collection('stationAreas').doc();
    const data = { uid: ref.id, stationId, name, order: order || 0, description: description || '', active: true };
    await ref.set(data);
    res.status(201).send({ message: 'Area created', uid: ref.id, area: data });
  } catch (error) {
    res.status(500).send({ error: 'Failed to create area', details: error.message });
  }
});

app.get('/api/station-area/list/:stationId', verifyToken, async (req, res) => {
  try {
    const { stationId } = req.params;
    const snapshot = await db.collection('stationAreas').where('stationId', '==', stationId).get();
    const areas = [];
    snapshot.forEach(doc => areas.push(doc.data()));
    // Sort in-memory to avoid composite index requirement
    areas.sort((a, b) => (a.order || 0) - (b.order || 0));
    res.status(200).json({ count: areas.length, areas });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch areas', details: error.message });
  }
});

// 12.4: Station Zone CRUD
app.post('/api/station-zone/create', verifyToken, async (req, res) => {
  try {
    const { stationId, areaId, areaName, name, description } = req.body;
    if (!stationId || !areaId || !name) return res.status(400).send({ error: 'stationId, areaId, name required' });
    const ref = db.collection('stationZones').doc();
    const data = { uid: ref.id, stationId, areaId, areaName: areaName || '', name, description: description || '', active: true };
    await ref.set(data);
    res.status(201).send({ message: 'Zone created', uid: ref.id, zone: data });
  } catch (error) {
    res.status(500).send({ error: 'Failed to create zone', details: error.message });
  }
});

app.get('/api/station-zone/list/:stationId', verifyToken, async (req, res) => {
  try {
    const { stationId } = req.params;
    const { areaId } = req.query;
    const snapshot = await db.collection('stationZones').where('stationId', '==', stationId).get();
    const zones = [];
    snapshot.forEach(doc => zones.push(doc.data()));
    // Filter by areaId in-memory to avoid composite index requirement
    const filtered = areaId ? zones.filter(z => z.areaId === areaId) : zones;
    res.status(200).json({ count: filtered.length, zones: filtered });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch zones', details: error.message });
  }
});

// 12.5: Contractor Mapping
app.post('/api/station-contractor/map', verifyToken, async (req, res) => {
  try {
    const { stationId, areaId, zoneId, entityId, entityName, serviceType } = req.body;
    if (!stationId || !entityId) return res.status(400).send({ error: 'stationId and entityId required' });
    const ref = db.collection('stationContractors').doc();
    const data = { uid: ref.id, stationId, areaId: areaId || '', zoneId: zoneId || '', entityId, entityName: entityName || '', serviceType: serviceType || 'Station Cleaning', startDate: new Date().toISOString(), active: true };
    await ref.set(data);
    res.status(201).send({ message: 'Contractor mapped', uid: ref.id });
  } catch (error) {
    res.status(500).send({ error: 'Failed to map contractor', details: error.message });
  }
});

app.get('/api/station-contractor/list/:stationId', verifyToken, async (req, res) => {
  try {
    const { stationId } = req.params;
    const snapshot = await db.collection('stationContractors').where('stationId', '==', stationId).get();
    const mappings = [];
    snapshot.forEach(doc => mappings.push(doc.data()));
    res.status(200).json({ count: mappings.length, mappings });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch mappings', details: error.message });
  }
});

// 12.6: Cleaning Schedule
app.post('/api/station-schedule/create', verifyToken, async (req, res) => {
  try {
    const { stationId, areaId, zoneId, frequency, shift, entityId, entityName, supervisorId, supervisorName, startTime, endTime, daysOfWeek } = req.body;
    if (!stationId) return res.status(400).send({ error: 'stationId required' });
    const ref = db.collection('stationSchedules').doc();
    const data = { uid: ref.id, stationId, areaId: areaId || '', zoneId: zoneId || '', frequency: frequency || 'daily', shift: shift || 'Morning', entityId: entityId || '', entityName: entityName || '', supervisorId: supervisorId || '', supervisorName: supervisorName || '', startTime: startTime || '', endTime: endTime || '', daysOfWeek: daysOfWeek || [], active: true, createdAt: new Date().toISOString() };
    await ref.set(data);
    res.status(201).send({ message: 'Schedule created', uid: ref.id });
  } catch (error) {
    res.status(500).send({ error: 'Failed to create schedule', details: error.message });
  }
});

app.get('/api/station-schedule/list/:stationId', verifyToken, async (req, res) => {
  try {
    const { stationId } = req.params;
    const snapshot = await db.collection('stationSchedules').where('stationId', '==', stationId).get();
    const schedules = [];
    snapshot.forEach(doc => schedules.push(doc.data()));
    res.status(200).json({ count: schedules.length, schedules });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch schedules', details: error.message });
  }
});

// 12.7: Station Cleaning Form - Create
app.post('/api/station-cleaning-form/create', verifyToken, async (req, res) => {
  try {
    const { stationId, stationName, areaId, areaName, zoneId, zoneName, division, depot, contractId, contractNumber, cleaningDate, shift, startTime, endTime, manpowerCount, machineCount, areaCovered, areaUncleaned, garbageCollected, remarks, latitude, longitude, deviceId, gpsAddress, photos, activities } = req.body;
    const { uid, fullName, entityId, entityName } = req.user;
    if (!stationId || !division) return res.status(400).send({ error: 'stationId and division required' });

    // Get station code
    const stationDoc = await db.collection('stations').doc(stationId).get();
    const stationCode = stationDoc.exists ? stationDoc.data().stationCode : 'STN';
    const formId = await generateStationFormId(stationCode);

    const ref = db.collection('stationCleaningForms').doc();
    const data = {
      uid: ref.id, formId, stationId, stationName: stationName || '', areaId: areaId || '', areaName: areaName || '',
      zoneId: zoneId || '', zoneName: zoneName || '', division, depot: depot || '',
      contractId: contractId || '', contractNumber: contractNumber || '',
      entityId: entityId || '', entityName: entityName || '',
      submittedBy: uid, submittedByName: fullName,
      status: 'draft',
      cleaningDate: cleaningDate || '', shift: shift || '', startTime: startTime || '', endTime: endTime || '',
      manpowerCount: manpowerCount || 0, machineCount: machineCount || 0,
      areaCovered: areaCovered || 0, areaUncleaned: areaUncleaned || 0, garbageCollected: garbageCollected || 0,
      remarks: remarks || '', latitude: latitude || 0, longitude: longitude || 0,
      deviceId: deviceId || '', gpsAddress: gpsAddress || '',
      photos: photos || [], activities: activities || [],
      auditLog: [{ action: 'CREATED', performedBy: uid, performedByName: fullName, timestamp: new Date().toISOString(), details: `Form ${formId} created` }],
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    };
    await ref.set(data);
    res.status(201).send({ message: 'Station cleaning form created', uid: ref.id, formId });
  } catch (error) {
    console.error('(StationForm) Error creating:', error);
    res.status(500).send({ error: 'Failed to create form', details: error.message });
  }
});

// 12.8: Submit Station Cleaning Form
app.post('/api/station-cleaning-form/submit/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('stationCleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    const form = doc.data();
    if (form.status !== 'draft') return res.status(400).send({ error: 'Only draft forms can be submitted' });

    await ref.update({
      status: 'submitted', updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({ action: 'SUBMITTED', performedBy: req.user.uid, performedByName: req.user.fullName, timestamp: new Date().toISOString(), details: 'Submitted for review' })
    });

    // Notify supervisors
    notifySupervisors(form.division, 'New Station Cleaning Form', `Form ${form.formId} for ${form.stationName} submitted.`, 'station_form_submitted', { stationFormId: uid, formId: form.formId, stationName: form.stationName });

    res.status(200).send({ message: 'Form submitted' });
  } catch (error) {
    res.status(500).send({ error: 'Failed to submit', details: error.message });
  }
});

// 12.9: Approve Station Cleaning Form
app.post('/api/station-cleaning-form/approve/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { uid: approverId, fullName } = req.user;
    const ref = db.collection('stationCleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    const form = doc.data();
    if (form.status !== 'submitted') return res.status(400).send({ error: 'Only submitted forms can be approved' });

    await ref.update({
      status: 'approved', approvedBy: approverId, approvedByName: fullName, approvedAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({ action: 'APPROVED', performedBy: approverId, performedByName: fullName, timestamp: new Date().toISOString(), details: `Approved by ${fullName}` })
    });
    res.status(200).send({ message: 'Form approved' });
  } catch (error) {
    res.status(500).send({ error: 'Failed to approve', details: error.message });
  }
});

// 12.10: Reject Station Cleaning Form
app.post('/api/station-cleaning-form/reject/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { reason } = req.body;
    const { uid: rejectorId, fullName } = req.user;
    const ref = db.collection('stationCleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    const form = doc.data();
    if (form.status !== 'submitted') return res.status(400).send({ error: 'Only submitted forms can be rejected' });

    await ref.update({
      status: 'rejected', rejectedBy: rejectorId, rejectedByName: fullName, rejectedAt: new Date().toISOString(), rejectionReason: reason || 'No reason', updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({ action: 'REJECTED', performedBy: rejectorId, performedByName: fullName, timestamp: new Date().toISOString(), details: `Rejected: ${reason || 'No reason'}` })
    });
    res.status(200).send({ message: 'Form rejected' });
  } catch (error) {
    res.status(500).send({ error: 'Failed to reject', details: error.message });
  }
});

// 12.11: Score Station Cleaning Form (built-in scoring)
app.post('/api/station-cleaning-form/score/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { scoringData, totalScore, grade } = req.body;
    const { uid: scorerId, fullName } = req.user;
    const ref = db.collection('stationCleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    const form = doc.data();
    if (form.status !== 'approved') return res.status(400).send({ error: 'Only approved forms can be scored' });

    const computedGrade = grade || (totalScore >= 90 ? 'A' : totalScore >= 80 ? 'B' : totalScore >= 70 ? 'C' : 'D');

    await ref.update({
      status: 'scored', score: totalScore, grade: computedGrade, scoringData: scoringData || { criteria: [], totalScore, grade: computedGrade },
      scoredBy: scorerId, scoredByName: fullName, scoringAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({ action: 'SCORED', performedBy: scorerId, performedByName: fullName, timestamp: new Date().toISOString(), details: `Score: ${totalScore} (Grade: ${computedGrade})` })
    });
    res.status(200).send({ message: 'Score submitted', grade: computedGrade });
  } catch (error) {
    res.status(500).send({ error: 'Failed to score', details: error.message });
  }
});

// 12.12: Lock Station Cleaning Form
app.post('/api/station-cleaning-form/lock/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('stationCleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    const form = doc.data();
    if (form.status !== 'scored') return res.status(400).send({ error: 'Only scored forms can be locked' });

    await ref.update({
      status: 'locked', lockedAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({ action: 'LOCKED', performedBy: req.user.uid, performedByName: req.user.fullName, timestamp: new Date().toISOString(), details: 'Form locked. Ready for billing.' })
    });
    res.status(200).send({ message: 'Form locked' });
  } catch (error) {
    res.status(500).send({ error: 'Failed to lock', details: error.message });
  }
});

// 12.13: Get Station Cleaning Forms List
app.get('/api/station-cleaning-form/list', verifyToken, async (req, res) => {
  try {
    const { status, stationId, areaId, zoneId, division } = req.query;
    const { role, division: userDiv, entityId, userType } = req.user;
    const snapshot = await db.collection('stationCleaningForms').get();
    let forms = [];
    snapshot.forEach(doc => forms.push(doc.data()));

    // Filter
    if (userType === 'contractor') {
      forms = forms.filter(f => f.entityId === entityId);
    } else if (!(role || '').toLowerCase().includes('master')) {
      forms = forms.filter(f => f.division === userDiv);
    }

    if (status) forms = forms.filter(f => f.status === status);
    if (stationId) forms = forms.filter(f => f.stationId === stationId);
    if (areaId) forms = forms.filter(f => f.areaId === areaId);
    if (zoneId) forms = forms.filter(f => f.zoneId === zoneId);
    if (division) forms = forms.filter(f => f.division === division);

    // Sort by createdAt desc
    forms.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));

    res.status(200).json({ count: forms.length, forms });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch forms', details: error.message });
  }
});

// 12.14: Get Station Cleaning Form Details
app.get('/api/station-cleaning-form/details/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const doc = await db.collection('stationCleaningForms').doc(uid).get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    res.status(200).json({ form: doc.data() });
  } catch (error) {
    res.status(500).send({ error: 'Failed to fetch form', details: error.message });
  }
});

// 12.15: Station Dashboard
app.get('/api/station-dashboard', verifyToken, async (req, res) => {
  try {
    const { role, division: userDiv, zone: userZone, entityId, userType } = req.user;
    const userRole = (role || '').toLowerCase();

    let stationQuery = db.collection('stations');
    let formQuery = db.collection('stationCleaningForms');
    let areaQuery = db.collection('stationAreas');
    let zoneQuery = db.collection('stationZones');
    let schedQuery = db.collection('stationSchedules');
    let contrQuery = db.collection('stationContractors');

    if (!userRole.includes('master')) {
      stationQuery = stationQuery.where('division', '==', userDiv);
      formQuery = formQuery.where('division', '==', userDiv);
    }
    if (userType === 'contractor') {
      formQuery = formQuery.where('entityId', '==', entityId);
    }

    const [stationSnap, formSnap, areaSnap, zoneSnap, schedSnap, contrSnap] = await Promise.all([
      stationQuery.get(), formQuery.get(), areaQuery.get(), zoneQuery.get(), schedQuery.get(), contrQuery.get(),
    ]);

    let totalScore = 0, scoredCount = 0;
    let draftForms = 0, submittedForms = 0, approvedForms = 0, scoredForms = 0, lockedForms = 0, rejectedForms = 0;

    formSnap.forEach(doc => {
      const d = doc.data();
      switch (d.status) {
        case 'draft': draftForms++; break;
        case 'submitted': submittedForms++; break;
        case 'approved': approvedForms++; break;
        case 'scored': scoredForms++; totalScore += d.score || 0; scoredCount++; break;
        case 'locked': lockedForms++; totalScore += d.score || 0; scoredCount++; break;
        case 'rejected': rejectedForms++; break;
      }
    });

    res.status(200).json({
      totalStations: stationSnap.size, activeStations: stationSnap.docs.filter(d => d.data().active !== false).length,
      totalAreas: areaSnap.size, totalZones: zoneSnap.size,
      draftForms, submittedForms, approvedForms, scoredForms, lockedForms, rejectedForms,
      pendingReview: submittedForms,
      schedules: schedSnap.size, contractorMappings: contrSnap.size,
      averageScore: scoredCount > 0 ? Math.round((totalScore / scoredCount) * 100) / 100 : 0,
    });
  } catch (error) {
    console.error('(StationDashboard) Error:', error);
    res.status(500).send({ error: 'Failed to fetch dashboard', details: error.message });
  }
});

// --- Notification Helper ---
async function createNotification({ userId, title, body, type, data, entityId }) {
  try {
    const notifRef = db.collection('notifications').doc();
    await notifRef.set({
      uid: notifRef.id,
      userId: userId || null,
      entityId: entityId || null,
      title,
      body,
      type: type || 'billing',
      data: data || {},
      read: false,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });
    return notifRef.id;
  } catch (e) {
    console.error('(Notification) Error creating notification:', e.message);
  }
}

async function notifyContractAdmins(contractId, title, body, type, data) {
  try {
    const contractDoc = await db.collection('contracts').doc(contractId).get();
    if (!contractDoc.exists) return;
    const contract = contractDoc.data();
    const entityId = contract.entityId || contract.agencyId;
    if (entityId) {
      await createNotification({ entityId, title, body, type, data });
      const userSnapshot = await db.collection('users')
        .where('entityId', '==', entityId)
        .limit(5)
        .get();
      userSnapshot.forEach(doc => {
        createNotification({ userId: doc.id, title, body, type, data });
      });
    }
  } catch (e) {
    console.error('(Notification) Error notifying contract admins:', e.message);
  }
}

// 10.13: Get Notifications for current user
app.get('/api/notifications', verifyToken, async (req, res) => {
  try {
    const { uid, entityId, role, division, zone } = req.user;
    let query = db.collection('notifications');

    if (req.query.all === 'true') {
      // no filter
    } else if (entityId) {
      query = query.where('entityId', '==', entityId);
    } else {
      query = query.where('userId', '==', uid);
    }

    const snapshot = await query.get();
    let notifications = [];
    snapshot.forEach(doc => notifications.push(doc.data()));

    // Sort by createdAt descending (in-memory to avoid composite index requirement)
    notifications.sort((a, b) => {
      const aTime = a.createdAt?.toMillis?.() || a.createdAt || 0;
      const bTime = b.createdAt?.toMillis?.() || b.createdAt || 0;
      return bTime - aTime;
    });
    notifications = notifications.slice(0, 50);

    // Convert Firestore Timestamps to ISO strings for JSON-safe serialization
    const serialized = notifications.map(n => ({
      ...n,
      createdAt: n.createdAt?.toDate?.()?.toISOString() || n.createdAt || null,
    }));

    res.status(200).json({ count: notifications.length, notifications: serialized });
  } catch (error) {
    console.error('(Notification) Error fetching notifications:', error);
    res.status(500).send({ error: 'Failed to fetch notifications', details: error.message });
  }
});

// 10.14: Mark notification as read
app.post('/api/notifications/:uid/read', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    await db.collection('notifications').doc(uid).update({ read: true });
    res.status(200).send({ message: 'Notification marked as read' });
  } catch (error) {
    console.error('(Notification) Error marking notification as read:', error);
    res.status(500).send({ error: 'Failed to mark notification', details: error.message });
  }
});

// 10.15: Mark all notifications as read
app.post('/api/notifications/read-all', verifyToken, async (req, res) => {
  try {
    const { uid, entityId } = req.user;
    let query = db.collection('notifications').where('read', '==', false);
    if (entityId) {
      query = query.where('entityId', '==', entityId);
    } else {
      query = query.where('userId', '==', uid);
    }
    const snapshot = await query.get();
    const batch = db.batch();
    snapshot.forEach(doc => {
      batch.update(doc.ref, { read: true });
    });
    const count = snapshot.size;
    if (count > 0) await batch.commit();
    res.status(200).send({ message: `${count} notifications marked as read` });
  } catch (error) {
    console.error('(Notification) Error marking all as read:', error);
    res.status(500).send({ error: 'Failed to mark notifications', details: error.message });
  }
});

// 10.16: Get unread notification count
app.get('/api/notifications/unread-count', verifyToken, async (req, res) => {
  try {
    const { uid, entityId } = req.user;
    let query = db.collection('notifications').where('read', '==', false);
    if (entityId) {
      query = query.where('entityId', '==', entityId);
    } else {
      query = query.where('userId', '==', uid);
    }
    const snapshot = await query.get();
    let unreadCount = snapshot.size;
    res.status(200).json({ count: unreadCount });
  } catch (error) {
    console.error('(Notification) Error fetching unread count:', error);
    res.status(500).send({ error: 'Failed to fetch unread count', details: error.message });
  }
});

//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Section 11: Cleaning Form APIs >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

async function notifySupervisors(division, title, body, type, data) {
  try {
    const userSnapshot = await db.collection('users')
      .where('division', '==', division)
      .where('role', '==', 'Railway Supervisor')
      .get();
    userSnapshot.forEach(doc => {
      createNotification({ userId: doc.id, title, body, type, data });
    });
  } catch (e) {
    console.error('(CleaningForm) Error notifying supervisors:', e.message);
  }
}

async function generateCleaningFormId(formType, division) {
  const prefix = formType === 'coach' ? 'CF' : 'PF';
  const div = (division || 'GEN').substring(0, 3).toUpperCase();
  const date = new Date();
  const seq = date.getFullYear().toString().slice(-2) + String(date.getMonth() + 1).padStart(2, '0');
  const counterRef = db.collection('counters').doc(`cleaningForm_${prefix}_${seq}`);
  const nextNum = await db.runTransaction(async transaction => {
    const counter = await transaction.get(counterRef);
    const nextVal = (counter.data()?.value || 0) + 1;
    transaction.set(counterRef, { value: nextVal }, { merge: true });
    return nextVal;
  });
  return `${prefix}-${div}-${seq}-${String(nextNum).padStart(4, '0')}`;
}

// 11.1: Create Cleaning Form (Draft or Submit)
app.post('/api/cleaning-form/create', verifyToken, async (req, res) => {
  try {
    const { formType, division, depot, contractId, contractNumber, cleaningDate, cleaningShift, startTime, endTime, manpowerCount, machineCount, remarks, latitude, longitude, deviceId, gpsAddress, coachDetails, premiseDetails, photos } = req.body;
    const { uid, fullName, entityId, entityName } = req.user;

    if (!formType) {
      return res.status(400).send({ error: 'formType is required.' });
    }

    const formId = await generateCleaningFormId(formType, division);
    const ref = db.collection('cleaningForms').doc();

    const formData = {
      uid: ref.id,
      formId,
      formType,
      division,
      depot: depot || '',
      contractId,
      contractNumber: contractNumber || '',
      entityId: entityId || '',
      entityName: entityName || '',
      submittedBy: uid,
      submittedByName: fullName,
      status: 'draft',
      cleaningDate: cleaningDate || '',
      cleaningShift: cleaningShift || '',
      startTime: startTime || '',
      endTime: endTime || '',
      manpowerCount: manpowerCount || 0,
      machineCount: machineCount || 0,
      remarks: remarks || '',
      latitude: latitude || 0,
      longitude: longitude || 0,
      deviceId: deviceId || '',
      gpsAddress: gpsAddress || '',
      photos: photos || [],
      auditLog: [{
        action: 'CREATED',
        performedBy: uid,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Form ${formId} created as draft`
      }],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    if (formType === 'coach' && coachDetails) formData.coachDetails = coachDetails;
    if (formType === 'premise' && premiseDetails) formData.premiseDetails = premiseDetails;

    await ref.set(formData);

    console.log(`(CleaningForm) Created ${formType} form ${formId}`);
    res.status(201).send({ message: 'Form created successfully', uid: ref.id, formId });
  } catch (error) {
    console.error('(CleaningForm) Error creating form:', error);
    res.status(500).send({ error: 'Failed to create form', details: error.message });
  }
});

// 11.2: Save Draft (update existing)
app.put('/api/cleaning-form/save-draft/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'draft') return res.status(400).send({ error: 'Can only save draft for forms in draft status' });

    const allowedFields = ['cleaningDate', 'cleaningShift', 'startTime', 'endTime', 'manpowerCount', 'machineCount', 'remarks', 'latitude', 'longitude', 'deviceId', 'gpsAddress', 'coachDetails', 'premiseDetails', 'photos'];
    const updates = { updatedAt: new Date().toISOString() };
    for (const key of allowedFields) {
      if (req.body[key] !== undefined) updates[key] = req.body[key];
    }

    updates.auditLog = admin.firestore.FieldValue.arrayUnion({
      action: 'DRAFT_SAVED',
      performedBy: req.user.uid,
      performedByName: req.user.fullName,
      timestamp: new Date().toISOString(),
      details: 'Draft saved'
    });

    await ref.update(updates);
    res.status(200).send({ message: 'Draft saved successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error saving draft:', error);
    res.status(500).send({ error: 'Failed to save draft', details: error.message });
  }
});

// 11.3: Submit Form
app.post('/api/cleaning-form/submit/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'draft') return res.status(400).send({ error: 'Only draft forms can be submitted' });

    await ref.update({
      status: 'submitted',
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'SUBMITTED',
        performedBy: req.user.uid,
        performedByName: req.user.fullName,
        timestamp: new Date().toISOString(),
        details: 'Form submitted for review'
      })
    });

    // Notify supervisors in this division
    notifySupervisors(form.division, 'New Cleaning Form', `Form ${form.formId} (${form.formType}) has been submitted for review.`, 'cleaning_form_submitted', { formId: form.formId, cleaningFormId: uid, formType: form.formType, division: form.division });

    console.log(`(CleaningForm) Submitted ${form.formId}`);
    res.status(200).send({ message: 'Form submitted successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error submitting form:', error);
    res.status(500).send({ error: 'Failed to submit form', details: error.message });
  }
});

// 11.4: Approve Form (opens scoring)
app.post('/api/cleaning-form/approve/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { uid: approverId, fullName } = req.user;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'submitted') return res.status(400).send({ error: 'Only submitted forms can be approved' });

    await ref.update({
      status: 'approved',
      approvedBy: approverId,
      approvedByName: fullName,
      approvedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'APPROVED',
        performedBy: approverId,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Form approved by ${fullName}. Scoring section opened.`
      })
    });

    // Notify contractor
    notifyContractAdmins(form.contractId, 'Form Approved', `Form ${form.formId} has been approved. Scoring is now open.`, 'cleaning_form_approved', { cleaningFormId: uid, formId: form.formId });

    console.log(`(CleaningForm) Approved ${form.formId}`);
    res.status(200).send({ message: 'Form approved successfully. Scoring section is now open.' });
  } catch (error) {
    console.error('(CleaningForm) Error approving form:', error);
    res.status(500).send({ error: 'Failed to approve form', details: error.message });
  }
});

// 11.5: Reject Form
app.post('/api/cleaning-form/reject/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { reason } = req.body;
    const { uid: rejectorId, fullName } = req.user;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'submitted') return res.status(400).send({ error: 'Only submitted forms can be rejected' });

    await ref.update({
      status: 'rejected',
      rejectedBy: rejectorId,
      rejectedByName: fullName,
      rejectedAt: new Date().toISOString(),
      rejectionReason: reason || 'No reason provided',
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'REJECTED',
        performedBy: rejectorId,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Form rejected by ${fullName}. Reason: ${reason || 'Not specified'}`
      })
    });

    // Notify contractor
    notifyContractAdmins(form.contractId, 'Form Rejected', `Form ${form.formId} has been rejected. Reason: ${reason || 'Not specified'}`, 'cleaning_form_rejected', { cleaningFormId: uid, formId: form.formId, reason: reason || '' });

    console.log(`(CleaningForm) Rejected ${form.formId}`);
    res.status(200).send({ message: 'Form rejected successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error rejecting form:', error);
    res.status(500).send({ error: 'Failed to reject form', details: error.message });
  }
});

// 11.6: Submit Score (after approval)
app.post('/api/cleaning-form/score/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const { scoringData, totalScore, maxTotalScore, remarks, grade } = req.body;
    const { uid: scorerId, fullName } = req.user;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'approved') return res.status(400).send({ error: 'Scoring is only allowed for approved forms' });

    const calculatedGrade = grade || (totalScore >= 90 ? 'A' : totalScore >= 80 ? 'B' : totalScore >= 70 ? 'C' : totalScore >= 60 ? 'D' : 'F');

    await ref.update({
      status: 'scored',
      score: totalScore,
      grade: calculatedGrade,
      scoringData: scoringData || { criteria: [], totalScore, maxTotalScore, remarks, grade: calculatedGrade },
      scoredBy: scorerId,
      scoredByName: fullName,
      scoringAt: new Date().toISOString(),
      remarks: remarks || form.remarks,
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'SCORED',
        performedBy: scorerId,
        performedByName: fullName,
        timestamp: new Date().toISOString(),
        details: `Score submitted: ${totalScore}/${maxTotalScore} (Grade: ${calculatedGrade})`
      })
    });

    // Notify: scoring complete, needs contractor acknowledgement
    notifyContractAdmins(form.contractId, 'Form Scored', `Form ${form.formId} scored ${totalScore}/${maxTotalScore} (Grade: ${calculatedGrade}). Please acknowledge.`, 'cleaning_form_scored', { cleaningFormId: uid, formId: form.formId, score: totalScore, grade: calculatedGrade });

    console.log(`(CleaningForm) Scored ${form.formId}: ${totalScore}/${maxTotalScore} (${calculatedGrade})`);
    res.status(200).send({ message: 'Score submitted successfully', grade: calculatedGrade });
  } catch (error) {
    console.error('(CleaningForm) Error scoring form:', error);
    res.status(500).send({ error: 'Failed to submit score', details: error.message });
  }
});

// 11.7: Contractor Acknowledge Score
app.post('/api/cleaning-form/acknowledge/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'scored') return res.status(400).send({ error: 'Only scored forms can be acknowledged' });

    await ref.update({
      status: 'contractorApproved',
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'CONTRACTOR_ACKNOWLEDGED',
        performedBy: req.user.uid,
        performedByName: req.user.fullName,
        timestamp: new Date().toISOString(),
        details: `Contractor acknowledged score: ${form.score}`
      })
    });

    console.log(`(CleaningForm) Contractor acknowledged ${form.formId}`);
    res.status(200).send({ message: 'Score acknowledged successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error acknowledging score:', error);
    res.status(500).send({ error: 'Failed to acknowledge score', details: error.message });
  }
});

// 11.8: Auto-Approve after 30 min timeout (can be called by cron or as fallback)
app.post('/api/cleaning-form/auto-approve/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'scored') return res.status(400).send({ error: 'Only scored forms can be auto-approved' });

    const scoringAt = new Date(form.scoringAt);
    const now = new Date();
    const diffMs = now - scoringAt;
    if (diffMs < 30 * 60 * 1000) {
      return res.status(400).send({ error: `Auto-approval requires 30 minutes after scoring. ${Math.ceil((30 * 60 * 1000 - diffMs) / 60000)} minutes remaining.` });
    }

    await ref.update({
      status: 'autoApproved',
      autoApprovedAt: now.toISOString(),
      updatedAt: now.toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'AUTO_APPROVED',
        performedBy: 'system',
        performedByName: 'System',
        timestamp: now.toISOString(),
        details: 'Auto-approved after 30-minute timeout'
      })
    });

    console.log(`(CleaningForm) Auto-approved ${form.formId}`);
    res.status(200).send({ message: 'Form auto-approved successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error auto-approving:', error);
    res.status(500).send({ error: 'Failed to auto-approve', details: error.message });
  }
});

// 11.9: Lock Form (final state - no further edits)
app.post('/api/cleaning-form/lock/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const ref = db.collection('cleaningForms').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    if (form.status !== 'contractorApproved' && form.status !== 'autoApproved' && form.status !== 'scored') {
      return res.status(400).send({ error: 'Form must be scored and acknowledged before locking' });
    }

    await ref.update({
      status: 'locked',
      lockedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      auditLog: admin.firestore.FieldValue.arrayUnion({
        action: 'LOCKED',
        performedBy: req.user.uid,
        performedByName: req.user.fullName,
        timestamp: new Date().toISOString(),
        details: 'Form locked. Ready for billing.'
      })
    });

    // Notify: form locked, ready for billing
    notifyContractAdmins(form.contractId, 'Form Locked', `Form ${form.formId} has been locked and is ready for billing.`, 'cleaning_form_locked', { cleaningFormId: uid, formId: form.formId, score: form.score, grade: form.grade });

    console.log(`(CleaningForm) Locked ${form.formId}`);
    res.status(200).send({ message: 'Form locked successfully' });
  } catch (error) {
    console.error('(CleaningForm) Error locking form:', error);
    res.status(500).send({ error: 'Failed to lock form', details: error.message });
  }
});

// 11.10: Get Cleaning Form List (with filters)
app.get('/api/cleaning-form/list', verifyToken, async (req, res) => {
  try {
    const { status, formType, contractId, division, depot } = req.query;
    const { role, zone, division: userDiv, entityId, userType } = req.user;
    const userRole = (role || '').toLowerCase();

    let query = db.collection('cleaningForms');
    query = query.orderBy('createdAt', 'desc');

    if (userType === 'contractor') {
      query = query.where('entityId', '==', entityId);
    } else if (userRole.includes('supervisor')) {
      query = query.where('division', '==', userDiv);
    } else if (userRole.includes('admin') || userRole.includes('master')) {
      if (division) query = query.where('division', '==', division);
    }

    if (status) query = query.where('status', '==', status);
    if (formType) query = query.where('formType', '==', formType);
    if (contractId) query = query.where('contractId', '==', contractId);
    if (depot) query = query.where('depot', '==', depot);

    const snapshot = await query.get();
    const forms = [];
    snapshot.forEach(doc => forms.push(doc.data()));

    res.status(200).json({ count: forms.length, forms });
  } catch (error) {
    if (error.code === 'FAILED_PRECONDITION') {
      return res.status(400).json({ error: 'Index required. Check Firebase Console.' });
    }
    console.error('(CleaningForm) Error fetching list:', error);
    res.status(500).send({ error: 'Failed to fetch forms', details: error.message });
  }
});

// 11.11: Get Cleaning Form Details
app.get('/api/cleaning-form/details/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const doc = await db.collection('cleaningForms').doc(uid).get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });
    res.status(200).json({ form: doc.data() });
  } catch (error) {
    console.error('(CleaningForm) Error fetching details:', error);
    res.status(500).send({ error: 'Failed to fetch form details', details: error.message });
  }
});

// 11.12: Get Cleaning Form Dashboard Summary
app.get('/api/cleaning-form/dashboard', verifyToken, async (req, res) => {
  try {
    const { role, division: userDiv, zone, entityId, userType } = req.user;
    const userRole = (role || '').toLowerCase();

    let query = db.collection('cleaningForms');
    let countQuery = db.collection('cleaningForms');

    if (userType === 'contractor') {
      query = query.where('entityId', '==', entityId);
      countQuery = countQuery.where('entityId', '==', entityId);
    } else if (userRole.includes('supervisor')) {
      query = query.where('division', '==', userDiv);
      countQuery = countQuery.where('division', '==', userDiv);
    }

    const snapshot = await query.get();
    const summary = { draftForms: 0, submittedForms: 0, approvedForms: 0, rejectedForms: 0, scoredForms: 0, lockedForms: 0, pendingReview: 0, scoringPending: 0, totalScore: 0, scoredCount: 0, totalManpower: 0, totalMachine: 0 };

    snapshot.forEach(doc => {
      const d = doc.data();
      switch (d.status) {
        case 'draft': summary.draftForms++; break;
        case 'submitted': summary.submittedForms++; summary.pendingReview++; break;
        case 'approved': summary.approvedForms++; summary.scoringPending++; break;
        case 'scored': summary.scoredForms++; summary.totalScore += d.score || 0; summary.scoredCount++; break;
        case 'contractorApproved': summary.scoredForms++; summary.totalScore += d.score || 0; summary.scoredCount++; break;
        case 'autoApproved': summary.scoredForms++; summary.totalScore += d.score || 0; summary.scoredCount++; break;
        case 'locked': summary.lockedForms++; summary.totalScore += d.score || 0; summary.scoredCount++; break;
        case 'rejected': summary.rejectedForms++; break;
      }
      summary.totalManpower += d.manpowerCount || 0;
      summary.totalMachine += d.machineCount || 0;
    });

    const averageScore = summary.scoredCount > 0 ? (summary.totalScore / summary.scoredCount) : 0;

    res.status(200).json({ ...summary, averageScore: Math.round(averageScore * 100) / 100 });
  } catch (error) {
    console.error('(CleaningForm) Error fetching dashboard:', error);
    res.status(500).send({ error: 'Failed to fetch dashboard', details: error.message });
  }
});

// 11.13: Get Cleaning Form Report Data
app.get('/api/cleaning-form/report/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const doc = await db.collection('cleaningForms').doc(uid).get();
    if (!doc.exists) return res.status(404).send({ error: 'Form not found' });

    const form = doc.data();
    res.status(200).json({ report: form });
  } catch (error) {
    console.error('(CleaningForm) Error fetching report:', error);
    res.status(500).send({ error: 'Failed to fetch report', details: error.message });
  }
});

// =======================================================
// == DIVISION MANAGEMENT CRUD
// =======================================================
// POST /api/divisions — Create a new division
app.post('/api/divisions', verifyToken, async (req, res) => {
  try {
    const { name, zone, code } = req.body;
    if (!name || !zone) {
      return res.status(400).json({ error: 'Division name and zone are required.' });
    }
    const ref = db.collection('divisions').doc();
    await ref.set({
      divisionId: ref.id,
      name,
      zone,
      code: code || name.substring(0, 3).toUpperCase(),
      createdBy: req.user.uid,
      createdByName: req.user.fullName,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
    res.status(201).json({ message: 'Division created', divisionId: ref.id });
  } catch (error) {
    res.status(500).json({ error: 'Failed to create division', details: error.message });
  }
});

// GET /api/divisions — List all divisions
app.get('/api/divisions', verifyToken, async (req, res) => {
  try {
    const { zone } = req.query;
    let query = db.collection('divisions');
    if (zone) query = query.where('zone', '==', zone);
    const snapshot = await query.get();
    const divisions = [];
    snapshot.forEach(doc => divisions.push(doc.data()));
    divisions.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    res.status(200).json({ count: divisions.length, divisions });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch divisions', details: error.message });
  }
});

// GET /api/divisions/:id — Get single division
app.get('/api/divisions/:id', verifyToken, async (req, res) => {
  try {
    const doc = await db.collection('divisions').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ error: 'Division not found' });
    res.status(200).json({ division: doc.data() });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch division', details: error.message });
  }
});

// PUT /api/divisions/:id — Update division
app.put('/api/divisions/:id', verifyToken, async (req, res) => {
  try {
    const { name, zone, code } = req.body;
    const ref = db.collection('divisions').doc(req.params.id);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Division not found' });
    const updateData = {};
    if (name !== undefined) updateData.name = name;
    if (zone !== undefined) updateData.zone = zone;
    if (code !== undefined) updateData.code = code;
    updateData.updatedAt = new Date().toISOString();
    await ref.update(updateData);
    res.status(200).json({ message: 'Division updated' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to update division', details: error.message });
  }
});

// DELETE /api/divisions/:id — Delete division
app.delete('/api/divisions/:id', verifyToken, async (req, res) => {
  try {
    const ref = db.collection('divisions').doc(req.params.id);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Division not found' });
    await ref.delete();
    res.status(200).json({ message: 'Division deleted' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to delete division', details: error.message });
  }
});

// =======================================================
// == PROFILE SELF-EDIT
// =======================================================
app.post('/api/user/update-profile', verifyToken, async (req, res) => {
  try {
    const { fullName, designation, mobile } = req.body;
    const { uid } = req.user;
    const ref = db.collection('users').doc(uid);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'User not found' });
    const updateData = {};
    if (fullName !== undefined) updateData.fullName = fullName;
    if (designation !== undefined) updateData.designation = designation;
    if (mobile !== undefined) {
      if (!/^\d{10}$/.test(mobile)) {
        return res.status(400).json({ error: 'Invalid mobile number. Must be 10 digits.' });
      }
      updateData.mobile = mobile;
    }
    if (Object.keys(updateData).length === 0) {
      return res.status(400).json({ error: 'No fields to update.' });
    }
    updateData.updatedAt = new Date().toISOString();
    updateData.updatedBy = uid;
    await ref.update(updateData);
    const updated = await ref.get();
    res.status(200).json({ message: 'Profile updated successfully', user: updated.data() });
  } catch (error) {
    res.status(500).json({ error: 'Failed to update profile', details: error.message });
  }
});

// =======================================================
// == CHANGE PASSWORD (logged-in user)
// =======================================================
app.post('/api/auth/change-password', verifyToken, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;
    if (!currentPassword || !newPassword) {
      return res.status(400).json({ error: 'Current password and new password are required.' });
    }
    if (newPassword.length < 6) {
      return res.status(400).json({ error: 'New password must be at least 6 characters.' });
    }
    const { uid } = req.user;
    const userRef = db.collection('users').doc(uid);
    const userDoc = await userRef.get();
    if (!userDoc.exists) return res.status(404).json({ error: 'User not found' });
    const userData = userDoc.data();
    if (userData.password !== currentPassword) {
      return res.status(403).json({ error: 'Current password is incorrect.' });
    }
    // Update password in Firestore
    await userRef.update({ password: newPassword, updatedAt: new Date().toISOString() });
    // Try to update Firebase Auth password
    try {
      await admin.auth().updateUser(uid, { password: newPassword });
    } catch (authError) {
      // Non-critical — password is stored in Firestore as fallback
    }
    // Audit log
    const auditRef = db.collection('auditLogs').doc();
    await auditRef.set({
      logId: auditRef.id,
      action: 'PASSWORD_CHANGED',
      performedBy: uid,
      performedByName: req.user.fullName,
      targetUser: uid,
      targetUserName: req.user.fullName,
      timestamp: new Date().toISOString(),
      details: 'User changed their password'
    });
    res.status(200).json({ message: 'Password changed successfully.' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to change password', details: error.message });
  }
});

// =======================================================
// == COMPLAINT ESCALATION WORKFLOW
// =======================================================
// PATCH /api/obhs/complaints/assign/:complaintId — Assign complaint
app.patch('/api/obhs/complaints/assign/:complaintId', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.params;
    const { assignedTo, assignedToName, remarks } = req.body;
    const allowedRoles = ['Admin', 'Supervisor', 'Company Master'];
    if (!allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Only Admin, Supervisor, or Company Master can assign complaints.' });
    }
    if (!assignedTo) return res.status(400).json({ error: 'assignedTo (user UID) is required.' });
    const ref = db.collection('obhs_complaints').doc(complaintId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Complaint not found' });
    await ref.update({
      status: 'IN_PROGRESS',
      assignedTo,
      assignedToName: assignedToName || 'Assigned User',
      assignedBy: req.user.uid,
      assignedByName: req.user.fullName,
      assignedAt: new Date().toISOString(),
      assignmentRemarks: remarks || '',
      updatedAt: new Date().toISOString()
    });
    // Audit log
    const auditRef = db.collection('auditLogs').doc();
    await auditRef.set({
      logId: auditRef.id,
      action: 'COMPLAINT_ASSIGNED',
      performedBy: req.user.uid,
      performedByName: req.user.fullName,
      targetUser: assignedTo,
      targetEntity: complaintId,
      targetEntityType: 'obhs_complaints',
      timestamp: new Date().toISOString(),
      details: `Complaint ${complaintId} assigned to ${assignedToName || assignedTo}`
    });
    res.status(200).json({ message: 'Complaint assigned and status set to IN_PROGRESS.' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to assign complaint', details: error.message });
  }
});

// PATCH /api/obhs/complaints/escalate/:complaintId — Escalate complaint
app.patch('/api/obhs/complaints/escalate/:complaintId', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.params;
    const { escalationReason, escalatedTo } = req.body;
    const ref = db.collection('obhs_complaints').doc(complaintId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Complaint not found' });
    const data = doc.data();
    if (data.status === 'CLOSED') {
      return res.status(400).json({ error: 'Cannot escalate a closed complaint.' });
    }
    const escalationLevel = (data.escalationLevel || 0) + 1;
    await ref.update({
      status: 'ESCALATED',
      escalationLevel,
      escalationReason: escalationReason || 'No reason provided',
      escalatedBy: req.user.uid,
      escalatedByName: req.user.fullName,
      escalatedAt: new Date().toISOString(),
      escalatedTo: escalatedTo || null,
      updatedAt: new Date().toISOString()
    });
    const auditRef = db.collection('auditLogs').doc();
    await auditRef.set({
      logId: auditRef.id,
      action: 'COMPLAINT_ESCALATED',
      performedBy: req.user.uid,
      performedByName: req.user.fullName,
      targetEntity: complaintId,
      targetEntityType: 'obhs_complaints',
      timestamp: new Date().toISOString(),
      details: `Complaint escalated to level ${escalationLevel}. Reason: ${escalationReason || 'Not specified'}`
    });
    res.status(200).json({ message: `Complaint escalated to level ${escalationLevel}.` });
  } catch (error) {
    res.status(500).json({ error: 'Failed to escalate complaint', details: error.message });
  }
});

// Update the resolve endpoint to allow resolving from IN_PROGRESS and ESCALATED too
// (The existing PATCH /api/obhs/complaints/resolve/:complaintId already works for OPEN/IN_PROGRESS/ESCALATED)

// =======================================================
// == PDF INVOICE GENERATION
// =======================================================
app.get('/api/billing/invoice-pdf/:uid', verifyToken, async (req, res) => {
  try {
    const { uid } = req.params;
    const doc = await db.collection('billingReports').doc(uid).get();
    if (!doc.exists) return res.status(404).json({ error: 'Billing report not found' });
    const bill = doc.data();
    if (bill.status !== 'APPROVED' && !bill.invoiceNumber) {
      return res.status(400).json({ error: 'Invoice not yet generated. Please generate invoice first.' });
    }
    const invoiceNumber = bill.invoiceNumber || `INV-${uid.substring(0, 8)}`;
    const docPdf = new PDFDocument({ margin: 50 });
    const fileName = `invoice_${invoiceNumber}.pdf`;
    const tmpDir = path.join(process.cwd(), 'tmp');
    if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
    const filePath = path.join(tmpDir, fileName);
    const stream = fs.createWriteStream(filePath);
    docPdf.pipe(stream);
    // Header
    docPdf.fontSize(22).fillColor('#1f4e78').text('SWACHH RAILWAYS', { align: 'center' });
    docPdf.fontSize(10).fillColor('#666').text('Indian Railways – OBHS Billing System', { align: 'center' });
    docPdf.moveDown();
    docPdf.fontSize(16).fillColor('#1f4e78').text('INVOICE', { align: 'center' });
    docPdf.moveDown();
    // Invoice info
    docPdf.fontSize(10).fillColor('#333');
    docPdf.text(`Invoice Number: ${invoiceNumber}`);
    docPdf.text(`Date: ${new Date(bill.invoiceGeneratedAt || new Date()).toLocaleDateString('en-IN')}`);
    docPdf.text(`Period: ${bill.period || 'N/A'}`);
    docPdf.moveDown();
    // Horizontal rule
    docPdf.moveTo(50, docPdf.y).lineTo(545, docPdf.y).strokeColor('#ddd').stroke();
    docPdf.moveDown();
    // Bill details
    docPdf.fontSize(12).fillColor('#1f4e78').text('Bill Details', { underline: true });
    docPdf.moveDown(0.5);
    docPdf.fontSize(10).fillColor('#333');
    const details = [
      ['Contract', bill.contractNumber || 'N/A'],
      ['Entity', bill.entityName || 'N/A'],
      ['Zone', bill.zone || 'N/A'],
      ['Division', bill.division || 'N/A'],
      ['Contract Value', `₹${(bill.contractValue || 0).toLocaleString('en-IN')}`],
      ['Overall Score', `${bill.overallScore || 0}%`],
      ['Grade', bill.grade || 'N/A'],
    ];
    details.forEach(([label, value]) => {
      docPdf.text(`${label}: ${value}`, { continued: false });
    });
    docPdf.moveDown();
    // Deductions table
    if (bill.deductions && bill.deductions.length > 0) {
      docPdf.fontSize(12).fillColor('#1f4e78').text('Deductions', { underline: true });
      docPdf.moveDown(0.5);
      docPdf.fontSize(9).fillColor('#333');
      const tableTop = docPdf.y;
      const col1 = 50, col2 = 200, col3 = 320, col4 = 420, col5 = 500;
      docPdf.font('Helvetica-Bold');
      docPdf.text('Type', col1, tableTop);
      docPdf.text('Description', col2, tableTop);
      docPdf.text('Count', col3, tableTop);
      docPdf.text('Rate', col4, tableTop);
      docPdf.text('Amount', col5, tableTop);
      docPdf.font('Helvetica');
      let y = tableTop + 15;
      bill.deductions.forEach(d => {
        docPdf.text(d.type || 'N/A', col1, y);
        docPdf.text(d.description || 'N/A', col2, y);
        docPdf.text(`${d.count || 0}`, col3, y);
        docPdf.text(`₹${(d.rate || 0).toLocaleString('en-IN')}`, col4, y);
        docPdf.text(`₹${(d.amount || 0).toLocaleString('en-IN')}`, col5, y);
        y += 15;
      });
      docPdf.y = y;
    }
    docPdf.moveDown();
    // Payable
    docPdf.moveTo(50, docPdf.y).lineTo(545, docPdf.y).strokeColor('#ddd').stroke();
    docPdf.moveDown();
    docPdf.fontSize(14).fillColor('#28a745').text(
      `FINAL PAYABLE: ₹${(bill.finalPayable || 0).toLocaleString('en-IN')}`,
      { align: 'center' }
    );
    docPdf.moveDown(2);
    // Footer
    docPdf.fontSize(8).fillColor('#999').text(
      'Swachh Railways – Contractor Employee Portal | Indian Railways',
      { align: 'center' }
    );
    docPdf.text(`Generated on ${new Date().toLocaleString('en-IN')}`, { align: 'center' });
    docPdf.end();
    stream.on('finish', () => {
      res.download(filePath, fileName, (err) => {
        if (err) {
          res.status(500).json({ error: 'Failed to download invoice' });
        }
        // Cleanup after download
        fs.unlink(filePath, () => { });
      });
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to generate invoice PDF', details: error.message });
  }
});

// =======================================================
// == CENTRALIZED AUDIT LOG
// =======================================================
// GET /api/audit/logs — Query audit logs (aggregated from centralized collection + form documents)
app.get('/api/audit/logs', verifyToken, async (req, res) => {
  try {
    const { action, targetEntity, targetUser, limit: limitParam, offset: offsetParam } = req.query;
    const maxLimit = Math.min(parseInt(limitParam) || 50, 200);

    // 1. Fetch from centralized auditLogs collection
    let centralQuery = db.collection('auditLogs').orderBy('timestamp', 'desc');
    if (action) centralQuery = centralQuery.where('action', '==', action);
    if (targetEntity) centralQuery = centralQuery.where('targetEntity', '==', targetEntity);
    if (targetUser) centralQuery = centralQuery.where('targetUser', '==', targetUser);
    const centralSnapshot = await centralQuery.limit(maxLimit).get();
    const logs = [];
    centralSnapshot.forEach(doc => logs.push(doc.data()));

    // 2. Fetch embedded audit logs from form collections for comprehensive view
    const formCollections = ['coachForms', 'premisesForms', 'stationCleaningForms'];
    const formPromises = formCollections.map(collName => {
      let q = db.collection(collName).orderBy('updatedAt', 'desc').limit(50);
      if (targetEntity === 'cleaning_form' || targetEntity === 'CleaningForm') {
        // no additional filter
      }
      return q.get();
    });
    const formSnapshots = await Promise.allSettled(formPromises);
    const formAuditLogs = [];
    for (const result of formSnapshots) {
      if (result.status !== 'fulfilled') continue;
      result.value.forEach(doc => {
        const data = doc.data();
        if (data.auditLog && Array.isArray(data.auditLog)) {
          for (const entry of data.auditLog) {
            if (entry && entry.action) {
              if (action && entry.action !== action) continue;
              if (targetUser && entry.performedBy !== targetUser) continue;
              formAuditLogs.push({
                ...entry,
                _source: collName,
                _formId: data.uid || doc.id,
                _formType: data.formType || collName,
                timestamp: entry.timestamp || data.updatedAt || new Date().toISOString(),
              });
            }
          }
        }
      });
    }

    // 3. Merge and sort by timestamp desc
    logs.push(...formAuditLogs);
    logs.sort((a, b) => {
      const tA = a.timestamp || '';
      const tB = b.timestamp || '';
      return tB.localeCompare(tA);
    });

    // 4. Apply limit
    const limited = logs.slice(0, maxLimit);

    res.status(200).json({ count: limited.length, logs: limited });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch audit logs', details: error.message });
  }
});

// GET /api/audit/logs/stats — Audit log summary stats (aggregated)
app.get('/api/audit/logs/stats', verifyToken, async (req, res) => {
  try {
    const centralSnapshot = await db.collection('auditLogs').get();
    const stats = {};
    centralSnapshot.forEach(doc => {
      const d = doc.data();
      stats[d.action] = (stats[d.action] || 0) + 1;
    });

    // Also aggregate from form collections
    const formCollections = ['coachForms', 'premisesForms', 'stationCleaningForms'];
    const formPromises = formCollections.map(collName =>
      db.collection(collName).get()
    );
    const formSnapshots = await Promise.allSettled(formPromises);
    for (const result of formSnapshots) {
      if (result.status !== 'fulfilled') continue;
      result.value.forEach(doc => {
        const data = doc.data();
        if (data.auditLog && Array.isArray(data.auditLog)) {
          for (const entry of data.auditLog) {
            if (entry && entry.action) {
              stats[entry.action] = (stats[entry.action] || 0) + 1;
            }
          }
        }
      });
    }

    let total = 0;
    for (const key of Object.keys(stats)) {
      total += stats[key];
    }

    res.status(200).json({ stats, total });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch audit stats', details: error.message });
  }
});

// =======================================================
// == AI COMPLAINT ROUTING (Keyword-based auto-categorization)
// =======================================================
app.post('/api/obhs/complaints/auto-route', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.body;
    if (!complaintId) return res.status(400).json({ error: 'complaintId is required.' });
    const ref = db.collection('obhs_complaints').doc(complaintId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Complaint not found' });
    const complaint = doc.data();
    // Keyword-based auto-categorization
    const description = (complaint.description || '').toLowerCase();
    const category = (complaint.category || '').toLowerCase();
    const text = `${category} ${description}`;
    const keywordMap = [
      { keywords: ['toilet', 'bathroom', 'restroom', 'washroom', 'sanitary', 'urinal', 'foul smell', 'stink'], taskType: 'Toilet Cleaning', priority: 2 },
      { keywords: ['clean', 'dirty', 'dust', 'sweep', 'mop', 'garbage', 'waste', 'litter', 'rubbish', 'spill'], taskType: 'Coach Cleaning', priority: 3 },
      { keywords: ['linen', 'bedsheet', 'pillow', 'blanket', 'bed roll', 'curtain'], taskType: 'Linen Distribution', priority: 3 },
      { keywords: ['light', 'fan', 'ac', 'air conditioner', 'cooling', 'electrical', 'plug', 'charging', 'switch'], taskType: 'Electrical Issue', priority: 2 },
      { keywords: ['water', 'drinking', 'supply', 'tap', 'leak', 'pipeline', 'overflow'], taskType: 'Water Issue', priority: 2 },
      { keywords: ['seat', 'berth', 'broken', 'damage', 'repair', 'maintenance', 'crack', 'rust'], taskType: 'Maintenance Issue', priority: 2 },
      { keywords: ['security', 'theft', 'safety', 'suspicious', 'unauthorized'], taskType: 'Security Issue', priority: 1 },
      { keywords: ['staff', 'behaviour', 'rude', 'misbehave', 'attitude', 'service'], taskType: 'Staff Behaviour', priority: 3 },
    ];
    let assignedTaskType = 'Coach Cleaning';
    let assignedPriority = 'medium';
    let bestScore = 0;
    keywordMap.forEach(entry => {
      const score = entry.keywords.reduce((sum, kw) => sum + (text.includes(kw) ? 1 : 0), 0);
      if (score > bestScore) {
        bestScore = score;
        assignedTaskType = entry.taskType;
        assignedPriority = entry.priority;
      }
    });
    // Update complaint with routing suggestion
    await ref.update({
      suggestedTaskType: assignedTaskType,
      suggestedPriority: assignedPriority,
      autoRoutedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
    // Audit log
    const auditRef = db.collection('auditLogs').doc();
    await auditRef.set({
      logId: auditRef.id,
      action: 'COMPLAINT_AUTO_ROUTED',
      performedBy: 'system',
      performedByName: 'System',
      targetEntity: complaintId,
      targetEntityType: 'obhs_complaints',
      timestamp: new Date().toISOString(),
      details: `Auto-routed complaint ${complaintId} → Task: ${assignedTaskType}, Priority: ${assignedPriority}`
    });
    res.status(200).json({
      message: 'Complaint auto-routed',
      suggestedTaskType: assignedTaskType,
      suggestedPriority: assignedPriority
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to auto-route complaint', details: error.message });
  }
});

// =======================================================
// == EVIDENCE MANAGEMENT SYSTEM
// =======================================================

// GET /api/evidence/types — List supported evidence types (no auth required for this metadata)
app.get('/api/evidence/types', (req, res) => {
  res.json({
    success: true,
    types: evidence.evidenceUtils.EVIDENCE_TYPES
  });
});

// POST /api/evidence/upload — Upload evidence with WEBP conversion, thumbnail, dedup
app.post('/api/evidence/upload', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const {
      trainNumber, coach, taskId, taskType, evidenceType,
      contractor, runInstanceId, complaintId, attendanceId,
      gpsLat, gpsLng, remarks
    } = req.body;

    const user = req.user;
    const division = user.division || 'Unknown';
    const date = new Date().toISOString().split('T')[0];

    // Duplicate check
    const fileHash = evidence.computeFileHash(req.file.buffer);
    const existing = await evidence.findDuplicateByHash(db, fileHash);
    if (existing) {
      return res.status(200).json({
        success: true,
        duplicate: true,
        evidenceId: existing.id,
        message: 'Duplicate image detected. Returning existing record.',
        ...existing
      });
    }

    // Process image (WEBP + thumbnail + compression)
    const processed = await evidence.processImage(req.file.buffer);

    // Build storage path: Division/Train/Date/Coach/Task/EvidenceType/
    const filename = `${uuidv4()}_${Date.now()}`;
    const storageRef = evidence.buildStoragePath(
      division, trainNumber || 'UNKNOWN', date,
      coach || 'GENERAL', taskId || 'no-task', evidenceType || 'Before', filename
    );

    // Upload to Firebase Storage
    const storageResult = await evidence.uploadToStorage(bucket, storageRef, processed);

    // Create metadata document
    const meta = evidence.buildEvidenceMetadata({
      storageRef,
      trainNumber,
      coach,
      taskId,
      taskType,
      evidenceType: evidenceType || 'Before',
      uploadedBy: user.uid,
      uploadedByName: user.fullName || user.name,
      uploadedByRole: user.role,
      division,
      contractor,
      runInstanceId,
      fileHash,
      complaintId,
      attendanceId,
      gpsLat,
      gpsLng,
      remarks,
      originalSize: processed.originalSize,
      compressedSize: processed.compressedSize,
      thumbnailSize: processed.thumbnailSize,
      compressionRatio: processed.compressionRatio,
      width: processed.width,
      height: processed.height,
      originalFormat: processed.originalFormat,
      webpUrl: storageResult.webpUrl,
      thumbUrl: storageResult.thumbUrl,
      webpPath: storageResult.webpPath,
      thumbPath: storageResult.thumbPath
    });

    const docRef = await db.collection('evidence_metadata').add(meta);
    await evidence.logAudit(db, 'EVIDENCE_UPLOADED', {
      evidenceId: docRef.id, userId: user.uid, userName: user.fullName,
      userRole: user.role, trainNumber, coach, taskId, complaintId,
      details: `Uploaded ${evidenceType || 'Before'} evidence (${(processed.compressionRatio)}% compression)`
    });

    res.status(201).json({
      success: true,
      evidenceId: docRef.id,
      webpUrl: storageResult.webpUrl,
      thumbUrl: storageResult.thumbUrl,
      originalSize: processed.originalSize,
      compressedSize: processed.compressedSize,
      compressionRatio: processed.compressionRatio,
      width: processed.width,
      height: processed.height,
      message: 'Evidence uploaded and compressed successfully'
    });

  } catch (error) {
    console.error('Evidence upload error:', error);
    res.status(500).json({ error: 'Failed to upload evidence', details: error.message });
  }
});

// POST /api/evidence/upload/base64 — Upload base64-encoded evidence
app.post('/api/evidence/upload/base64', verifyToken, async (req, res) => {
  try {
    const { image, ...metadata } = req.body;
    if (!image) return res.status(400).json({ error: 'No image data provided' });

    const base64Data = image.replace(/^data:image\/\w+;base64,/, '');
    const buffer = Buffer.from(base64Data, 'base64');

    // Create a fake file object
    req.file = { buffer, mimetype: 'image/png', originalname: `evidence_${Date.now()}.png` };
    req.body = metadata;

    // Forward to the main upload handler logic
    const handler = app._router.stack.find(l => l.route?.path === '/api/evidence/upload' && l.route.methods?.post);
    if (handler) {
      return handler.route.stack[0].handle(req, res);
    }
    throw new Error('Upload handler not found');
  } catch (error) {
    console.error('Base64 evidence upload error:', error);
    res.status(500).json({ error: 'Failed to process base64 evidence', details: error.message });
  }
});

// GET /api/evidence/:id — Get single evidence record (active/archive/long-term)
app.get('/api/evidence/:id', verifyToken, async (req, res) => {
  try {
    const record = await evidence.getEvidenceById(db, req.params.id);
    if (!record) return res.status(404).json({ error: 'Evidence not found' });

    await evidence.logAudit(db, 'EVIDENCE_VIEWED', {
      evidenceId: req.params.id, userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, trainNumber: record.trainNumber,
      details: 'Evidence record viewed'
    });

    res.json({ success: true, evidence: record });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch evidence', details: error.message });
  }
});

// PUT /api/evidence/:id — Update evidence metadata (edit remarks, etc.)
app.put('/api/evidence/:id', verifyToken, async (req, res) => {
  try {
    const { id } = req.params;
    const allowed = ['remarks', 'gpsLat', 'gpsLng', 'evidenceType'];
    const updates = {};
    for (const key of allowed) {
      if (req.body[key] !== undefined) updates[key] = req.body[key];
    }
    updates.updatedAt = admin.firestore.FieldValue.serverTimestamp();

    const doc = await db.collection('evidence_metadata').doc(id).get();
    if (!doc.exists) return res.status(404).json({ error: 'Evidence not found' });

    await doc.ref.update(updates);
    await evidence.logAudit(db, 'EVIDENCE_UPDATED', {
      evidenceId: id, userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, details: `Updated fields: ${Object.keys(updates).join(', ')}`
    });

    res.json({ success: true, message: 'Evidence updated' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to update evidence', details: error.message });
  }
});

// DELETE /api/evidence/:id — Soft delete evidence
app.delete('/api/evidence/:id', verifyToken, async (req, res) => {
  try {
    const { id } = req.params;
    const doc = await db.collection('evidence_metadata').doc(id).get();
    if (!doc.exists) return res.status(404).json({ error: 'Evidence not found' });

    await doc.ref.update({
      deleted: true,
      deletedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await evidence.logAudit(db, 'EVIDENCE_DELETED', {
      evidenceId: id, userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, details: 'Evidence soft-deleted'
    });

    res.json({ success: true, message: 'Evidence deleted' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to delete evidence', details: error.message });
  }
});

// GET /api/evidence/search — Search evidence across active/archive
app.get('/api/evidence/search', verifyToken, async (req, res) => {
  try {
    const params = {
      trainNumber: req.query.trainNumber,
      coach: req.query.coach,
      dateFrom: req.query.dateFrom,
      dateTo: req.query.dateTo,
      contractor: req.query.contractor,
      worker: req.query.worker,
      supervisor: req.query.supervisor,
      taskType: req.query.taskType,
      evidenceType: req.query.evidenceType,
      complaintId: req.query.complaintId,
      runInstanceId: req.query.runInstanceId,
      storageTier: req.query.storageTier,
      uploadedBy: req.query.uploadedBy,
      limit: parseInt(req.query.limit) || 50,
      offset: parseInt(req.query.offset) || 0
    };

    const results = await evidence.searchEvidence(db, params);

    await evidence.logAudit(db, 'EVIDENCE_SEARCHED', {
      userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role,
      details: `Search: train=${params.trainNumber || '*'}, coach=${params.coach || '*'}, type=${params.evidenceType || '*'}`
    });

    res.json({ success: true, ...results });
  } catch (error) {
    res.status(500).json({ error: 'Search failed', details: error.message });
  }
});

// POST /api/evidence/:id/archive — Manually archive evidence
app.post('/api/evidence/:id/archive', verifyToken, async (req, res) => {
  try {
    const { id } = req.params;
    const doc = await db.collection('evidence_metadata').doc(id).get();
    if (!doc.exists) return res.status(404).json({ error: 'Evidence not found in active storage' });

    const data = doc.data();
    await db.collection('archive_evidence').doc(id).set({
      ...data, storageTier: evidence.STORAGE_TIER.ARCHIVE,
      archivedAt: admin.firestore.FieldValue.serverTimestamp()
    });
    await doc.ref.update({
      storageTier: evidence.STORAGE_TIER.ARCHIVE,
      archivedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await evidence.logAudit(db, 'EVIDENCE_ARCHIVED_MANUAL', {
      evidenceId: id, userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, details: 'Manually archived'
    });

    res.json({ success: true, message: 'Evidence archived' });
  } catch (error) {
    res.status(500).json({ error: 'Archive failed', details: error.message });
  }
});

// POST /api/evidence/:id/restore — Restore from archive/long-term
app.post('/api/evidence/:id/restore', verifyToken, async (req, res) => {
  try {
    const result = await evidence.restoreFromArchive(db, req.params.id);
    await evidence.logAudit(db, 'EVIDENCE_RESTORED_MANUAL', {
      evidenceId: req.params.id, userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, details: `Restored from ${result.tier}`
    });
    res.json({ success: true, ...result });
  } catch (error) {
    res.status(500).json({ error: 'Restore failed', details: error.message });
  }
});

// =======================================================
// == STORAGE ANALYTICS
// =======================================================

// GET /api/storage/analytics — Storage usage, trends, forecast
app.get('/api/storage/analytics', verifyToken, async (req, res) => {
  try {
    const analytics = await evidence.getStorageAnalytics(db);
    res.json({ success: true, ...analytics });
  } catch (error) {
    res.status(500).json({ error: 'Analytics failed', details: error.message });
  }
});

// GET /api/storage/per-train — Image count per train
app.get('/api/storage/per-train', verifyToken, async (req, res) => {
  try {
    const snap = await db.collection('evidence_metadata')
      .where('deleted', '==', false).get();
    const trainCounts = {};
    snap.docs.forEach(doc => {
      const t = doc.data().trainNumber || 'UNKNOWN';
      trainCounts[t] = (trainCounts[t] || 0) + 1;
    });
    const sorted = Object.entries(trainCounts)
      .map(([train, count]) => ({ train, count }))
      .sort((a, b) => b.count - a.count);
    res.json({ success: true, trains: sorted });
  } catch (error) {
    res.status(500).json({ error: 'Failed', details: error.message });
  }
});

// GET /api/storage/per-contractor — Storage per contractor
app.get('/api/storage/per-contractor', verifyToken, async (req, res) => {
  try {
    const snap = await db.collection('evidence_metadata')
      .where('deleted', '==', false).get();
    const contractorStats = {};
    snap.docs.forEach(doc => {
      const d = doc.data();
      const c = d.contractor || 'Unknown';
      if (!contractorStats[c]) contractorStats[c] = { count: 0, totalBytes: 0 };
      contractorStats[c].count++;
      contractorStats[c].totalBytes += d.compressedSize || d.originalSize || 0;
    });
    res.json({ success: true, contractors: contractorStats });
  } catch (error) {
    res.status(500).json({ error: 'Failed', details: error.message });
  }
});

// GET /api/storage/daily-upload-count — Daily upload counts
app.get('/api/storage/daily-upload-count', verifyToken, async (req, res) => {
  try {
    const from = req.query.from || new Date(Date.now() - 30 * 86400000).toISOString().split('T')[0];
    const snap = await db.collection('evidence_metadata')
      .where('deleted', '==', false).get();
    const dailyCounts = {};
    snap.docs.forEach(doc => {
      const d = doc.data();
      if (d.uploadedAt?.toDate) {
        const day = d.uploadedAt.toDate().toISOString().split('T')[0];
        if (day >= from) dailyCounts[day] = (dailyCounts[day] || 0) + 1;
      }
    });
    res.json({ success: true, from, daily: dailyCounts });
  } catch (error) {
    res.status(500).json({ error: 'Failed', details: error.message });
  }
});

// =======================================================
// == AUDIT TRAIL
// =======================================================

// GET /api/audit/evidence — Evidence audit log
app.get('/api/audit/evidence', verifyToken, async (req, res) => {
  try {
    const { evidenceId, action, userId, limit = 50 } = req.query;
    let query = db.collection('audit_evidence').orderBy('timestamp', 'desc').limit(parseInt(limit));

    if (evidenceId) query = query.where('evidenceId', '==', evidenceId);
    if (action) query = query.where('action', '==', action);
    if (userId) query = query.where('userId', '==', userId);

    const snap = await query.get();
    const logs = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    res.json({ success: true, count: logs.length, logs });
  } catch (error) {
    res.status(500).json({ error: 'Audit fetch failed', details: error.message });
  }
});

// =======================================================
// == REPORTING ENGINE
// =======================================================

// POST /api/reports/generate — Generate PDF report
app.post('/api/reports/generate', verifyToken, async (req, res) => {
  try {
    const reportData = {
      title: req.body.title || 'EVIDENCE REPORT',
      generatedBy: req.user.fullName || req.user.name,
      trainNumber: req.body.trainNumber,
      coach: req.body.coach,
      contractor: req.body.contractor,
      periodStart: req.body.periodStart,
      periodEnd: req.body.periodEnd,
      sections: req.body.sections || [],
      images: req.body.images || []
    };

    const pdfBuffer = await evidence.generatePDFReport(reportData);

    // Store report record
    const reportRef = db.collection('report_history').doc();
    await reportRef.set({
      reportId: reportRef.id,
      title: reportData.title,
      type: req.body.reportType || 'custom',
      generatedBy: req.user.uid,
      generatedByName: req.user.fullName,
      generatedByRole: req.user.role,
      trainNumber: reportData.trainNumber || null,
      coach: reportData.coach || null,
      contractor: reportData.contractor || null,
      periodStart: reportData.periodStart || null,
      periodEnd: reportData.periodEnd || null,
      fileSize: pdfBuffer.length,
      generatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await evidence.logAudit(db, 'REPORT_GENERATED', {
      userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, trainNumber: reportData.trainNumber,
      details: `Generated report: ${reportData.title}`
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${reportData.title.replace(/[^a-zA-Z0-9]/g, '_')}.pdf"`);
    res.send(pdfBuffer);

  } catch (error) {
    res.status(500).json({ error: 'Report generation failed', details: error.message });
  }
});

// POST /api/reports/email — Send email report with role-based routing
app.post('/api/reports/email', verifyToken, async (req, res) => {
  try {
    const {
      reportType, subject, bodyHtml, sections, images,
      periodStart, periodEnd, trainNumber, coach, contractor,
      to, cc, bcc
    } = req.body;

    const sgMail = (await import('@sendgrid/mail')).default;
    if (process.env.SENDGRID_API_KEY) {
      sgMail.setApiKey(process.env.SENDGRID_API_KEY);
    }

    // Determine recipients based on role
    const recipients = to && to.length > 0
      ? to
      : evidence.getRecipientsForRole(req.user.role, { email: req.user.email });

    // Generate PDF and attach
    const pdfBuffer = await evidence.generatePDFReport({
      title: subject || `${reportType} Report`,
      generatedBy: req.user.fullName,
      trainNumber, coach, contractor,
      periodStart, periodEnd,
      sections: sections || [],
      images: images || []
    });

    const msg = {
      to: recipients,
      cc: cc || [],
      bcc: bcc || [],
      from: process.env.EMAIL_FROM || 'noreply@swachhrailways.com',
      subject: subject || `${reportType.toUpperCase()} Report - ${new Date().toISOString().split('T')[0]}`,
      html: bodyHtml || `<p>Please find attached the ${reportType} report.</p>`,
      attachments: [{
        content: pdfBuffer.toString('base64'),
        filename: `${reportType}_${new Date().toISOString().split('T')[0]}.pdf`,
        type: 'application/pdf',
        disposition: 'attachment'
      }]
    };

    let sendResult = { messageId: 'simulated' };
    if (process.env.SENDGRID_API_KEY) {
      sendResult = await sgMail.send(msg);
    }

    // Store email history
    const emailRef = db.collection('email_history').doc();
    await emailRef.set({
      emailId: emailRef.id,
      reportType,
      subject: msg.subject,
      to: msg.to,
      cc: msg.cc,
      sentBy: req.user.uid,
      sentByName: req.user.fullName,
      sentByRole: req.user.role,
      pdfSize: pdfBuffer.length,
      sendgridMessageId: sendResult.messageId || sendResult[0]?.statusCode || 'simulated',
      sentAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await evidence.logAudit(db, 'REPORT_EMAILED', {
      userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role, trainNumber,
      details: `Emailed ${reportType} report to ${recipients.join(', ')}`
    });

    res.json({
      success: true,
      message: 'Report emailed successfully',
      emailId: emailRef.id,
      recipients,
      pdfSize: pdfBuffer.length
    });

  } catch (error) {
    console.error('Email report error:', error);
    res.status(500).json({ error: 'Failed to send email report', details: error.message });
  }
});

// GET /api/reports/history — Report generation history
app.get('/api/reports/history', verifyToken, async (req, res) => {
  try {
    const { limit = 50, type, trainNumber } = req.query;
    let query = db.collection('report_history').orderBy('generatedAt', 'desc').limit(parseInt(limit));

    if (type) query = query.where('type', '==', type);
    if (trainNumber) query = query.where('trainNumber', '==', trainNumber);

    const snap = await query.get();
    const reports = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    res.json({ success: true, count: reports.length, reports });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch report history', details: error.message });
  }
});

// GET /api/reports/email-history — Email sending history
app.get('/api/reports/email-history', verifyToken, async (req, res) => {
  try {
    const { limit = 50 } = req.query;
    const snap = await db.collection('email_history')
      .orderBy('sentAt', 'desc')
      .limit(parseInt(limit))
      .get();
    const emails = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    res.json({ success: true, count: emails.length, emails });
  } catch (error) {
    res.status(500).json({ error: 'Failed', details: error.message });
  }
});

// =======================================================
// == OFFICIAL AUDIT REPORTS (PDF & Excel)
// =======================================================

app.post('/api/reports/generate/audit', verifyToken, async (req, res) => {
  try {
    const { reportType, runInstanceId, workerId, format = 'pdf' } = req.body;
    if (!reportType || !runInstanceId) {
      return res.status(400).json({ error: 'reportType and runInstanceId are required' });
    }

    const reportService = new ReportService(db);
    let data;
    let pdfBuffer;
    const pdfRenderer = new PDFRenderer();

    if (reportType === 'ATTENDANCE_AUDIT') {
      data = await reportService.getAttendanceAuditData(runInstanceId, workerId);
      pdfBuffer = await pdfRenderer.generateAttendanceAudit(data);
    } else if (reportType === 'WORKER_ACTIVITY') {
      // For now, mapping to Attendance data to match the pattern (Needs separate activity logic)
      data = await reportService.getAttendanceAuditData(runInstanceId, workerId); 
      pdfBuffer = await pdfRenderer.generateWorkerActivityAudit(data);
    } else if (reportType === 'COMPLAINT_AUDIT') {
      // For now, mapping to Attendance data to match the pattern (Needs separate complaint logic)
      data = await reportService.getAttendanceAuditData(runInstanceId, workerId); 
      pdfBuffer = await pdfRenderer.generateComplaintAudit(data);
    } else if (reportType === 'OPERATIONAL_AUDIT') {
      data = await reportService.getOperationalAuditData(runInstanceId);
      pdfBuffer = await pdfRenderer.generateEnterpriseOperationalAudit(data);
    } else {
      return res.status(400).json({ error: 'Invalid reportType' });
    }

    if (format === 'excel') {
      const excelGen = new ExcelGenerator();
      excelGen.createSummarySheet(data, reportType.replace('_', ' '));
      const excelBuffer = await excelGen.getBuffer();
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${data.meta.reportId}.xlsx"`);
      return res.send(excelBuffer);
    }

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${data.meta.reportId}.pdf"`);
    res.send(pdfBuffer);

  } catch (error) {
    console.error('(AuditReport) Error:', error);
    res.status(500).json({ error: 'Failed to generate audit report', details: error.message });
  }
});

app.post('/api/reports/send-email', verifyToken, async (req, res) => {
  try {
    const { reportType, runInstanceId, workerId, emailTo = 'hirenkodwani@gmail.com' } = req.body;
    
    if (!reportType || !runInstanceId) {
      return res.status(400).json({ error: 'reportType and runInstanceId are required' });
    }

    const reportService = new ReportService(db);
    const pdfRenderer = new PDFRenderer();
    const excelGen = new ExcelGenerator();
    let data;
    let pdfBuffer;

    if (reportType === 'ATTENDANCE_AUDIT') {
      data = await reportService.getAttendanceAuditData(runInstanceId, workerId);
      pdfBuffer = await pdfRenderer.generateAttendanceAudit(data);
    } else if (reportType === 'OPERATIONAL_AUDIT') {
      data = await reportService.getOperationalAuditData(runInstanceId);
      pdfBuffer = await pdfRenderer.generateEnterpriseOperationalAudit(data);
    } else if (reportType === 'WORKER_ACTIVITY_AUDIT') {
      data = await reportService.getWorkerActivityAuditData(runInstanceId, workerId);
      pdfBuffer = await pdfRenderer.generateWorkerActivityAudit(data);
    } else if (reportType === 'COMPLAINT_AUDIT') {
      data = await reportService.getComplaintAuditData(runInstanceId);
      pdfBuffer = await pdfRenderer.generateComplaintAudit(data);
    } else {
      // Fallback
      data = await reportService.getOperationalAuditData(runInstanceId);
      pdfBuffer = await pdfRenderer.generateEnterpriseOperationalAudit(data);
    }

    excelGen.createSummarySheet(data, reportType.replace('_', ' '));
    const excelBuffer = await excelGen.getBuffer();

    const emailHtml = `
      <h2>OBHS Official Audit Report</h2>
      <p>Please find attached the latest <b>${reportType}</b> for Run Instance ID: ${runInstanceId}.</p>
      <p><b>Status:</b> ${data.kpi.overallStatus}</p>
      <br>
      <p>Generated by: ${req.user.fullName || 'OBHS System'}</p>
    `;

    const emailResponse = await resend.emails.send({
      from: 'Swachh Railways <reports@swachhrailways.com>',
      to: [emailTo],
      subject: `OBHS Compliance Report | ${data.trainInfo['Train Name']} | ${data.meta.auditStatus}`,
      html: emailHtml,
      attachments: [
        {
          filename: `${data.meta.reportId}.pdf`,
          content: pdfBuffer
        },
        {
          filename: `${data.meta.reportId}.xlsx`,
          content: excelBuffer
        }
      ]
    });

    res.json({ success: true, message: 'Email sent successfully to Higher Authority', emailResponse });
  } catch (error) {
    console.error('(SendEmail) Error:', error);
    res.status(500).json({ error: 'Failed to send email', details: error.message });
  }
});

// =======================================================
// == SCHEDULED REPORTS (Daily / Weekly / Monthly)
// =======================================================

// POST /api/reports/generate/daily — Generate daily train report
app.post('/api/reports/generate/daily', verifyToken, async (req, res) => {
  try {
    const { trainNumber, runInstanceId } = req.body;
    if (!trainNumber) return res.status(400).json({ error: 'trainNumber required' });

    const today = new Date().toISOString().split('T')[0];

    // Gather task data
    const tasksSnap = await db.collection('obhs_task_instances')
      .where('trainNumber', '==', trainNumber)
      .where('date', '==', today)
      .get();

    const tasks = tasksSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    const completed = tasks.filter(t => t.status === 'completed');
    const pending = tasks.filter(t => t.status === 'pending' || t.status === 'assigned');
    const escalated = tasks.filter(t => t.status === 'escalated');

    // Gather complaint data
    const complaintsSnap = await db.collection('obhs_complaints')
      .where('trainNumber', '==', trainNumber)
      .where('date', '==', today)
      .get();
    const complaints = complaintsSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    // Gather evidence data
    const evidenceSnap = await db.collection('evidence_metadata')
      .where('trainNumber', '==', trainNumber)
      .where('deleted', '==', false)
      .get();
    const evidenceDocs = evidenceSnap.docs.map(d => d.data());

    // Gather supervisor approvals
    const approvalsSnap = await db.collection('task_approvals')
      .where('trainNumber', '==', trainNumber)
      .where('date', '==', today)
      .get();

    const sections = [
      {
        heading: 'Task Summary',
        colWidths: [150, 80, 80, 80, 80],
        headers: ['Metric', 'Count'],
        fields: ['label', 'value'],
        items: [
          { label: 'Tasks Generated', value: tasks.length },
          { label: 'Tasks Completed', value: completed.length },
          { label: 'Pending Tasks', value: pending.length },
          { label: 'Escalations', value: escalated.length }
        ]
      },
      {
        heading: 'Complaint Summary',
        colWidths: [150, 80, 80, 80, 80],
        headers: ['Metric', 'Count'],
        fields: ['label', 'value'],
        items: [
          { label: 'Total Complaints', value: complaints.length },
          { label: 'Images Uploaded', value: evidenceDocs.length }
        ]
      },
      {
        heading: 'Evidence Summary',
        colWidths: [150, 80, 80, 80, 80],
        headers: ['Evidence Type', 'Count'],
        fields: ['type', 'count'],
        items: Object.entries(
          evidenceDocs.reduce((acc, e) => {
            acc[e.evidenceType] = (acc[e.evidenceType] || 0) + 1;
            return acc;
          }, {})
        ).map(([type, count]) => ({ type, count }))
      }
    ];

    const pdfBuffer = await evidence.generatePDFReport({
      title: `DAILY TRAIN REPORT - ${trainNumber}`,
      generatedBy: req.user.fullName,
      trainNumber,
      periodStart: today,
      periodEnd: today,
      sections
    });

    // Store report
    const reportRef = db.collection('report_history').doc();
    await reportRef.set({
      reportId: reportRef.id,
      title: `Daily Report - ${trainNumber}`,
      type: 'daily',
      generatedBy: req.user.uid,
      generatedByName: req.user.fullName,
      trainNumber,
      periodStart: today,
      periodEnd: today,
      summary: {
        totalTasks: tasks.length,
        completedTasks: completed.length,
        pendingTasks: pending.length,
        escalations: escalated.length,
        complaints: complaints.length,
        images: evidenceDocs.length
      },
      generatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="DAILY_${trainNumber}_${today}.pdf"`);
    res.send(pdfBuffer);

  } catch (error) {
    res.status(500).json({ error: 'Daily report failed', details: error.message });
  }
});

// POST /api/reports/generate/weekly — Generate weekly contractor report
app.post('/api/reports/generate/weekly', verifyToken, async (req, res) => {
  try {
    const { contractor } = req.body;
    if (!contractor) return res.status(400).json({ error: 'contractor required' });

    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - 7 * 86400000);
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];

    // Gather evidence
    const evidenceSnap = await db.collection('evidence_metadata')
      .where('contractor', '==', contractor)
      .where('deleted', '==', false)
      .get();
    const evidenceDocs = evidenceSnap.docs.map(d => d.data());

    // Gather complaints
    const complaintsSnap = await db.collection('obhs_complaints')
      .where('contractorId', '==', contractor)
      .get();
    const complaints = complaintsSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    // Gather attendance
    const attendanceSnap = await db.collection('attendance_records')
      .where('contractorId', '==', contractor)
      .where('date', '>=', startStr)
      .where('date', '<=', endStr)
      .get();
    const attendance = attendanceSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    const presentCount = attendance.filter(a => a.status === 'present').length;
    const attendanceCompliance = attendance.length > 0
      ? ((presentCount / attendance.length) * 100).toFixed(1)
      : 'N/A';

    // Evidence compliance
    const tasksSnap = await db.collection('obhs_task_instances')
      .where('contractorId', '==', contractor)
      .get();
    const tasks = tasksSnap.docs.map(d => ({ id: d.id, ...d.data() }));
    const completedTasks = tasks.filter(t => t.status === 'completed');
    const completionPct = tasks.length > 0
      ? ((completedTasks.length / tasks.length) * 100).toFixed(1)
      : 'N/A';

    const sections = [
      {
        heading: 'Weekly Performance Summary',
        items: [
          { label: 'Contractor', value: contractor },
          { label: 'Period', value: `${startStr} to ${endStr}` },
          { label: 'Completion %', value: `${completionPct}%` },
          { label: 'Attendance Compliance', value: `${attendanceCompliance}%` },
          { label: 'Evidence Uploaded', value: evidenceDocs.length },
          { label: 'Complaints', value: complaints.length }
        ].map(i => ({ label: i.label, value: String(i.value) }))
      }
    ];

    const pdfBuffer = await evidence.generatePDFReport({
      title: `WEEKLY CONTRACTOR REPORT`,
      generatedBy: req.user.fullName,
      contractor,
      periodStart: startStr,
      periodEnd: endStr,
      sections
    });

    const reportRef = db.collection('report_history').doc();
    await reportRef.set({
      reportId: reportRef.id,
      title: `Weekly Report - ${contractor}`,
      type: 'weekly',
      generatedBy: req.user.uid,
      generatedByName: req.user.fullName,
      contractor,
      periodStart: startStr,
      periodEnd: endStr,
      summary: {
        completionPct,
        attendanceCompliance,
        evidenceCount: evidenceDocs.length,
        complaints: complaints.length
      },
      generatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="WEEKLY_${contractor}_${endStr}.pdf"`);
    res.send(pdfBuffer);

  } catch (error) {
    res.status(500).json({ error: 'Weekly report failed', details: error.message });
  }
});

// POST /api/reports/generate/monthly — Generate monthly division report
app.post('/api/reports/generate/monthly', verifyToken, async (req, res) => {
  try {
    const { division } = req.body;
    if (!division) return res.status(400).json({ error: 'division required' });

    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - 30 * 86400000);
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];

    // Gather evidence
    const evidenceSnap = await db.collection('evidence_metadata')
      .where('division', '==', division)
      .where('deleted', '==', false)
      .get();
    const evidenceDocs = evidenceSnap.docs.map(d => d.data());

    // Contractor breakdown
    const contractorBreakdown = {};
    evidenceDocs.forEach(e => {
      const c = e.contractor || 'Unknown';
      if (!contractorBreakdown[c]) contractorBreakdown[c] = { evidenceCount: 0, totalBytes: 0 };
      contractorBreakdown[c].evidenceCount++;
      contractorBreakdown[c].totalBytes += e.compressedSize || e.originalSize || 0;
    });

    // Train breakdown
    const trainBreakdown = {};
    evidenceDocs.forEach(e => {
      const t = e.trainNumber || 'Unknown';
      trainBreakdown[t] = (trainBreakdown[t] || 0) + 1;
    });

    const sections = [
      {
        heading: 'Evidence Summary',
        items: [
          { label: 'Total Evidence', value: String(evidenceDocs.length) },
          { label: 'Unique Trains', value: String(Object.keys(trainBreakdown).length) },
          { label: 'Unique Contractors', value: String(Object.keys(contractorBreakdown).length) }
        ]
      },
      {
        heading: 'Contractor Comparison',
        colWidths: [120, 80, 100],
        headers: ['Contractor', 'Evidence', 'Storage (MB)'],
        fields: ['name', 'evidence', 'storage'],
        items: Object.entries(contractorBreakdown).map(([name, data]) => ({
          name,
          evidence: String(data.evidenceCount),
          storage: (data.totalBytes / (1024 * 1024)).toFixed(2)
        }))
      },
      {
        heading: 'Train Evidence Count',
        colWidths: [120, 80],
        headers: ['Train', 'Images'],
        fields: ['train', 'count'],
        items: Object.entries(trainBreakdown)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 20)
          .map(([train, count]) => ({ train, count: String(count) }))
      }
    ];

    const pdfBuffer = await evidence.generatePDFReport({
      title: `MONTHLY DIVISION REPORT`,
      generatedBy: req.user.fullName,
      periodStart: startStr,
      periodEnd: endStr,
      sections
    });

    const reportRef = db.collection('report_history').doc();
    await reportRef.set({
      reportId: reportRef.id,
      title: `Monthly Report - ${division}`,
      type: 'monthly',
      generatedBy: req.user.uid,
      generatedByName: req.user.fullName,
      division,
      periodStart: startStr,
      periodEnd: endStr,
      summary: {
        totalEvidence: evidenceDocs.length,
        contractors: Object.keys(contractorBreakdown).length,
        trains: Object.keys(trainBreakdown).length
      },
      generatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="MONTHLY_${division}_${endStr}.pdf"`);
    res.send(pdfBuffer);

  } catch (error) {
    res.status(500).json({ error: 'Monthly report failed', details: error.message });
  }
});

// =======================================================
// == BACKUP OPERATIONS
// =======================================================

// POST /api/backup/evidence — Trigger manual backup
app.post('/api/backup/evidence', verifyToken, async (req, res) => {
  try {
    const type = req.body.type || 'daily';
    const result = await evidence.performBackup(db, type);

    await evidence.logAudit(db, 'BACKUP_PERFORMED', {
      userId: req.user.uid, userName: req.user.fullName,
      userRole: req.user.role,
      details: `Backup type: ${type}, records: ${JSON.stringify(result.recordCounts)}`
    });

    res.json({ success: true, ...result });
  } catch (error) {
    res.status(500).json({ error: 'Backup failed', details: error.message });
  }
});

// GET /api/backup/logs — Backup history
app.get('/api/backup/logs', verifyToken, async (req, res) => {
  try {
    const { limit = 20 } = req.query;
    const snap = await db.collection('backup_logs')
      .orderBy('timestamp', 'desc')
      .limit(parseInt(limit))
      .get();
    const logs = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    res.json({ success: true, count: logs.length, logs });
  } catch (error) {
    res.status(500).json({ error: 'Failed', details: error.message });
  }
});

// -------------------------------------------------------
// ═══════════════════════════════════════════════════════════════════════════
// == PHASE 3: NEW MODULE APIS - Garbage / Water / Safety / Petty Repair
// ═══════════════════════════════════════════════════════════════════════════

// ─── GARBAGE TASK APIS ────────────────────────────────────────────────────

// Get garbage tasks for a run instance
app.get('/api/obhs/garbage-tasks', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('garbage_tasks')
      .where('runInstanceId', '==', runInstanceId)
      .orderBy('scheduledTime')
      .get();
    const tasks = [];
    snapshot.forEach(doc => tasks.push(doc.data()));
    res.status(200).json({ success: true, count: tasks.length, tasks });
  } catch (error) {
    console.error('(GarbageTasks) Error:', error);
    res.status(500).json({ error: 'Failed to fetch garbage tasks', details: error.message });
  }
});

// Complete a garbage task (with evidence)
app.post('/api/obhs/garbage-tasks/complete', verifyToken, async (req, res) => {
  try {
    const { taskId, beforePhoto, afterPhoto, latitude, longitude } = req.body;
    if (!taskId || !beforePhoto || !afterPhoto) {
      return res.status(400).json({ error: 'taskId, beforePhoto, and afterPhoto are required.' });
    }
    const ref = db.collection('garbage_tasks').doc(taskId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Garbage task not found.' });
    }
    await ref.update({
      status: 'COMPLETED',
      beforePhoto,
      afterPhoto,
      gpsLatitude: latitude || null,
      gpsLongitude: longitude || null,
      completedAt: new Date().toISOString()
    });
    res.status(200).json({ success: true, message: 'Garbage task completed', taskId });
  } catch (error) {
    console.error('(GarbageComplete) Error:', error);
    res.status(500).json({ error: 'Failed to complete garbage task', details: error.message });
  }
});

// Get pre-terminal garbage tasks
app.get('/api/obhs/garbage-tasks/pre-terminal', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('garbage_tasks')
      .where('runInstanceId', '==', runInstanceId)
      .where('isPreTerminal', '==', true)
      .get();
    const tasks = [];
    snapshot.forEach(doc => tasks.push(doc.data()));
    res.status(200).json({ success: true, count: tasks.length, tasks });
  } catch (error) {
    console.error('(PreTerminalGarbage) Error:', error);
    res.status(500).json({ error: 'Failed to fetch pre-terminal tasks', details: error.message });
  }
});

// ─── WATER CHECK APIS ─────────────────────────────────────────────────────

// Get water checks for a run instance
app.get('/api/obhs/water-checks', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('water_checks')
      .where('runInstanceId', '==', runInstanceId)
      .orderBy('checkTime')
      .get();
    const checks = [];
    snapshot.forEach(doc => checks.push(doc.data()));
    res.status(200).json({ success: true, count: checks.length, checks });
  } catch (error) {
    console.error('(WaterChecks) Error:', error);
    res.status(500).json({ error: 'Failed to fetch water checks', details: error.message });
  }
});

// Submit a water check
app.post('/api/obhs/water-checks/submit', verifyToken, async (req, res) => {
  try {
    const { checkId, waterStatus, lowWaterAlert, wateringPointSchedule, photoUrl } = req.body;
    if (!checkId || !waterStatus) {
      return res.status(400).json({ error: 'checkId and waterStatus are required.' });
    }
    const validStatuses = ['full', 'low', 'empty'];
    if (!validStatuses.includes(waterStatus)) {
      return res.status(400).json({ error: 'Invalid waterStatus. Must be: full, low, or empty.' });
    }
    const ref = db.collection('water_checks').doc(checkId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Water check not found.' });
    }
    await ref.update({
      waterStatus,
      lowWaterAlert: lowWaterAlert || (waterStatus === 'empty') || (waterStatus === 'low'),
      wateringPointSchedule: wateringPointSchedule || null,
      photoUrl: photoUrl || null,
      status: 'COMPLETED',
      completedAt: new Date().toISOString()
    });
    res.status(200).json({ success: true, message: 'Water check submitted', checkId });
  } catch (error) {
    console.error('(WaterSubmit) Error:', error);
    res.status(500).json({ error: 'Failed to submit water check', details: error.message });
  }
});

// Get low water alerts
app.get('/api/obhs/water-checks/alerts', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('water_checks')
      .where('runInstanceId', '==', runInstanceId)
      .where('lowWaterAlert', '==', true)
      .get();
    const alerts = [];
    snapshot.forEach(doc => alerts.push(doc.data()));
    res.status(200).json({ success: true, count: alerts.length, alerts });
  } catch (error) {
    console.error('(WaterAlerts) Error:', error);
    res.status(500).json({ error: 'Failed to fetch water alerts', details: error.message });
  }
});

// ─── SAFETY CHECK APIS ────────────────────────────────────────────────────

// Get safety checks for a run instance
app.get('/api/obhs/safety-checks', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('safety_checks')
      .where('runInstanceId', '==', runInstanceId)
      .orderBy('scheduledTime')
      .get();
    const checks = [];
    snapshot.forEach(doc => checks.push(doc.data()));
    res.status(200).json({ success: true, count: checks.length, checks });
  } catch (error) {
    console.error('(SafetyChecks) Error:', error);
    res.status(500).json({ error: 'Failed to fetch safety checks', details: error.message });
  }
});

// Submit a safety check
app.post('/api/obhs/safety-checks/submit', verifyToken, async (req, res) => {
  try {
    const { checkId, fireExtinguisherStatus, fsdsStatus, cctvStatus, emergencyEquipmentStatus, photos, deficiencyReports, remarks } = req.body;
    if (!checkId) {
      return res.status(400).json({ error: 'checkId is required.' });
    }
    const ref = db.collection('safety_checks').doc(checkId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Safety check not found.' });
    }
    await ref.update({
      fireExtinguisherStatus: fireExtinguisherStatus || null,
      fsdsStatus: fsdsStatus || null,
      cctvStatus: cctvStatus || null,
      emergencyEquipmentStatus: emergencyEquipmentStatus || null,
      photos: photos || [],
      deficiencyReports: deficiencyReports || [],
      status: 'COMPLETED',
      remarks: remarks || null,
      completedAt: new Date().toISOString()
    });
    res.status(200).json({ success: true, message: 'Safety check submitted', checkId });
  } catch (error) {
    console.error('(SafetySubmit) Error:', error);
    res.status(500).json({ error: 'Failed to submit safety check', details: error.message });
  }
});

// Report a safety deficiency
app.post('/api/obhs/safety-checks/report-deficiency', verifyToken, async (req, res) => {
  try {
    const { checkId, deficiencyReport, photoUrl } = req.body;
    if (!checkId || !deficiencyReport) {
      return res.status(400).json({ error: 'checkId and deficiencyReport are required.' });
    }
    const ref = db.collection('safety_checks').doc(checkId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Safety check not found.' });
    }
    await ref.update({
      deficiencyReports: admin.firestore.FieldValue.arrayUnion({
        report: deficiencyReport,
        photoUrl: photoUrl || null,
        reportedAt: new Date().toISOString(),
        reportedBy: req.user.uid
      })
    });
    res.status(200).json({ success: true, message: 'Deficiency reported', checkId });
  } catch (error) {
    console.error('(SafetyDeficiency) Error:', error);
    res.status(500).json({ error: 'Failed to report deficiency', details: error.message });
  }
});

// ─── PETTY REPAIR APIS ────────────────────────────────────────────────────

// Get petty repairs for a run instance
app.get('/api/obhs/petty-repairs', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) {
      return res.status(400).json({ error: 'runInstanceId query parameter is required.' });
    }
    const snapshot = await db.collection('petty_repairs')
      .where('runInstanceId', '==', runInstanceId)
      .orderBy('inspectionTime')
      .get();
    const repairs = [];
    snapshot.forEach(doc => repairs.push(doc.data()));
    res.status(200).json({ success: true, count: repairs.length, repairs });
  } catch (error) {
    console.error('(PettyRepairs) Error:', error);
    res.status(500).json({ error: 'Failed to fetch petty repairs', details: error.message });
  }
});

// Submit a petty repair inspection
app.post('/api/obhs/petty-repairs/submit', verifyToken, async (req, res) => {
  try {
    const { repairId, items, remarks } = req.body;
    if (!repairId) {
      return res.status(400).json({ error: 'repairId is required.' });
    }
    const ref = db.collection('petty_repairs').doc(repairId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Petty repair record not found.' });
    }
    const updateData = {
      items: items || {},
      status: items ? 'COMPLETED' : 'PENDING',
      remarks: remarks || null,
      updatedAt: new Date().toISOString()
    };
    await ref.update(updateData);
    res.status(200).json({ success: true, message: 'Petty repair inspection submitted', repairId });
  } catch (error) {
    console.error('(RepairSubmit) Error:', error);
    res.status(500).json({ error: 'Failed to submit repair inspection', details: error.message });
  }
});

// Escalate a petty repair
app.post('/api/obhs/petty-repairs/escalate', verifyToken, async (req, res) => {
  try {
    const { repairId, escalatedTo } = req.body;
    if (!repairId) {
      return res.status(400).json({ error: 'repairId is required.' });
    }
    const ref = db.collection('petty_repairs').doc(repairId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Petty repair record not found.' });
    }
    await ref.update({
      isEscalated: true,
      escalatedTo: escalatedTo || 'supervisor',
      status: 'ESCALATED',
      updatedAt: new Date().toISOString()
    });
    res.status(200).json({ success: true, message: 'Repair escalated', repairId });
  } catch (error) {
    console.error('(RepairEscalate) Error:', error);
    res.status(500).json({ error: 'Failed to escalate repair', details: error.message });
  }
});


// ═══════════════════════════════════════════════════════════════════════════
// == PHASE 4: COMPLAINTS SLA, RATINGS, SUPERVISOR DASHBOARD, ANALYTICS
// ═══════════════════════════════════════════════════════════════════════════

// ─── MULTI-SOURCE RATINGS ─────────────────────────────────────────────────

// Submit a rating from any source (passenger, tte, supervisor, railway_official)
app.post('/api/obhs/ratings/submit', verifyToken, async (req, res) => {
  try {
    const { runInstanceId, coachNo, employeeId, source, rating, remarks } = req.body;
    if (!runInstanceId || !source || !rating) {
      return res.status(400).json({ error: 'runInstanceId, source, and rating are required.' });
    }
    const validSources = ['passenger', 'tte', 'supervisor', 'railway_official'];
    if (!validSources.includes(source)) {
      return res.status(400).json({ error: `Invalid source. Must be one of: ${validSources.join(', ')}` });
    }
    if (rating < 1 || rating > 5) {
      return res.status(400).json({ error: 'Rating must be between 1 and 5.' });
    }
    const ref = db.collection('ratings').doc();
    const data = {
      id: ref.id,
      runInstanceId,
      coachNo: coachNo || null,
      employeeId: employeeId || req.user.uid,
      taskId: req.body.taskId || null,
      source,
      rating: Number(rating),
      journeyId: runInstanceId,
      remarks: remarks || null,
      createdAt: new Date().toISOString()
    };
    await ref.set(data);
    res.status(201).json({ success: true, message: 'Rating submitted', ratingId: ref.id });
  } catch (error) {
    console.error('(RatingsSubmit) Error:', error);
    res.status(500).json({ error: 'Failed to submit rating', details: error.message });
  }
});

// Get ratings for an employee/work instance
app.get('/api/obhs/ratings', verifyToken, async (req, res) => {
  try {
    const { employeeId, runInstanceId, source } = req.query;
    let query = db.collection('ratings');
    if (employeeId) query = query.where('employeeId', '==', employeeId);
    if (runInstanceId) query = query.where('runInstanceId', '==', runInstanceId);
    if (source) query = query.where('source', '==', source);
    const snapshot = await query.get();
    const ratings = [];
    snapshot.forEach(doc => ratings.push(doc.data()));
    res.status(200).json({ success: true, count: ratings.length, ratings });
  } catch (error) {
    console.error('(RatingsGet) Error:', error);
    res.status(500).json({ error: 'Failed to fetch ratings', details: error.message });
  }
});

// Get performance score for an employee
app.get('/api/obhs/ratings/performance/:employeeId', verifyToken, async (req, res) => {
  try {
    const { employeeId } = req.params;
    const snapshot = await db.collection('ratings').where('employeeId', '==', employeeId).get();
    let totalRating = 0, count = 0;
    let sourceBreakdown = { passenger: { total: 0, count: 0 }, tte: { total: 0, count: 0 }, supervisor: { total: 0, count: 0 }, railway_official: { total: 0, count: 0 } };
    snapshot.forEach(doc => {
      const r = doc.data();
      totalRating += r.rating;
      count++;
      if (sourceBreakdown[r.source]) {
        sourceBreakdown[r.source].total += r.rating;
        sourceBreakdown[r.source].count++;
      }
    });
    const avgRating = count > 0 ? parseFloat((totalRating / count).toFixed(2)) : 0;
    const breakdown = {};
    for (const [key, val] of Object.entries(sourceBreakdown)) {
      breakdown[key] = {
        average: val.count > 0 ? parseFloat((val.total / val.count).toFixed(2)) : 0,
        count: val.count
      };
    }
    res.status(200).json({
      success: true,
      employeeId,
      averageRating: avgRating,
      totalRatings: count,
      sourceBreakdown: breakdown
    });
  } catch (error) {
    console.error('(PerformanceScore) Error:', error);
    res.status(500).json({ error: 'Failed to get performance score', details: error.message });
  }
});

// ─── COMPLAINTS SLA ───────────────────────────────────────────────────────

// Update complaint with SLA info and status
app.patch('/api/obhs/complaints/sla-update/:complaintId', verifyToken, async (req, res) => {
  try {
    const { complaintId } = req.params;
    const { status, resolutionNotes, resolutionPhotoUrl, escalatedTo } = req.body;
    const ref = db.collection('obhs_complaints').doc(complaintId);
    const doc = await ref.get();
    if (!doc.exists) {
      return res.status(404).json({ error: 'Complaint not found.' });
    }
    const complaint = doc.data();
    const createdAt = new Date(complaint.createdAt);
    const now = new Date();
    const resolutionTimeMinutes = Math.round((now - createdAt) / (1000 * 60));

    const slaMap = {
      'Cleaning': 30, 'Garbage': 20, 'Water': 30,
      'Petty Repair': 60, 'Toilet': 30, 'Electrical': 60,
      'Toilet Cleaning': 30, 'Coach Cleaning': 30
    };
    const slaMinutes = slaMap[complaint.category] || 30;
    const slaBreached = resolutionTimeMinutes > slaMinutes;

    const updateData = {
      status: status || complaint.status,
      updatedAt: now.toISOString(),
      slaMinutes: slaMinutes,
      resolutionTimeMinutes,
      slaBreached,
      resolutionNotes: resolutionNotes || complaint.resolutionNotes || null,
      resolutionPhotoUrl: resolutionPhotoUrl || complaint.resolutionPhotoUrl || null
    };
    if (escalatedTo) updateData.escalatedTo = escalatedTo;
    if (status === 'RESOLVED' || status === 'CLOSED') {
      updateData.resolvedAt = now.toISOString();
    }
    await ref.update(updateData);
    res.status(200).json({
      success: true,
      message: 'Complaint SLA updated',
      complaintId,
      resolutionTimeMinutes,
      slaBreached,
      slaMinutes
    });
  } catch (error) {
    console.error('(SLAUpdate) Error:', error);
    res.status(500).json({ error: 'Failed to update complaint SLA', details: error.message });
  }
});

// Get SLA compliance report
app.get('/api/obhs/complaints/sla-report', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    let query = db.collection('obhs_complaints');
    if (runInstanceId) query = query.where('runInstanceId', '==', runInstanceId);
    const snapshot = await query.get();
    let total = 0, withinSLA = 0, breached = 0, open = 0;
    const categoryStats = {};
    snapshot.forEach(doc => {
      const c = doc.data();
      total++;
      if (c.status === 'OPEN' || c.status === 'ESCALATED') open++;
      if (c.slaBreached) breached++;
      else if (c.status === 'RESOLVED' || c.status === 'CLOSED') withinSLA++;
      const cat = c.category || 'Other';
      if (!categoryStats[cat]) categoryStats[cat] = { total: 0, breached: 0, withinSLA: 0 };
      categoryStats[cat].total++;
      if (c.slaBreached) categoryStats[cat].breached++;
      else if (c.status === 'RESOLVED' || c.status === 'CLOSED') categoryStats[cat].withinSLA++;
    });
    const slaComplianceRate = (total - open) > 0 ? parseFloat((withinSLA / (total - open) * 100).toFixed(1)) : 0;
    res.status(200).json({
      success: true,
      total,
      open,
      resolved: withinSLA + breached,
      withinSLA,
      slaBreached: breached,
      slaComplianceRate,
      categoryStats
    });
  } catch (error) {
    console.error('(SLAReport) Error:', error);
    res.status(500).json({ error: 'Failed to get SLA report', details: error.message });
  }
});

// ─── SUPERVISOR DASHBOARD ─────────────────────────────────────────────────

// Get supervisor dashboard summary
app.get('/api/obhs/supervisor/dashboard', verifyToken, async (req, res) => {
  try {
    const allowedRoles = ['Admin', 'Supervisor', 'Company Master', 'Railway Supervisor'];
    if (!allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Access denied. Supervisor role required.' });
    }

    const activeRunsSnap = await db.collection('RunInstance').where('status', '==', 'Active').get();
    const activeTrains = activeRunsSnap.size;
    const activeWorkersSet = new Set();
    activeRunsSnap.forEach(doc => {
      const data = doc.data();
      if (data.coaches) data.coaches.forEach(c => { if (c.workerId) activeWorkersSet.add(c.workerId); });
    });

    const overdueHeadersSnap = await db.collection('task_headers').where('status', '==', 'OVERDUE').get();
    const openDetailsSnap = await db.collection('task_details').where('status', 'in', ['OPEN', 'OVERDUE']).get();
    let missedTasks = 0, overdueTasks = 0;
    openDetailsSnap.forEach(doc => {
      if (doc.data().status === 'OVERDUE') overdueTasks++;
      else missedTasks++;
    });

    const escalatedDetailsSnap = await db.collection('task_details').where('status', '==', 'ESCALATED').get();
    const escalatedTasks = escalatedDetailsSnap.size;

    const openComplaintsSnap = await db.collection('obhs_complaints').where('status', '==', 'OPEN').get();
    const openComplaints = openComplaintsSnap.size;

    const ratingsSnap = await db.collection('ratings').get();
    let totalRating = 0, ratingCount = 0;
    ratingsSnap.forEach(doc => { totalRating += doc.data().rating || 0; ratingCount++; });
    const averageRating = ratingCount > 0 ? parseFloat((totalRating / ratingCount).toFixed(2)) : 0;

    const lowWaterSnap = await db.collection('water_checks').where('status', '==', 'PENDING').where('lowWaterAlert', '==', true).get();
    const waterIssues = lowWaterSnap.size;

    const safetySnap = await db.collection('safety_checks').where('status', '==', 'PENDING').get();
    let safetyIssues = 0;
    safetySnap.forEach(doc => {
      const d = doc.data();
      if (d.fireExtinguisherStatus === 'deficient' || d.fsdsStatus === 'deficient') safetyIssues++;
    });

    res.status(200).json({
      success: true,
      dashboard: {
        activeTrains,
        activeWorkers: activeWorkersSet.size,
        missedTasks,
        overdueTasks,
        escalatedTasks,
        openComplaints,
        averageRating,
        waterIssues,
        safetyIssues,
        lastUpdated: new Date().toISOString()
      }
    });
  } catch (error) {
    console.error('(SupervisorDashboard) Error:', error);
    res.status(500).json({ error: 'Failed to get dashboard', details: error.message });
  }
});

// ─── ANALYTICS ENDPOINTS ──────────────────────────────────────────────────

// Janitor performance analytics
app.get('/api/obhs/analytics/janitor-performance', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) return res.status(400).json({ error: 'runInstanceId is required.' });

    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) return res.status(404).json({ error: 'Run instance not found.' });
    const runData = runDoc.data();
    const janitors = {};

    for (const coach of (runData.coaches || [])) {
      if (coach.workerId && !janitors[coach.workerId]) {
        janitors[coach.workerId] = { janitorId: coach.workerId, janitorName: coach.workerName || 'Unknown', tasksCompleted: 0, tasksMissed: 0, tasksOverdue: 0, totalTasks: 0, averageRating: 0, completionPercentage: 0, coachCleanlinessScore: 0 };
      }
    }

    const detailsSnap = await db.collection('task_details').where('runInstanceId', '==', runInstanceId).get();
    detailsSnap.forEach(doc => {
      const d = doc.data();
      const wId = d.workerId;
      if (wId && janitors[wId]) {
        janitors[wId].totalTasks++;
        if (d.status === 'COMPLETED') janitors[wId].tasksCompleted++;
        else if (d.status === 'OVERDUE') janitors[wId].tasksOverdue++;
        else if (d.status === 'PLANNED') janitors[wId].tasksMissed++;
      }
    });

    const ratingsSnap = await db.collection('ratings').where('runInstanceId', '==', runInstanceId).get();
    const ratingMap = {}, ratingCountMap = {};
    ratingsSnap.forEach(doc => {
      const r = doc.data();
      const eid = r.employeeId;
      if (eid) { ratingMap[eid] = (ratingMap[eid] || 0) + r.rating; ratingCountMap[eid] = (ratingCountMap[eid] || 0) + 1; }
    });

    const result = Object.values(janitors).map(j => ({
      ...j,
      averageRating: ratingCountMap[j.janitorId] ? parseFloat((ratingMap[j.janitorId] / ratingCountMap[j.janitorId]).toFixed(2)) : 0,
      completionPercentage: j.totalTasks > 0 ? parseFloat(((j.tasksCompleted / j.totalTasks) * 100).toFixed(1)) : 0
    }));

    res.status(200).json({ success: true, performance: result });
  } catch (error) {
    console.error('(JanitorPerformance) Error:', error);
    res.status(500).json({ error: 'Failed to get janitor performance', details: error.message });
  }
});

// Coach cleanliness analytics
app.get('/api/obhs/analytics/coach-cleanliness', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) return res.status(400).json({ error: 'runInstanceId is required.' });

    const detailsSnap = await db.collection('task_details').where('runInstanceId', '==', runInstanceId).get();
    const coaches = {};
    detailsSnap.forEach(doc => {
      const d = doc.data();
      const cn = d.coachNo;
      if (!cn) return;
      if (!coaches[cn]) coaches[cn] = { coachNo: cn, totalToiletTasks: 0, toiletCompletions: 0, totalTasks: 0, completedTasks: 0 };
      coaches[cn].totalTasks++;
      if (d.status === 'COMPLETED') {
        coaches[cn].completedTasks++;
        if (d.taskType === 'toilet_cleaning') coaches[cn].toiletCompletions++;
      }
      if (d.taskType === 'toilet_cleaning') coaches[cn].totalToiletTasks++;
    });

    const garbageSnap = await db.collection('garbage_tasks').where('runInstanceId', '==', runInstanceId).get();
    const garbageMap = {};
    garbageSnap.forEach(doc => { const g = doc.data(); if (g.coachNo) garbageMap[g.coachNo] = (garbageMap[g.coachNo] || 0) + 1; });

    const waterSnap = await db.collection('water_checks').where('runInstanceId', '==', runInstanceId).where('lowWaterAlert', '==', true).get();
    const waterIssuesMap = {};
    waterSnap.forEach(doc => { const w = doc.data(); if (w.coachNo) waterIssuesMap[w.coachNo] = (waterIssuesMap[w.coachNo] || 0) + 1; });

    const result = Object.values(coaches).map(c => ({ ...c, cleanlinessScore: c.totalTasks > 0 ? parseFloat(((c.completedTasks / c.totalTasks) * 100).toFixed(1)) : 0, garbageIssues: garbageMap[c.coachNo] || 0, waterIssues: waterIssuesMap[c.coachNo] || 0 }));
    res.status(200).json({ success: true, coaches: result });
  } catch (error) {
    console.error('(CoachCleanliness) Error:', error);
    res.status(500).json({ error: 'Failed to get coach cleanliness', details: error.message });
  }
});

// Attendance compliance analytics
app.get('/api/obhs/analytics/attendance-compliance', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) return res.status(400).json({ error: 'runInstanceId is required.' });

    const snapshot = await db.collection('obhs_attendance').where('runInstanceId', '==', runInstanceId).get();
    let total = 0, onTime = 0, late = 0;
    snapshot.forEach(doc => {
      const d = doc.data();
      total++;
      const startLate = d.startAttendance && d.startAttendance.isLate;
      const midLate = d.midAttendance && d.midAttendance.isLate;
      const endLate = d.endAttendance && d.endAttendance.isLate;
      if (startLate || midLate || endLate) late++;
      else onTime++;
    });

    res.status(200).json({
      success: true, total, onTime, late,
      complianceRate: total > 0 ? parseFloat(((onTime / total) * 100).toFixed(1)) : 0
    });
  } catch (error) {
    console.error('(AttendanceCompliance) Error:', error);
    res.status(500).json({ error: 'Failed to get attendance compliance', details: error.message });
  }
});

// Task completion percentage
app.get('/api/obhs/analytics/task-completion', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    if (!runInstanceId) return res.status(400).json({ error: 'runInstanceId is required.' });

    const snapshot = await db.collection('task_details').where('runInstanceId', '==', runInstanceId).get();
    let total = 0, completed = 0, overdue = 0, planned = 0, open = 0, escalated = 0;
    snapshot.forEach(doc => {
      const d = doc.data(); total++;
      switch (d.status) { case 'COMPLETED': completed++; break; case 'OVERDUE': overdue++; break; case 'PLANNED': planned++; break; case 'OPEN': open++; break; case 'ESCALATED': escalated++; break; }
    });

    res.status(200).json({
      success: true, total, completed, overdue, planned, open, escalated,
      completionRate: total > 0 ? parseFloat(((completed / total) * 100).toFixed(1)) : 0,
      overdueRate: total > 0 ? parseFloat(((overdue / total) * 100).toFixed(1)) : 0
    });
  } catch (error) {
    console.error('(TaskCompletion) Error:', error);
    res.status(500).json({ error: 'Failed to get task completion', details: error.message });
  }
});

// Passenger rating trend
app.get('/api/obhs/analytics/passenger-rating-trend', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    const query = db.collection('obhs_feedbacks');
    const snapshot = runInstanceId ? await query.where('runInstanceId', '==', runInstanceId).get() : await query.get();
    const dailyRates = {};
    snapshot.forEach(doc => {
      const f = doc.data();
      if (f.feedbackType === 'PASSENGER' && f.overallRating) {
        const date = f.date || new Date(f.createdAt).toISOString().split('T')[0];
        if (!dailyRates[date]) dailyRates[date] = { total: 0, count: 0 };
        dailyRates[date].total += f.overallRating;
        dailyRates[date].count++;
      }
    });
    const trend = Object.entries(dailyRates).sort(([a], [b]) => a.localeCompare(b)).map(([date, data]) => ({ date, averageRating: parseFloat((data.total / data.count).toFixed(2)), count: data.count }));
    const overallAvg = trend.length > 0 ? parseFloat((trend.reduce((s, d) => s + d.averageRating, 0) / trend.length).toFixed(2)) : 0;
    res.status(200).json({ success: true, trend, overallAverageRating: overallAvg });
  } catch (error) {
    console.error('(RatingTrend) Error:', error);
    res.status(500).json({ error: 'Failed to get rating trend', details: error.message });
  }
});

// Penalty risk report
app.get('/api/obhs/analytics/penalty-risk', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.query;
    const risks = [];

    let detailQuery = db.collection('task_details').where('status', '==', 'OVERDUE');
    if (runInstanceId) detailQuery = detailQuery.where('runInstanceId', '==', runInstanceId);
    const overdueSnap = await detailQuery.get();
    const overdueByWorker = {};
    overdueSnap.forEach(doc => {
      const d = doc.data();
      const key = d.workerId || 'unknown';
      if (!overdueByWorker[key]) overdueByWorker[key] = { workerId: key, workerName: d.workerName || 'Unknown', overdueCount: 0, missedCount: 0, penaltyAmount: 0 };
      overdueByWorker[key].overdueCount++;
      overdueByWorker[key].penaltyAmount += 50;
    });

    let attendQuery = db.collection('obhs_attendance');
    if (runInstanceId) attendQuery = attendQuery.where('runInstanceId', '==', runInstanceId);
    const attendSnap = await attendQuery.get();
    attendSnap.forEach(doc => {
      const d = doc.data();
      const key = d.uid ? d.uid.split('_')[1] : 'unknown';
      if (!d.startAttendance || !d.endAttendance) {
        if (!overdueByWorker[key]) overdueByWorker[key] = { workerId: key, workerName: 'Unknown', overdueCount: 0, missedCount: 0, penaltyAmount: 0 };
        overdueByWorker[key].missedCount++;
        overdueByWorker[key].penaltyAmount += 100;
      }
    });

    Object.values(overdueByWorker).forEach(w => risks.push(w));
    res.status(200).json({ success: true, riskReport: risks, totalPenaltyRisk: risks.reduce((s, r) => s + r.penaltyAmount, 0) });
  } catch (error) {
    console.error('(PenaltyRisk) Error:', error);
    res.status(500).json({ error: 'Failed to get penalty risk report', details: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// == COMPREHENSIVE REPORT API (WITH RBAC)
// ═══════════════════════════════════════════════════════════════════════════

app.get('/api/obhs/reports/comprehensive/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const userRole = req.user.role ? req.user.role.toUpperCase() : '';
    const userTrainId = req.user.trainId;
    const userTrainIds = req.user.trainIds || [];

    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) return res.status(404).json({ error: 'RunInstance not found' });
    const runData = runDoc.data();

    // ─── RBAC ENFORCEMENT ───
    if (userRole === 'CTS' || userRole === 'CONTRACTOR SUPERVISOR') {
      if (runData.parentTrainId !== userTrainId) {
        return res.status(403).json({ error: 'Access denied. You can only view reports for your assigned train.' });
      }
    } else if (userRole === 'RAILWAY SUPERVISOR') {
      if (!userTrainIds.includes(runData.parentTrainId) && runData.parentTrainId !== userTrainId) {
        return res.status(403).json({ error: 'Access denied. Train not in your assigned trains list.' });
      }
    } else if (!userRole.includes('ADMIN') && !userRole.includes('COMPANY MASTER')) {
      return res.status(403).json({ error: 'Access denied. Insufficient privileges.' });
    }

    // ─── AGGREGATION ───
    const attendSnap = await db.collection('obhs_attendance').where('runInstanceId', '==', runInstanceId).get();
    const attendance = { total: 0, present: 0, late: 0, missed: 0, records: [] };
    attendSnap.forEach(doc => {
      const d = doc.data();
      attendance.total++;
      if (d.status === 'LATE') attendance.late++;
      else if (d.status === 'PRESENT') attendance.present++;
      attendance.records.push({ workerName: d.workerName, status: d.status, lateByMinutes: d.lateByMinutes });
    });

    const taskSnap = await db.collection('obhs_tasks').where('runInstanceId', '==', runInstanceId).get();
    const tasks = { total: 0, completed: 0, overdue: 0, missed: 0 };
    taskSnap.forEach(doc => {
      const d = doc.data();
      tasks.total++;
      if (d.status === 'COMPLETED' || d.status === 'APPROVED') tasks.completed++;
      else if (d.status === 'OVERDUE') tasks.overdue++;
      else tasks.missed++;
    });

    const complaintSnap = await db.collection('obhs_complaints').where('runInstanceId', '==', runInstanceId).get();
    const complaints = { total: 0, open: 0, resolved: 0 };
    complaintSnap.forEach(doc => {
      const d = doc.data();
      complaints.total++;
      if (d.status === 'OPEN') complaints.open++;
      else complaints.resolved++;
    });

    const ratingSnap = await db.collection('obhs_feedbacks').where('runInstanceId', '==', runInstanceId).get();
    const ratings = { total: 0, sum: 0, average: 0 };
    ratingSnap.forEach(doc => {
      const d = doc.data();
      ratings.total++;
      ratings.sum += (d.passengerScore || d.officialScore || 0);
    });
    if (ratings.total > 0) ratings.average = (ratings.sum / ratings.total).toFixed(2);

    res.status(200).json({
      success: true,
      runInstanceId,
      trainId: runData.parentTrainId,
      date: runData.departureDate,
      attendance,
      tasks,
      complaints,
      ratings
    });
  } catch (error) {
    console.error('(ComprehensiveReport) Error:', error);
    res.status(500).json({ error: 'Failed to generate comprehensive report', details: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// == NEW ARCHITECTURE: MCC OBHS CONTRACTOR MANAGEMENT SYSTEM
// ═══════════════════════════════════════════════════════════════════════════

// ─── ROLE HIERARCHY ───────────────────────────────────────────────────────
const ROLE_HIERARCHY = {
  'CM': { level: 5, label: 'Contractor Master', manages: ['CA', 'CTS', 'CS', 'janitor', 'attendant'] },
  'CA': { level: 4, label: 'Contractor Admin', manages: ['CTS', 'CS', 'janitor', 'attendant'] },
  'CTS': { level: 3, label: 'Contractor Train Supervisor', manages: ['CS', 'janitor', 'attendant'] },
  'CS': { level: 2, label: 'Contractor Supervisor', manages: ['janitor', 'attendant'] },
  'janitor': { level: 1, label: 'Janitor', manages: [] },
  'attendant': { level: 1, label: 'Coach Attendant', manages: [] }
};

function hasRolePermission(actorRole, targetRole) {
  const actor = ROLE_HIERARCHY[actorRole];
  if (!actor) return false;
  if (actorRole === targetRole) return true;
  return actor.manages.includes(targetRole);
}

// ─── AUDIT LOG HELPER ─────────────────────────────────────────────────────
async function logAudit({ action, entityType, entityId, actorId, actorRole, actorName, oldValue, newValue, details }) {
  try {
    const ref = db.collection('audit_logs').doc();
    const entry = {
      logId: ref.id,
      action,
      entityType,
      entityId,
      actorId: actorId || 'system',
      actorRole: actorRole || 'system',
      actorName: actorName || 'System',
      oldValue: oldValue || null,
      newValue: newValue || null,
      details: details || '',
      ip: null,
      userAgent: null,
      timestamp: new Date().toISOString()
    };
    await ref.set(entry);
    return ref.id;
  } catch (err) {
    console.error('(AuditLog) Error:', err.message);
  }
}

// ─── AUTO-CLOSE PARENT TASK HEADER ───────────────────────────────────────
async function autoCloseParentTask(detailRef, detailDoc) {
  try {
    const detailData = detailDoc.data();
    if (!detailData.headerId) return;

    const headerRef = db.collection('task_headers').doc(detailData.headerId);
    const headerDoc = await headerRef.get();
    if (!headerDoc.exists) return;

    const headerData = headerDoc.data();
    const childIds = headerData.childTaskIds || [];

    // Check all children
    let allCompleted = true;
    let anyOverdue = false;
    for (const cId of childIds) {
      const cDoc = await db.collection('task_details').doc(cId).get();
      if (!cDoc.exists) { allCompleted = false; continue; }
      const cData = cDoc.data();
      if (cData.status === 'OVERDUE') anyOverdue = true;
      if (cData.status !== 'COMPLETED') allCompleted = false;
    }

    let newStatus = headerData.status;
    if (allCompleted) {
      newStatus = 'COMPLETED';
    } else if (anyOverdue) {
      newStatus = 'PARTIAL';
    }

    if (newStatus !== headerData.status) {
      await headerRef.update({
        status: newStatus,
        updatedAt: new Date().toISOString()
      });
    }
  } catch (err) {
    console.error('(AutoCloseParent) Error:', err.message);
  }
}

// ─── GENERATE CLOSURE TASKS ──────────────────────────────────────────────
async function generateClosureTasks(runInstanceId, arrivalTime, runData) {
  try {
    const batch = db.batch();
    let count = 0;
    const coaches = runData.coaches || [];

    const closureTasks = [
      { code: 'FINAL_INSPECTION', name: 'Final Coach Inspection', category: 'closure' },
      { code: 'FINAL_GARBAGE', name: 'Final Garbage Clearance Verification', category: 'closure' },
      { code: 'LOST_FOUND', name: 'Lost & Found Check', category: 'closure' },
      { code: 'INVENTORY', name: 'Inventory & Consumable Reconciliation', category: 'closure' },
      { code: 'CLOSURE_REPORT', name: 'Journey Closure Report', category: 'closure' }
    ];

    for (const coach of coaches) {
      if (!coach.workerId) continue;
      const coachNo = coach.coachPosition || coach.coachNo || 'N/A';

      for (const task of closureTasks) {
        const taskId = `${runInstanceId}_${task.code}_${coachNo}`;
        const ref = db.collection('closure_tasks').doc(taskId);
        batch.set(ref, {
          taskId, runInstanceId, coachNo,
          taskCode: task.code, taskName: task.name,
          category: task.category,
          workerId: coach.workerId, workerName: coach.workerName || 'Unknown',
          status: 'PENDING',
          beforePhoto: null, afterPhoto: null,
          gpsLatitude: null, gpsLongitude: null,
          supervisorVerified: false,
          completedAt: null, createdAt: new Date().toISOString()
        });
        count++;
      }
    }

    if (count > 0) {
      await batch.commit();
      console.log(`[Closure] Generated ${count} closure tasks for ${runInstanceId}`);
    }
  } catch (err) {
    console.error('(GenerateClosure) Error:', err.message);
  }
}

// ─── SEED DEFAULT TASK MASTERS ────────────────────────────────────────────
async function seedTaskMasters() {
  try {
    const existing = await db.collection('task_masters').limit(1).get();
    if (!existing.empty) {
      console.log('[TaskMasters] Already seeded, skipping.');
      return;
    }

    const masters = [
      // ── TOILET CLEANING (Janitor) ──
      {
        taskCode: 'TOILET_CLEAN_06', taskName: 'Toilet Cleaning 06:00',
        category: 'cleaning', subCategory: 'toilet',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'toilet_cleaning',
        frequencyRules: [
          { startHour: 6, endHour: 9, intervalMinutes: 60, label: 'Peak Morning' }
        ],
        scheduledTime: '06:00',
        checklistTemplate: [
          { item: 'toilet_pan_cleaned', label: 'Toilet Pan Cleaned', required: true },
          { item: 'washbasin_cleaned', label: 'Washbasin Cleaned', required: true },
          { item: 'mirror_cleaned', label: 'Mirror Cleaned', required: false },
          { item: 'floor_wiped', label: 'Floor Wiped', required: true },
          { item: 'deodorant_applied', label: 'Deodorant Applied', required: false }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'TOILET_CLEAN_07', taskName: 'Toilet Cleaning 07:00',
        category: 'cleaning', subCategory: 'toilet',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'toilet_cleaning',
        frequencyRules: [
          { startHour: 6, endHour: 9, intervalMinutes: 60, label: 'Peak Morning' }
        ],
        scheduledTime: '07:00',
        checklistTemplate: [
          { item: 'toilet_pan_cleaned', label: 'Toilet Pan Cleaned', required: true },
          { item: 'washbasin_cleaned', label: 'Washbasin Cleaned', required: true },
          { item: 'floor_wiped', label: 'Floor Wiped', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'TOILET_CLEAN_08', taskName: 'Toilet Cleaning 08:00',
        category: 'cleaning', subCategory: 'toilet',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'toilet_cleaning',
        frequencyRules: [
          { startHour: 6, endHour: 9, intervalMinutes: 60, label: 'Peak Morning' }
        ],
        scheduledTime: '08:00',
        checklistTemplate: [
          { item: 'toilet_pan_cleaned', label: 'Toilet Pan Cleaned', required: true },
          { item: 'washbasin_cleaned', label: 'Washbasin Cleaned', required: true },
          { item: 'floor_wiped', label: 'Floor Wiped', required: true }
        ],
        active: true, isSystem: true
      },
      // ── COACH CLEANING (Janitor) Normal hours every 2hr ──
      {
        taskCode: 'COACH_CLEAN_09', taskName: 'Coach Cleaning 09:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '09:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true },
          { item: 'windows_cleaned', label: 'Windows Cleaned', required: false }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_11', taskName: 'Coach Cleaning 11:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '11:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_13', taskName: 'Coach Cleaning 13:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '13:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_15', taskName: 'Coach Cleaning 15:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '15:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_17', taskName: 'Coach Cleaning 17:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '17:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_19', taskName: 'Coach Cleaning 19:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 9, endHour: 20, intervalMinutes: 120, label: 'Normal Hours' }
        ],
        scheduledTime: '19:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      // ── EVENING PEAK (Janitor) every 1hr ──
      {
        taskCode: 'COACH_CLEAN_20', taskName: 'Coach Cleaning 20:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 20, endHour: 22, intervalMinutes: 60, label: 'Evening Peak' }
        ],
        scheduledTime: '20:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'COACH_CLEAN_21', taskName: 'Coach Cleaning 21:00',
        category: 'cleaning', subCategory: 'compartment',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'coach_cleaning',
        frequencyRules: [
          { startHour: 20, endHour: 22, intervalMinutes: 60, label: 'Evening Peak' }
        ],
        scheduledTime: '21:00',
        checklistTemplate: [
          { item: 'floor_swept', label: 'Floor Swept', required: true },
          { item: 'seats_dusted', label: 'Seats Dusted', required: true },
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true }
        ],
        active: true, isSystem: true
      },
      // ── GARBAGE MANAGEMENT (Janitor) ──
      {
        taskCode: 'GARBAGE_07', taskName: 'Garbage Collection 07:00',
        category: 'garbage', subCategory: 'collection',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 4, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'garbage',
        frequencyRules: [], scheduledTime: '07:00',
        checklistTemplate: [
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true },
          { item: 'garbage_bagged', label: 'Garbage Bagged', required: true },
          { item: 'area_clean', label: 'Area Clean After Collection', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'GARBAGE_13', taskName: 'Garbage Collection 13:00',
        category: 'garbage', subCategory: 'collection',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 4, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'garbage',
        frequencyRules: [], scheduledTime: '13:00',
        checklistTemplate: [
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true },
          { item: 'garbage_bagged', label: 'Garbage Bagged', required: true },
          { item: 'area_clean', label: 'Area Clean After Collection', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'GARBAGE_19', taskName: 'Garbage Collection 19:00',
        category: 'garbage', subCategory: 'collection',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 4, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'garbage',
        frequencyRules: [], scheduledTime: '19:00',
        checklistTemplate: [
          { item: 'dustbin_emptied', label: 'Dustbin Emptied', required: true },
          { item: 'garbage_bagged', label: 'Garbage Bagged', required: true },
          { item: 'area_clean', label: 'Area Clean After Collection', required: true }
        ],
        active: true, isSystem: true
      },
      // ── WATER CHECK (Supervisor) ──
      {
        taskCode: 'WATER_CHECK_08', taskName: 'Morning Water Check 08:00',
        category: 'water', subCategory: 'check',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'water',
        frequencyRules: [], scheduledTime: '08:00',
        checklistTemplate: [
          { item: 'tank_level', label: 'Water Tank Level', required: true },
          { item: 'chiller_status', label: 'Water Chiller Status', required: true },
          { item: 'cooler_status', label: 'Cooler Status', required: false }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'WATER_CHECK_16', taskName: 'Evening Water Check 16:00',
        category: 'water', subCategory: 'check',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'water',
        frequencyRules: [], scheduledTime: '16:00',
        checklistTemplate: [
          { item: 'tank_level', label: 'Water Tank Level', required: true },
          { item: 'chiller_status', label: 'Water Chiller Status', required: true },
          { item: 'cooler_status', label: 'Cooler Status', required: false }
        ],
        active: true, isSystem: true
      },
      // ── SAFETY INSPECTION (Supervisor) ──
      {
        taskCode: 'SAFETY_10', taskName: 'Morning Safety Inspection 10:00',
        category: 'safety', subCategory: 'inspection',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 1, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'safety',
        frequencyRules: [], scheduledTime: '10:00',
        checklistTemplate: [
          { item: 'fire_extinguisher', label: 'Fire Extinguisher Status', required: true },
          { item: 'cctv', label: 'CCTV Functioning', required: true },
          { item: 'fsds', label: 'FSDS System', required: true },
          { item: 'emergency_equipment', label: 'Emergency Equipment', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'SAFETY_18', taskName: 'Evening Safety Inspection 18:00',
        category: 'safety', subCategory: 'inspection',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 1, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'safety',
        frequencyRules: [], scheduledTime: '18:00',
        checklistTemplate: [
          { item: 'fire_extinguisher', label: 'Fire Extinguisher Status', required: true },
          { item: 'cctv', label: 'CCTV Functioning', required: true },
          { item: 'fsds', label: 'FSDS System', required: true },
          { item: 'emergency_equipment', label: 'Emergency Equipment', required: true }
        ],
        active: true, isSystem: true
      },
      // ── PETTY REPAIR (Supervisor) ──
      {
        taskCode: 'REPAIR_10', taskName: 'Morning Petty Repair Check 10:00',
        category: 'repair', subCategory: 'inspection',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'repair',
        frequencyRules: [], scheduledTime: '10:00',
        checklistTemplate: [
          { item: 'latches', label: 'Door Latches', required: true },
          { item: 'windows', label: 'Windows', required: true },
          { item: 'fans', label: 'Fans', required: true },
          { item: 'lights', label: 'Lights', required: true }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'REPAIR_18', taskName: 'Evening Petty Repair Check 18:00',
        category: 'repair', subCategory: 'inspection',
        assignedRole: 'CS', applicableCoachTypes: ['all'],
        priority: 2, requiresSupervisorVerification: true,
        isComplaintDriven: false, conflictGroup: 'repair',
        frequencyRules: [], scheduledTime: '18:00',
        checklistTemplate: [
          { item: 'latches', label: 'Door Latches', required: true },
          { item: 'windows', label: 'Windows', required: true },
          { item: 'fans', label: 'Fans', required: true },
          { item: 'lights', label: 'Lights', required: true }
        ],
        active: true, isSystem: true
      },
      // ── ATTENDANT TASKS (AC Coaches Only) ──
      {
        taskCode: 'LINEN_DIST', taskName: 'Linen Distribution',
        category: 'passenger_service', subCategory: 'linen',
        assignedRole: 'attendant', applicableCoachTypes: ['ac'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'linen',
        frequencyRules: [], scheduledTime: '06:00',
        checklistTemplate: [
          { item: 'bedroll', label: 'Bedroll Distributed', required: true },
          { item: 'blanket', label: 'Blanket Distributed', required: true },
          { item: 'pillow', label: 'Pillow Distributed', required: true },
          { item: 'towel', label: 'Towel Distributed', required: false }
        ],
        active: true, isSystem: true
      },
      {
        taskCode: 'BEDROLL_COLLECT', taskName: 'Bedroll Collection',
        category: 'passenger_service', subCategory: 'linen',
        assignedRole: 'attendant', applicableCoachTypes: ['ac'],
        priority: 3, requiresSupervisorVerification: false,
        isComplaintDriven: false, conflictGroup: 'linen',
        frequencyRules: [], scheduledTime: '21:00',
        checklistTemplate: [
          { item: 'bedroll_collected', label: 'Bedroll Collected', required: true },
          { item: 'count_verified', label: 'Count Verified', required: true }
        ],
        active: true, isSystem: true
      },
      // ── NIGHT TASKS (Complaint-driven only) ──
      {
        taskCode: 'NIGHT_CLEAN_REQUEST', taskName: 'Night Cleaning (On Request)',
        category: 'cleaning', subCategory: 'night',
        assignedRole: 'janitor', applicableCoachTypes: ['all'],
        priority: 4, requiresSupervisorVerification: false,
        isComplaintDriven: true, conflictGroup: 'night_cleaning',
        frequencyRules: [
          { startHour: 22, endHour: 6, intervalMinutes: 0, label: 'Night Hours (Request Only)' }
        ],
        scheduledTime: null,
        checklistTemplate: [
          { item: 'toilet_cleaned', label: 'Toilet Cleaned', required: true },
          { item: 'area_wiped', label: 'Area Wiped', required: true }
        ],
        active: true, isSystem: true
      }
    ];

    const batch = db.batch();
    for (const master of masters) {
      const ref = db.collection('task_masters').doc(master.taskCode);
      batch.set(ref, { ...master, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() });
    }
    await batch.commit();
    console.log(`[TaskMasters] Seeded ${masters.length} default task masters`);
  } catch (err) {
    console.error('(SeedTaskMasters) Error:', err.message);
  }
}

// ─── GENERATE TASKS FROM MASTERS (Rule-driven, called on activation) ─────
async function generateTasksFromMasters(runData) {
  try {
    const mastersSnap = await db.collection('task_masters')
      .where('active', '==', true)
      .where('isComplaintDriven', '==', false)
      .get();

    if (mastersSnap.empty) return;

    const masters = [];
    mastersSnap.forEach(doc => masters.push(doc.data()));

    const batch = db.batch();
    let count = 0;
    const departureDate = runData.departureDate || new Date().toISOString().split('T')[0];
    const runId = runData.runInstanceId;
    const actualDeparture = runData.actualDeparture || new Date().toISOString();

    // Calculate journey timeline from actual departure
    const departureTime = new Date(actualDeparture);
    const depHour = departureTime.getHours();
    const depMinutes = departureTime.getMinutes();

    for (const coach of (runData.coaches || [])) {
      if (!coach.workerId) continue;
      const coachNo = coach.coachPosition || coach.coachNo || 'N/A';
      const workerId = coach.workerId;
      const workerName = coach.workerName || 'Unknown';
      const isAC = coach.coachType === 'AC' || coach.coachType === 'ac';

      for (const master of masters) {
        // Check if this task applies to this coach type
        if (!master.applicableCoachTypes.includes('all')) {
          if (isAC && !master.applicableCoachTypes.includes('ac')) continue;
          if (!isAC && !master.applicableCoachTypes.includes('non_ac')) continue;
        }

        // Check role match: task assigned to this worker's role
        const workerRole = coach.workerRole || 'janitor';
        if (master.assignedRole !== workerRole && master.assignedRole !== 'all') continue;

        // Build scheduled time from master or compute from frequency
        const taskTime = master.scheduledTime;
        if (!taskTime) continue;

        const [th, tm] = taskTime.split(':').map(Number);
        const taskHour = th;
        const taskMinute = tm || 0;

        // Skip tasks scheduled before actual departure (train hasn't left yet)
        if (taskHour < depHour || (taskHour === depHour && taskMinute < depMinutes)) continue;

        // Create task instance (parent)
        const taskId = `${runId}_${master.taskCode}_${coachNo}`;
        const taskRef = db.collection('task_instances').doc(taskId);
        batch.set(taskRef, {
          taskId, runInstanceId: runId, trainNo: runData.trainNo,
          taskMasterCode: master.taskCode, taskName: master.taskName,
          category: master.category, subCategory: master.subCategory,
          coachNo, workerId, workerName, workerRole,
          scheduledTime: taskTime, scheduledDate: departureDate,
          status: 'PLANNED',
          priority: master.priority,
          requiresSupervisorVerification: master.requiresSupervisorVerification,
          conflictGroup: master.conflictGroup,
          checklistTemplate: master.checklistTemplate,
          isParentTask: true, childTaskIds: [],
          beforePhoto: null, afterPhoto: null,
          gpsLatitude: null, gpsLongitude: null,
          supervisorVerified: false,
          supervisorId: null, supervisorRemarks: null,
          exceptionReason: null,
          checklistResponses: [],
          completedAt: null, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
        });

        // Create child task for this specific coach
        const childId = `${taskId}_child`;
        const childRef = db.collection('task_instances').doc(childId);
        batch.set(childRef, {
          taskId: childId, parentTaskId: taskId, runInstanceId: runId,
          taskMasterCode: master.taskCode, taskName: master.taskName,
          category: master.category, subCategory: master.subCategory,
          coachNo, workerId, workerName, workerRole,
          scheduledTime: taskTime, scheduledDate: departureDate,
          status: 'PLANNED', priority: master.priority,
          isParentTask: false,
          beforePhoto: null, afterPhoto: null,
          gpsLatitude: null, gpsLongitude: null,
          supervisorVerified: false,
          supervisorId: null, supervisorRemarks: null,
          exceptionReason: null,
          checklistResponses: [],
          completedAt: null, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
        });

        batch.update(taskRef, {
          childTaskIds: admin.firestore.FieldValue.arrayUnion(childId)
        });
        count += 2; // parent + child
      }
    }

    if (count > 0) {
      await batch.commit();
      console.log(`[TaskMasters] Generated ${count} task instances from masters for ${runId}`);
    }
  } catch (err) {
    console.error('(GenerateFromMasters) Error:', err.message);
  }
}

// ─── OVERLAPPING TASK CONFLICT RESOLUTION ────────────────────────────────
async function resolveTaskConflicts(runInstanceId) {
  try {
    const snapshot = await db.collection('task_instances')
      .where('runInstanceId', '==', runInstanceId)
      .where('isParentTask', '==', true)
      .where('status', '==', 'PLANNED')
      .get();

    const tasks = {};
    snapshot.forEach(doc => {
      const d = doc.data();
      const key = `${d.coachNo}_${d.conflictGroup || 'default'}`;
      if (!tasks[key]) tasks[key] = [];
      tasks[key].push(d);
    });

    const batch = db.batch();
    let resolved = 0;

    for (const [key, group] of Object.entries(tasks)) {
      if (group.length <= 1) continue;
      // Sort by scheduled time
      group.sort((a, b) => (a.scheduledTime || '00:00').localeCompare(b.scheduledTime || '00:00'));

      for (let i = 1; i < group.length; i++) {
        const prev = group[i - 1];
        const curr = group[i];
        if (prev.scheduledTime === curr.scheduledTime) {
          // Same time + same conflict group = overlap
          // Priority: 1=highest(Safety), 2=Complaint, 3=Policy, 4=Routine
          if (curr.priority > prev.priority) {
            batch.update(db.collection('task_instances').doc(curr.taskId), {
              status: 'DEFERRED',
              exceptionReason: `Deferred due to conflict with ${prev.taskName} (higher priority)`,
              updatedAt: new Date().toISOString()
            });
            resolved++;
          }
        }
      }
    }

    if (resolved > 0) {
      await batch.commit();
      console.log(`[ConflictResolve] Resolved ${resolved} overlapping tasks for ${runInstanceId}`);
    }
  } catch (err) {
    console.error('(ConflictResolve) Error:', err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// == NEW ARCHITECTURE APIs
// ═══════════════════════════════════════════════════════════════════════════

// ─── TASK MASTERS CRUD ───────────────────────────────────────────────────

// List all active task masters (CM, CA, CTS can manage)
app.get('/api/v2/task-masters', verifyToken, async (req, res) => {
  try {
    const { active, assignedRole, category } = req.query;
    let query = db.collection('task_masters');
    if (active === 'true') query = query.where('active', '==', true);
    if (assignedRole) query = query.where('assignedRole', '==', assignedRole);
    if (category) query = query.where('category', '==', category);
    const snap = await query.get();
    const masters = [];
    snap.forEach(doc => masters.push({ id: doc.id, ...doc.data() }));
    res.status(200).json({ success: true, count: masters.length, masters });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch task masters', details: error.message });
  }
});

// Create a custom task master (CM, CA)
app.post('/api/v2/task-masters', verifyToken, async (req, res) => {
  try {
    if (!['CM', 'CA'].includes(req.user.role)) {
      return res.status(403).json({ error: 'Only CM/CA can create task masters' });
    }
    const { taskCode, taskName, category, subCategory, assignedRole, applicableCoachTypes,
      priority, frequencyRules, scheduledTime, checklistTemplate,
      requiresSupervisorVerification, conflictGroup } = req.body;

    if (!taskCode || !taskName || !assignedRole) {
      return res.status(400).json({ error: 'taskCode, taskName, assignedRole are required' });
    }

    const ref = db.collection('task_masters').doc(taskCode);
    const existing = await ref.get();
    if (existing.exists) {
      return res.status(409).json({ error: 'Task master with this code already exists' });
    }

    const data = {
      taskCode, taskName, category: category || 'custom', subCategory: subCategory || null,
      assignedRole, applicableCoachTypes: applicableCoachTypes || ['all'],
      priority: priority || 'medium', requiresSupervisorVerification: requiresSupervisorVerification || false,
      isComplaintDriven: false, conflictGroup: conflictGroup || taskCode,
      frequencyRules: frequencyRules || [], scheduledTime: scheduledTime || null,
      checklistTemplate: checklistTemplate || [],
      active: true, isSystem: false,
      createdBy: req.user.uid, createdByName: req.user.fullName,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
    };
    await ref.set(data);

    await logAudit({
      action: 'TASK_MASTER_CREATED', entityType: 'task_masters',
      entityId: taskCode, actorId: req.user.uid, actorRole: req.user.role,
      actorName: req.user.fullName, newValue: data,
      details: `Created task master: ${taskName} (${taskCode})`
    });

    res.status(201).json({ success: true, message: 'Task master created', taskCode });
  } catch (error) {
    res.status(500).json({ error: 'Failed to create task master', details: error.message });
  }
});

// Update a task master
app.put('/api/v2/task-masters/:taskCode', verifyToken, async (req, res) => {
  try {
    if (!['CM', 'CA'].includes(req.user.role)) {
      return res.status(403).json({ error: 'Only CM/CA can update task masters' });
    }
    const { taskCode } = req.params;
    const ref = db.collection('task_masters').doc(taskCode);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task master not found' });

    const oldData = doc.data();
    const updates = req.body;
    delete updates.taskCode;
    delete updates.isSystem;
    updates.updatedAt = new Date().toISOString();
    updates.updatedBy = req.user.uid;

    await ref.update(updates);

    await logAudit({
      action: 'TASK_MASTER_UPDATED', entityType: 'task_masters',
      entityId: taskCode, actorId: req.user.uid, actorRole: req.user.role,
      actorName: req.user.fullName, oldValue: oldData, newValue: { ...oldData, ...updates },
      details: `Updated task master: ${taskCode}`
    });

    res.status(200).json({ success: true, message: 'Task master updated' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to update task master', details: error.message });
  }
});

// ─── CONTRACTOR WORKER ASSIGNMENTS ────────────────────────────────────────

// Assign workers to coaches for a journey (CA, CTS, CS)
app.post('/api/v2/assignments/create', verifyToken, async (req, res) => {
  try {
    const { runInstanceId, assignments } = req.body;
    // assignments: [{ workerId, workerRole, coachNos[] }]
    if (!runInstanceId || !assignments || !Array.isArray(assignments)) {
      return res.status(400).json({ error: 'runInstanceId and assignments array are required' });
    }

    const runRef = db.collection('RunInstance').doc(runInstanceId);
    const runDoc = await runRef.get();
    if (!runDoc.exists) return res.status(404).json({ error: 'Run instance not found' });
    if (runDoc.data().status !== 'PLANNED') {
      return res.status(400).json({ error: 'Can only assign workers to PLANNED journeys' });
    }

    const batch = db.batch();
    const results = [];

    for (const assign of assignments) {
      const { workerId, workerRole, coachNos } = assign;
      if (!workerId || !workerRole || !coachNos || !Array.isArray(coachNos)) {
        return res.status(400).json({ error: 'Each assignment needs workerId, workerRole, coachNos[]' });
      }

      // Validate role hierarchy
      if (!hasRolePermission(req.user.role, workerRole)) {
        return res.status(403).json({ error: `Cannot assign ${workerRole}. Your role (${req.user.role}) lacks permission.` });
      }

      // Janitor max 2 coaches
      if (workerRole === 'janitor' && coachNos.length > 2) {
        return res.status(400).json({ error: `Janitor ${workerId} cannot be assigned more than 2 coaches (got ${coachNos.length})` });
      }

      // Attendant only for AC coaches
      if (workerRole === 'attendant') {
        // Validate that all assigned coaches are AC type
        const runData = runDoc.data();
        for (const cn of coachNos) {
          const coach = runData.coaches.find(c => (c.coachPosition === cn || c.coachNo === cn));
          if (coach && coach.coachType !== 'AC' && coach.coachType !== 'ac') {
            return res.status(400).json({ error: `Attendant can only be assigned to AC coaches. Coach ${cn} is not AC.` });
          }
        }
      }

      const assignmentId = `${runInstanceId}_${workerId}_${Date.now()}`;
      const ref = db.collection('worker_assignments').doc(assignmentId);
      batch.set(ref, {
        assignmentId, runInstanceId, workerId, workerRole,
        coachNos, assignedBy: req.user.uid,
        assignedByName: req.user.fullName,
        status: 'ACTIVE',
        createdAt: new Date().toISOString()
      });

      // Update RunInstance coaches array
      for (const cn of coachNos) {
        const coachIndex = runDoc.data().coaches.findIndex(c => (c.coachPosition === cn || c.coachNo === cn));
        if (coachIndex !== -1) {
          const coachRef = db.collection('RunInstance').doc(runInstanceId);
          batch.update(coachRef, {
            [`coaches.${coachIndex}.workerId`]: workerId,
            [`coaches.${coachIndex}.workerName`]: req.body.workerName || assign.workerName || 'Unknown',
            [`coaches.${coachIndex}.workerRole`]: workerRole
          });
        }
      }

      results.push({ assignmentId, workerId, workerRole, coachNos });
    }

    await batch.commit();

    // Transition journey to ALLOCATED
    await runRef.update({ status: 'ALLOCATED', updatedAt: new Date().toISOString() });

    await logAudit({
      action: 'WORKERS_ASSIGNED', entityType: 'RunInstance',
      entityId: runInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      newValue: { assignments: results, status: 'ALLOCATED' },
      details: `Assigned ${assignments.length} workers to ${runInstanceId}`
    });

    res.status(201).json({ success: true, message: 'Workers assigned', assignments: results });
  } catch (error) {
    res.status(500).json({ error: 'Failed to assign workers', details: error.message });
  }
});

// Get assignments for a journey
app.get('/api/v2/assignments/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const snap = await db.collection('worker_assignments')
      .where('runInstanceId', '==', runInstanceId)
      .where('status', '==', 'ACTIVE')
      .get();
    const assignments = [];
    snap.forEach(doc => assignments.push(doc.data()));
    res.status(200).json({ success: true, count: assignments.length, assignments });
  } catch (error) {
    res.status(500).json({ error: 'Failed to get assignments', details: error.message });
  }
});

// ─── TASK INSTANCE APIS (Unified) ────────────────────────────────────────

// Get task instances for a journey
app.get('/api/v2/tasks/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const { coachNo, status, workerId } = req.query;
    let query = db.collection('task_instances')
      .where('runInstanceId', '==', runInstanceId);
    if (coachNo) query = query.where('coachNo', '==', coachNo);
    if (status) query = query.where('status', '==', status);
    if (workerId) query = query.where('workerId', '==', workerId);
    const snap = await query.get();
    const tasks = [];
    snap.forEach(doc => tasks.push(doc.data()));
    res.status(200).json({ success: true, count: tasks.length, tasks });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch tasks', details: error.message });
  }
});

// Submit task completion (with full evidence)
app.post('/api/v2/tasks/submit', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, checklistResponses, beforePhoto, afterPhoto,
      gpsLatitude, gpsLongitude, deviceTimestamp, deviceId, mobileNumber, comment } = req.body;

    if (!taskInstanceId || !beforePhoto || !afterPhoto || !gpsLatitude || !gpsLongitude) {
      return res.status(400).json({ error: 'taskInstanceId, beforePhoto, afterPhoto, gpsLatitude, gpsLongitude are mandatory' });
    }

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const taskData = doc.data();
    if (['COMPLETED', 'VERIFIED', 'CLOSED', 'NOT_APPLICABLE'].includes(taskData.status)) {
      return res.status(400).json({ error: `Task already in final state: ${taskData.status}` });
    }

    const updateData = {
      status: 'COMPLETED',
      beforePhoto, afterPhoto,
      gpsLatitude, gpsLongitude,
      deviceId: deviceId || null,
      mobileNumber: mobileNumber || null,
      checklistResponses: checklistResponses || [],
      comment: comment || null,
      completedBy: req.user.uid,
      completedByName: req.user.fullName,
      deviceTimestamp: deviceTimestamp || new Date().toISOString(),
      completedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await ref.update(updateData);

    // If this is a child task, check and update parent
    if (taskData.parentTaskId) {
      await updateParentTaskStatus(taskData.parentTaskId);
    }

    await logAudit({
      action: 'TASK_INSTANCE_COMPLETED', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: taskData.status },
      newValue: { status: 'COMPLETED', gpsLatitude, gpsLongitude },
      details: `Completed ${taskData.taskName} for Coach ${taskData.coachNo}`
    });

    res.status(200).json({ success: true, message: 'Task completed', taskInstanceId });
  } catch (error) {
    res.status(500).json({ error: 'Failed to submit task', details: error.message });
  }
});

// Worker start task (OPEN/OVERDUE → IN_PROGRESS)
app.post('/api/v2/tasks/start', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, gpsLatitude, gpsLongitude, deviceId } = req.body;
    if (!taskInstanceId) return res.status(400).json({ error: 'taskInstanceId required' });

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    if (!['OPEN', 'OVERDUE', 'REOPENED'].includes(oldData.status)) {
      return res.status(400).json({ error: `Cannot start task in status ${oldData.status}` });
    }

    await ref.update({
      status: 'IN_PROGRESS',
      startedAt: new Date().toISOString(),
      startedBy: req.user.uid,
      startedByName: req.user.fullName,
      gpsLatitude: gpsLatitude || null,
      gpsLongitude: gpsLongitude || null,
      deviceId: deviceId || null,
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: 'TASK_STARTED', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: oldData.status },
      newValue: { status: 'IN_PROGRESS' },
      details: `Started ${oldData.taskName} for Coach ${oldData.coachNo}`
    });

    res.status(200).json({ success: true, message: 'Task started', taskInstanceId, status: 'IN_PROGRESS' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to start task', details: error.message });
  }
});

// Supervisor verify a task (COMPLETED → VERIFIED or REJECTED)
app.post('/api/v2/tasks/verify', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, verified, remarks } = req.body;
    if (!taskInstanceId) return res.status(400).json({ error: 'taskInstanceId required' });

    const allowedRoles = ['CS', 'CTS', 'CA', 'CM'];
    if (!allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Only CS/CTS/CA/CM can verify tasks' });
    }

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    const isApproved = verified !== false;
    const newStatus = isApproved ? 'VERIFIED' : 'REJECTED';

    await ref.update({
      supervisorVerified: isApproved,
      supervisorId: req.user.uid,
      supervisorName: req.user.fullName,
      supervisorRemarks: remarks || null,
      verifiedAt: new Date().toISOString(),
      status: newStatus,
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: isApproved ? 'TASK_VERIFIED' : 'TASK_REJECTED',
      entityType: 'task_instances', entityId: taskInstanceId,
      actorId: req.user.uid, actorRole: req.user.role,
      actorName: req.user.fullName,
      oldValue: { status: oldData.status, supervisorVerified: oldData.supervisorVerified },
      newValue: { status: newStatus, supervisorVerified: isApproved, remarks },
      details: remarks || 'No remarks'
    });

    res.status(200).json({ success: true, message: isApproved ? 'Task verified' : 'Task rejected', status: newStatus });
  } catch (error) {
    res.status(500).json({ error: 'Failed to verify task', details: error.message });
  }
});

// Close task (VERIFIED → CLOSED)
app.post('/api/v2/tasks/close', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, remarks } = req.body;
    if (!taskInstanceId) return res.status(400).json({ error: 'taskInstanceId required' });

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    if (oldData.status !== 'VERIFIED') {
      return res.status(400).json({ error: 'Only VERIFIED tasks can be closed' });
    }

    await ref.update({
      status: 'CLOSED',
      closedAt: new Date().toISOString(),
      closedBy: req.user.uid,
      closedByName: req.user.fullName,
      closeRemarks: remarks || null,
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: 'TASK_CLOSED', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: oldData.status },
      newValue: { status: 'CLOSED' },
      details: remarks || 'Task closed'
    });

    res.status(200).json({ success: true, message: 'Task closed', taskInstanceId, status: 'CLOSED' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to close task', details: error.message });
  }
});

// Reopen task (any completed state → REOPENED)
app.post('/api/v2/tasks/reopen', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, reason } = req.body;
    if (!taskInstanceId || !reason) return res.status(400).json({ error: 'taskInstanceId and reason required' });

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    const completableStatuses = ['COMPLETED', 'VERIFIED', 'CLOSED', 'REJECTED'];
    if (!completableStatuses.includes(oldData.status)) {
      return res.status(400).json({ error: `Cannot reopen task in status ${oldData.status}` });
    }

    await ref.update({
      status: 'REOPENED',
      reopenedAt: new Date().toISOString(),
      reopenedBy: req.user.uid,
      reopenedByName: req.user.fullName,
      reopenReason: reason,
      supervisorVerified: false,
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: 'TASK_REOPENED', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: oldData.status },
      newValue: { status: 'REOPENED', reason },
      details: `Reopened: ${reason}`
    });

    res.status(200).json({ success: true, message: 'Task reopened', taskInstanceId, status: 'REOPENED' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to reopen task', details: error.message });
  }
});

// Mark task as not applicable
app.post('/api/v2/tasks/not-applicable', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId, reason } = req.body;
    if (!taskInstanceId || !reason) return res.status(400).json({ error: 'taskInstanceId and reason required' });

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    if (!['OPEN', 'PLANNED', 'OVERDUE'].includes(oldData.status)) {
      return res.status(400).json({ error: `Cannot mark task as N/A in status ${oldData.status}` });
    }

    await ref.update({
      status: 'NOT_APPLICABLE',
      notApplicableReason: reason,
      notApplicableMarkedBy: req.user.uid,
      notApplicableMarkedByName: req.user.fullName,
      notApplicableAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: 'TASK_NA', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: oldData.status },
      newValue: { status: 'NOT_APPLICABLE', reason },
      details: `Marked N/A: ${reason}`
    });

    res.status(200).json({ success: true, message: 'Task marked not applicable', taskInstanceId, status: 'NOT_APPLICABLE' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to mark task N/A', details: error.message });
  }
});

// Supervisor edit task instance (reassign worker, reschedule, update fields)
app.put('/api/v2/tasks/:taskInstanceId', verifyToken, async (req, res) => {
  try {
    const { taskInstanceId } = req.params;
    const allowedRoles = ['CS', 'CTS', 'CA', 'CM'];
    if (!allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Only CS/CTS/CA/CM can edit tasks' });
    }

    const ref = db.collection('task_instances').doc(taskInstanceId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Task instance not found' });

    const oldData = doc.data();
    const updates = { ...req.body };

    // Remove protected fields that cannot be edited by supervisor
    delete updates.taskId;
    delete updates.runInstanceId;
    delete updates.isParentTask;
    delete updates.parentTaskId;
    delete updates.childTaskIds;
    delete updates.createdAt;

    // If reassigning worker, resolve worker name
    if (updates.workerId && updates.workerId !== oldData.workerId) {
      if (!updates.workerName) {
        try {
          const workerDoc = await db.collection('users').doc(updates.workerId).get();
          if (workerDoc.exists) {
            updates.workerName = workerDoc.data().fullName || workerDoc.data().name || 'Unknown';
          }
        } catch (_) { updates.workerName = 'Unknown'; }
      }
    }

    updates.updatedAt = new Date().toISOString();
    updates.editedBy = req.user.uid;
    updates.editedByName = req.user.fullName;
    updates.editedAt = updates.updatedAt;

    await ref.update(updates);

    // Also update child tasks if worker reassigned on a parent task
    if (oldData.isParentTask && (updates.workerId || updates.workerName || updates.workerRole)) {
      const children = oldData.childTaskIds || [];
      const childBatch = db.batch();
      for (const cId of children) {
        const childRef = db.collection('task_instances').doc(cId);
        childBatch.update(childRef, {
          ...(updates.workerId && { workerId: updates.workerId }),
          ...(updates.workerName && { workerName: updates.workerName }),
          ...(updates.workerRole && { workerRole: updates.workerRole }),
          ...(updates.scheduledTime && { scheduledTime: updates.scheduledTime }),
          ...(updates.coachNo && { coachNo: updates.coachNo }),
          updatedAt: new Date().toISOString()
        });
      }
      await childBatch.commit();
    }

    await logAudit({
      action: 'TASK_EDITED', entityType: 'task_instances',
      entityId: taskInstanceId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: oldData, newValue: { ...oldData, ...updates },
      details: `Supervisor ${req.user.fullName} edited task ${taskInstanceId}`
    });

    res.status(200).json({ success: true, message: 'Task updated', taskInstanceId });
  } catch (error) {
    res.status(500).json({ error: 'Failed to edit task', details: error.message });
  }
});

// Helper: Update parent task status when child completes
async function updateParentTaskStatus(parentTaskId) {
  try {
    const parentRef = db.collection('task_instances').doc(parentTaskId);
    const parentDoc = await parentRef.get();
    if (!parentDoc.exists) return;

    const parentData = parentDoc.data();
    const childIds = parentData.childTaskIds || [];

    let allCompleted = true;
    let anyOverdue = false;
    let anyOpen = false;
    let anyRejected = false;
    let anyNA = false;
    let anyReopened = false;

    for (const cId of childIds) {
      const cDoc = await db.collection('task_instances').doc(cId).get();
      if (!cDoc.exists) { allCompleted = false; continue; }
      const s = cDoc.data().status;
      const finalStates = ['COMPLETED', 'VERIFIED', 'CLOSED', 'NOT_APPLICABLE'];
      if (!finalStates.includes(s)) allCompleted = false;
      if (s === 'OVERDUE') anyOverdue = true;
      if (s === 'OPEN' || s === 'PLANNED') anyOpen = true;
      if (s === 'REJECTED') anyRejected = true;
      if (s === 'NOT_APPLICABLE') anyNA = true;
      if (s === 'REOPENED') anyReopened = true;
    }

    let newStatus;
    if (allCompleted) newStatus = 'COMPLETED';
    else if (anyReopened) newStatus = 'REOPENED';
    else if (anyRejected) newStatus = 'PARTIAL';
    else if (anyOverdue) newStatus = 'OVERDUE';
    else if (anyOpen) newStatus = 'IN_PROGRESS';
    else newStatus = parentData.status;

    if (newStatus !== parentData.status) {
      await parentRef.update({ status: newStatus, updatedAt: new Date().toISOString() });
    }
  } catch (err) {
    console.error('(UpdateParentTask) Error:', err.message);
  }
}

// ─── ESCALATION APIS ──────────────────────────────────────────────────────

// Create escalation
app.post('/api/v2/escalations/create', verifyToken, async (req, res) => {
  try {
    const { sourceEntity, sourceId, reason, escalatedToRole, details } = req.body;
    if (!sourceEntity || !sourceId || !reason) {
      return res.status(400).json({ error: 'sourceEntity, sourceId, reason required' });
    }

    const ref = db.collection('escalations').doc();
    const escalation = {
      escalationId: ref.id,
      sourceEntity, sourceId,
      reason, details: details || '',
      escalatedBy: req.user.uid, escalatedByName: req.user.fullName,
      escalatedByRole: req.user.role,
      escalatedToRole: escalatedToRole || getNextEscalationRole(req.user.role),
      status: 'OPEN',
      resolvedAt: null, resolvedBy: null, resolution: null,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
    };
    await ref.set(escalation);

    // Update source entity status (non-blocking)
    if (sourceEntity === 'task_instance') {
      try {
        const srcDoc = await db.collection('task_instances').doc(sourceId).get();
        if (srcDoc.exists) {
          await srcDoc.ref.update({
            status: 'ESCALATED',
            escalationId: ref.id,
            updatedAt: new Date().toISOString()
          });
        }
      } catch (srcErr) {
        console.error('(Escalation) Source update warning:', srcErr.message);
      }
    }

    await logAudit({
      action: 'ESCALATION_CREATED', entityType: 'escalations',
      entityId: ref.id, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      newValue: escalation, details: `Escalated ${sourceEntity}:${sourceId} - ${reason}`
    });

    res.status(201).json({ success: true, message: 'Escalation created', escalationId: ref.id });
  } catch (error) {
    res.status(500).json({ error: 'Failed to create escalation', details: error.message });
  }
});

// Resolve escalation
app.post('/api/v2/escalations/resolve/:escalationId', verifyToken, async (req, res) => {
  try {
    const { escalationId } = req.params;
    const { resolution } = req.body;

    const ref = db.collection('escalations').doc(escalationId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Escalation not found' });

    const oldData = doc.data();
    await ref.update({
      status: 'RESOLVED',
      resolvedBy: req.user.uid,
      resolvedByName: req.user.fullName,
      resolution: resolution || 'Resolved',
      resolvedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    await logAudit({
      action: 'ESCALATION_RESOLVED', entityType: 'escalations',
      entityId: escalationId, actorId: req.user.uid,
      actorRole: req.user.role, actorName: req.user.fullName,
      oldValue: { status: oldData.status },
      newValue: { status: 'RESOLVED', resolution },
      details: resolution || 'Resolved'
    });

    res.status(200).json({ success: true, message: 'Escalation resolved' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to resolve escalation', details: error.message });
  }
});

// Get escalation chain
function getNextEscalationRole(currentRole) {
  const chain = ['janitor', 'CS', 'CTS', 'CA', 'CM'];
  const idx = chain.indexOf(currentRole);
  if (idx === -1 || idx >= chain.length - 1) return 'CM';
  return chain[idx + 1];
}

// Get escalations
app.get('/api/v2/escalations', verifyToken, async (req, res) => {
  try {
    const { status, sourceEntity } = req.query;
    let query = db.collection('escalations');
    if (status) query = query.where('status', '==', status);
    if (sourceEntity) query = query.where('sourceEntity', '==', sourceEntity);
    const snap = await query.orderBy('createdAt', 'desc').get();
    const list = [];
    snap.forEach(doc => list.push(doc.data()));
    res.status(200).json({ success: true, count: list.length, escalations: list });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch escalations', details: error.message });
  }
});

// ─── AUDIT LOG APIS ───────────────────────────────────────────────────────

// Get audit logs
app.get('/api/v2/audit-logs', verifyToken, async (req, res) => {
  try {
    const { entityType, entityId, action, limit: qLimit } = req.query;
    let query = db.collection('audit_logs');
    if (entityType) query = query.where('entityType', '==', entityType);
    if (entityId) query = query.where('entityId', '==', entityId);
    if (action) query = query.where('action', '==', action);
    const snap = await query.orderBy('timestamp', 'desc').limit(parseInt(qLimit) || 100).get();
    const logs = [];
    snap.forEach(doc => logs.push(doc.data()));
    res.status(200).json({ success: true, count: logs.length, logs });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch audit logs', details: error.message });
  }
});

// ─── JOURNEY CLOSURE TASK APIS ────────────────────────────────────────────

// Get closure tasks
app.get('/api/v2/closure-tasks/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const snap = await db.collection('closure_tasks')
      .where('runInstanceId', '==', runInstanceId)
      .get();
    const tasks = [];
    snap.forEach(doc => tasks.push(doc.data()));
    res.status(200).json({ success: true, count: tasks.length, tasks });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch closure tasks', details: error.message });
  }
});

// Complete a closure task
app.post('/api/v2/closure-tasks/complete', verifyToken, async (req, res) => {
  try {
    const { taskId, beforePhoto, afterPhoto, gpsLatitude, gpsLongitude } = req.body;
    if (!taskId) return res.status(400).json({ error: 'taskId required' });

    const ref = db.collection('closure_tasks').doc(taskId);
    const doc = await ref.get();
    if (!doc.exists) return res.status(404).json({ error: 'Closure task not found' });

    await ref.update({
      status: 'COMPLETED',
      beforePhoto: beforePhoto || null,
      afterPhoto: afterPhoto || null,
      gpsLatitude: gpsLatitude || null,
      gpsLongitude: gpsLongitude || null,
      completedAt: new Date().toISOString()
    });

    res.status(200).json({ success: true, message: 'Closure task completed', taskId });
  } catch (error) {
    res.status(500).json({ error: 'Failed to complete closure task', details: error.message });
  }
});

// ─── ROLE-BASED WORKER PROFILE ────────────────────────────────────────────

// Get worker's assigned tasks for their current journey
app.get('/api/v2/worker/my-tasks', verifyToken, async (req, res) => {
  try {
    const workerId = req.user.uid;
    const { status: filterStatus } = req.query;

    // Find active journey for this worker (as janitor or attendant)
    const runsSnap = await db.collection('RunInstance')
      .where('status', '==', 'Active')
      .get();

    let myRun = null;
    let myCoaches = [];
    let myCoachTasks = null; // per-coach task filter array
    runsSnap.forEach(doc => {
      const d = doc.data();
      const assigned = (d.coaches || []).filter(c => c.workerId === workerId || c.attendantId === workerId);
      if (assigned.length > 0) {
        myRun = { ...d, runInstanceId: doc.id };
        myCoaches = assigned.map(c => c.coachPosition || c.coachNo);
        // Collect per-coach task filters
        myCoachTasks = assigned.reduce((acc, c) => {
          const cn = c.coachPosition || c.coachNo;
          if (c.tasks && c.tasks.length > 0) acc[cn] = c.tasks.map(t => t.toLowerCase());
          return acc;
        }, {});
      }
    });

    if (!myRun) {
      return res.status(404).json({ success: false, error: 'No active journey found for this worker' });
    }

    let taskQuery = db.collection('task_instances')
      .where('runInstanceId', '==', myRun.runInstanceId)
      .where('isParentTask', '==', false);

    // If worker is the janitor (primary workerId)
    const janitorCoaches = myRun.coaches.filter(c => c.workerId === workerId).map(c => c.coachPosition || c.coachNo);
    // If worker is the attendant (secondary)
    const attendantCoaches = myRun.coaches.filter(c => c.attendantId === workerId).map(c => c.coachPosition || c.coachNo);
    const assignedCoachNos = [...new Set([...janitorCoaches, ...attendantCoaches])];

    const taskSnap = await taskQuery.get();
    let tasks = [];
    taskSnap.forEach(doc => tasks.push(doc.data()));

    // Filter by assigned coaches
    tasks = tasks.filter(t => assignedCoachNos.includes(t.coachNo));

    // Apply per-coach task filtering if defined
    if (myCoachTasks) {
      tasks = tasks.filter(t => {
        const allowed = myCoachTasks[t.coachNo];
        if (!allowed) return true; // no filter for this coach
        const taskName = (t.taskName || t.taskMasterCode || '').toLowerCase();
        return allowed.some(keyword => taskName.includes(keyword));
      });
    }

    if (filterStatus) tasks = tasks.filter(t => t.status === filterStatus);

    res.status(200).json({
      success: true,
      journey: {
        runInstanceId: myRun.runInstanceId,
        trainNo: myRun.trainNo, trainName: myRun.trainName,
        coaches: myCoaches,
        actualDeparture: myRun.actualDeparture,
        scheduledArrival: myRun.scheduledArrival
      },
      taskCount: tasks.length,
      tasks
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch worker tasks', details: error.message });
  }
});

// ─── JOURNEY TIMELINE (Delay-aware) ──────────────────────────────────────

// Get journey timeline with delay-adjusted task schedule
app.get('/api/v2/journey/timeline/:runInstanceId', verifyToken, async (req, res) => {
  try {
    const { runInstanceId } = req.params;
    const runDoc = await db.collection('RunInstance').doc(runInstanceId).get();
    if (!runDoc.exists) return res.status(404).json({ error: 'Journey not found' });

    const run = runDoc.data();
    const actualDeparture = run.actualDeparture ? new Date(run.actualDeparture) : null;
    const scheduledDeparture = run.scheduledDeparture ? new Date(run.scheduledDeparture) : null;
    const delayMs = actualDeparture && scheduledDeparture
      ? actualDeparture.getTime() - scheduledDeparture.getTime()
      : 0;
    const delayMinutes = Math.round(delayMs / 60000);

    // Fetch tasks
    const taskSnap = await db.collection('task_instances')
      .where('runInstanceId', '==', runInstanceId)
      .where('isParentTask', '==', true)
      .get();
    const tasks = [];
    taskSnap.forEach(doc => tasks.push(doc.data()));

    // Adjust task scheduled times for delay
    const adjustedTasks = tasks.map(t => {
      if (delayMinutes > 0 && t.scheduledTime) {
        const [h, m] = t.scheduledTime.split(':').map(Number);
        const totalMin = h * 60 + m + delayMinutes;
        const newH = Math.floor(totalMin / 60) % 24;
        const newM = totalMin % 60;
        return { ...t, originalScheduledTime: t.scheduledTime, adjustedTime: `${String(newH).padStart(2, '0')}:${String(newM).padStart(2, '0')}` };
      }
      return { ...t, adjustedTime: t.scheduledTime };
    });

    res.status(200).json({
      success: true,
      journey: {
        runInstanceId, trainNo: run.trainNo, trainName: run.trainName,
        scheduledDeparture: run.scheduledDeparture,
        actualDeparture: run.actualDeparture,
        scheduledArrival: run.scheduledArrival,
        actualArrival: run.actualArrival,
        delayMinutes,
        status: run.status
      },
      taskCount: adjustedTasks.length,
      tasks: adjustedTasks
    });
  } catch (error) {
    res.status(500).json({ error: 'Failed to get journey timeline', details: error.message });
  }
});

// =======================================================
// == WORKER ASSIGNMENTS (MCC Flow)
// =======================================================

// GET /api/obhs/worker/active-run — Get worker's active run with assigned coaches
app.get('/api/obhs/worker/active-run', verifyToken, async (req, res) => {
  try {
    const uid = req.user.uid;
    const allRuns = await db.collection('RunInstance')
      .where('status', 'in', ['ALLOCATED', 'ACTIVE', 'Active', 'active', 'READY', 'ready', 'Running', 'running'])
      .orderBy('createdAt', 'desc')
      .get();

    let runDoc = null;
    allRuns.forEach(doc => {
      const coaches = doc.data().coaches || [];
      if (coaches.some(c => c.workerId === uid)) {
        runDoc = doc;
      }
    });

    if (!runDoc) {
      return res.status(200).json({ success: true, hasAssignment: false, run: null });
    }
    const runData = { id: runDoc.id, ...runDoc.data() };
    const myCoaches = (runData.coaches || []).filter(c => c.workerId === uid);

    res.status(200).json({
      success: true,
      hasAssignment: true,
      run: {
        runInstanceId: runData.id,
        trainNo: runData.trainNo || runData.trainNumber || 'N/A',
        trainName: runData.trainName || runData.name || '',
        status: runData.status,
        departureDate: runData.departureDate || '',
        departureTime: runData.departureTime || '',
      },
      coaches: myCoaches.map(c => ({
        coachNo: c.coachPosition || c.coachNo || c.id || 'N/A',
        coachType: c.coachType || 'general',
        workerId: c.workerId,
        workerName: c.workerName || c.name || '',
        workerRole: c.workerRole || 'janitor'
      }))
    });
  } catch (error) {
    console.error('(WorkerActiveRun) Error:', error);
    res.status(500).json({ error: 'Failed to fetch worker assignment', details: error.message });
  }
});

// ─── SEED ON STARTUP ──────────────────────────────────────────────────────
seedTaskMasters();

// --- Server Start ---
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});

