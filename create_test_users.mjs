import admin from 'firebase-admin';
import { readFileSync } from 'fs';

const serviceAccount = process.env.FIREBASE_SERVICE_ACCOUNT 
  ? JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT) 
  : JSON.parse(readFileSync('./serviceAccountKey.json', 'utf8'));

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

await db.collection('users').doc('test-cm-uid').set({
  uid: 'test-cm-uid', fullName: 'Test CM', role: 'CM',
  userType: 'contractor', email: 'test@test.com',
  mobile: '9999999999', zone: 'NR', division: 'DELHI',
  status: 'APPROVED', createdAt: new Date().toISOString()
});

await db.collection('users').doc('test-jan-uid').set({
  uid: 'test-jan-uid', fullName: 'Test Janitor', role: 'janitor',
  userType: 'contractor', email: 'jan@test.com',
  mobile: '9999999998', status: 'APPROVED',
  createdAt: new Date().toISOString()
});

console.log('Users created');
admin.app().delete();
