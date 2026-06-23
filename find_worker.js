import { initializeApp, cert } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import { readFile } from 'fs/promises';

async function findWorker() {
  const serviceAccount = JSON.parse(await readFile('./serviceAccountKey.json', 'utf-8'));
  initializeApp({ credential: cert(serviceAccount) });
  const db = getFirestore();

  const snapshot = await db.collection('users').where('role', '==', 'Worker').limit(1).get();
  if (snapshot.empty) {
    console.log("No workers found.");
  } else {
    console.log("Worker found:", snapshot.docs[0].id, snapshot.docs[0].data());
  }
}
findWorker();
