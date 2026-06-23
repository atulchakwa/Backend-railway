import admin from 'firebase-admin';
import { readFile } from 'fs/promises';

async function seedData() {
  try {
    const serviceAccount = JSON.parse(await readFile('./serviceAccountKey.json', 'utf-8'));
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount)
    });
    const db = admin.firestore();

    console.log("Seeding Database...");

    // 1. Get a dummy worker (or create one)
    let workerId = "demo_worker_123";
    let workerName = "Demo Janitor";
    const workersSnapshot = await db.collection('users').where('role', '==', 'Worker').limit(1).get();
    if (!workersSnapshot.empty) {
      workerId = workersSnapshot.docs[0].id;
      workerName = workersSnapshot.docs[0].data().fullName || "Worker";
      console.log(`Using existing worker: ${workerName} (${workerId})`);
    } else {
      console.log("No existing worker found. Creating a demo worker.");
      await db.collection('users').doc(workerId).set({
        uid: workerId,
        fullName: workerName,
        email: "worker@demo.com",
        role: "Worker",
        designation: "Janitor",
        status: "APPROVED"
      });
    }

    // 2. Create a Train
    const trainRef = db.collection('trains').doc();
    const trainId = trainRef.id;
    await trainRef.set({
      trainNo: "12345",
      trainName: "Swachh Express",
      status: "active",
      TrainApplicableFor: "OBHS",
      coaches: ["S1", "S2", "B1"]
    });
    console.log(`Created Train: Swachh Express (${trainId})`);

    // 3. Create a Run Instance
    const instanceRef = db.collection('RunInstance').doc();
    const runInstanceId = instanceRef.id;
    const now = new Date();
    await instanceRef.set({
      instanceId: `OBHS-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-12345`,
      trainId: trainId,
      trainNo: "12345",
      trainName: "Swachh Express",
      status: "ACTIVE",
      createdAt: now.toISOString(),
      coaches: [
        { coachPosition: "S1", coachType: "Sleeper", workerId: workerId, workerName: workerName },
        { coachPosition: "S2", coachType: "Sleeper", workerId: null, workerName: null },
        { coachPosition: "B1", coachType: "AC 3-Tier", workerId: null, workerName: null }
      ]
    });
    console.log(`Created Run Instance: ${runInstanceId}`);

    // 4. Create 3 Task Instances for the assigned worker on Coach S1
    const tasks = [
      { taskName: "Toilet Cleaning", status: "PENDING" },
      { taskName: "Mopping Floor", status: "SUBMITTED", beforePhoto: "https://via.placeholder.com/150", afterPhoto: "https://via.placeholder.com/150" },
      { taskName: "Garbage Collection", status: "APPROVED", passengerScore: 8, supervisorScore: 9, consolidatedScore: 8.3 }
    ];

    for (const t of tasks) {
      const taskRef = db.collection('task_instances').doc();
      const taskData = {
        taskId: taskRef.id,
        runInstanceId: runInstanceId,
        trainId: trainId,
        coachId: "S1",
        workerId: workerId,
        workerName: workerName,
        taskName: t.taskName,
        status: t.status,
        createdAt: admin.firestore.FieldValue.serverTimestamp()
      };
      
      if (t.status === "SUBMITTED" || t.status === "APPROVED") {
        taskData.beforePhotoUrl = t.beforePhoto || "https://via.placeholder.com/150";
        taskData.afterPhotoUrl = t.afterPhoto || "https://via.placeholder.com/150";
        taskData.submittedAt = now.toISOString();
      }
      
      if (t.status === "APPROVED") {
        taskData.passengerScore = t.passengerScore;
        taskData.supervisorScore = t.supervisorScore;
        taskData.consolidatedScore = t.consolidatedScore;
      }

      await taskRef.set(taskData);
      console.log(`Created Task: ${t.taskName} [${t.status}]`);
    }

    console.log("Seeding completed successfully!");
    process.exit(0);
  } catch (error) {
    console.error("Seeding failed:", error);
    process.exit(1);
  }
}

seedData();
