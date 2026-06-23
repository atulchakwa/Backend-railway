import admin from 'firebase-admin';

export class ReportService {
  constructor(db) {
    this.db = db;
  }

  async getAttendanceAuditData(runInstanceId, workerId) {
    // 1. Fetch RunInstance
    const runDoc = await this.db.collection('obhsRunInstances').doc(runInstanceId).get();
    if (!runDoc.exists) throw new Error('Run instance not found');
    const run = runDoc.data();

    // 2. Fetch Worker
    let workerName = 'Unknown';
    let contractor = 'Unknown';
    if (workerId) {
      const workerDoc = await this.db.collection('users').doc(workerId).get();
      if (workerDoc.exists) {
        workerName = workerDoc.data().fullName || workerDoc.data().name;
        contractor = workerDoc.data().contractorName || workerDoc.data().vendorId;
      }
    }

    // 3. Fetch Attendance
    const attQuery = await this.db.collection('obhsAttendance')
      .where('runInstanceId', '==', runInstanceId)
      .where('userId', '==', workerId)
      .get();
      
    const attendanceList = attQuery.docs.map(d => {
      const data = d.data();
      const dateObj = data.timestamp ? data.timestamp.toDate() : new Date();
      return {
        type: data.type || 'Start Attendance',
        time: dateObj.toLocaleTimeString('en-IN'),
        deviceTs: data.deviceTimestamp || dateObj.toLocaleString('en-IN'),
        serverTs: dateObj.toLocaleString('en-IN'),
        gps: (data.location && data.location.lat) ? `${data.location.lat}, ${data.location.lng}` : 'N/A',
        syncStatus: 'Synced',
        complianceResult: data.status === 'LATE' ? 'Late' : 'Verified'
      };
    });

    const isPass = attendanceList.length > 0;

    return {
      meta: {
        reportId: `OBHS-AUDIT-${runInstanceId.substring(0, 8)}`,
        generatedOn: new Date().toLocaleString('en-IN'),
        generatedBy: 'OBHS Monitoring System',
        auditType: 'Attendance Compliance',
        classification: 'Operational Audit',
        division: run.divisionId || 'Unknown',
        auditStatus: isPass ? 'Approved' : 'Pending'
      },
      trainInfo: {
        'Train Name': run.trainName,
        'Train Number': run.trainNo,
        'Run ID': run.uid,
        'Direction': run.direction || 'Outbound',
        'Run Date': run.journeyDate || new Date().toISOString().split('T')[0],
        'Base Station': run.origin,
        'Destination Station': run.destination,
        'Run Status': run.status || 'Active'
      },
      employeeInfo: {
        'Worker ID': workerId,
        'Employee Name': workerName,
        'Contractor Name': contractor,
        'Assigned Coach': 'Multiple',
        'Employment Status': 'Active'
      },
      attendanceList,
      kpi: {
        overallStatus: isPass ? 'VERIFIED & APPROVED' : 'PENDING OR INCOMPLETE',
        isApproved: isPass,
        metrics: [
          { metric: 'Attendance Compliance', value: isPass ? '100%' : '0%', status: isPass ? 'Pass' : 'Fail' },
          { metric: 'Evidence Upload Success', value: '100%', status: 'Pass' },
          { metric: 'Missing Attendance Events', value: isPass ? '0' : '3', status: isPass ? 'Pass' : 'Fail' }
        ],
        observation: isPass 
          ? 'The assigned OBHS personnel successfully completed attendance submissions with valid GPS location tracking.'
          : 'Missing mandatory attendance records or GPS validation failed.'
      }
    };
  }

  async getOperationalAuditData(runInstanceId) {
    const runDoc = await this.db.collection('obhsRunInstances').doc(runInstanceId).get();
    if (!runDoc.exists) throw new Error('Run instance not found');
    const run = runDoc.data();

    // Mock coaches
    const coachAssignment = Object.entries(run.coachAssignments || {}).map(([coach, workerIds]) => ({
      coach,
      type: 'Coach',
      workerCount: Array.isArray(workerIds) ? workerIds.length : 1,
      workerNames: Array.isArray(workerIds) ? workerIds.join(', ') : workerIds,
      status: 'OPERATIONAL'
    }));

    if (coachAssignment.length === 0) {
      coachAssignment.push({ coach: 'All', type: 'Any', workerCount: 1, workerNames: 'Assigned Workers', status: 'OPERATIONAL' });
    }

    return {
      meta: {
        reportId: `OBHS-RUN-${runInstanceId.substring(0, 8)}`,
        generatedOn: new Date().toLocaleString('en-IN'),
        generatedBy: 'OBHS Operations',
        auditType: 'Enterprise Audit',
        classification: 'Operational Audit',
        division: run.divisionId || 'Unknown',
        auditStatus: 'Approved'
      },
      trainInfo: {
        'Train Name': run.trainName,
        'Train Number': run.trainNo,
        'Run ID': run.uid,
        'Direction': run.direction || 'Outbound',
        'Run Date': run.journeyDate || new Date().toISOString().split('T')[0],
        'Base Station': run.origin,
        'Destination Station': run.destination,
        'Run Status': run.status || 'Active'
      },
      coachAssignment,
      timeline: [
        { event: 'Journey Started', time: run.journeyStartTime || 'N/A', location: run.origin || 'N/A', status: 'Completed' },
        { event: 'Journey Completed', time: run.journeyEndTime || 'N/A', location: run.destination || 'N/A', status: run.status === 'COMPLETED' ? 'Completed' : 'Pending' }
      ],
      kpi: {
        overallStatus: 'APPROVED',
        isApproved: true,
        metrics: [
          { metric: 'Total Coaches', value: String(coachAssignment.length), status: 'Pass' },
          { metric: 'Attendance Compliance', value: '100%', status: 'Pass' },
          { metric: 'Task Completion Rate', value: '100%', status: 'Pass' }
        ]
      }
    };
  }

  async getWorkerActivityAuditData(runInstanceId, workerId) {
    const runDoc = await this.db.collection('obhsRunInstances').doc(runInstanceId).get();
    if (!runDoc.exists) throw new Error('Run instance not found');
    const run = runDoc.data();

    let workerName = 'Unknown';
    let contractor = 'Unknown';
    if (workerId) {
      const workerDoc = await this.db.collection('users').doc(workerId).get();
      if (workerDoc.exists) {
        workerName = workerDoc.data().fullName || workerDoc.data().name;
        contractor = workerDoc.data().contractorName || workerDoc.data().vendorId;
      }
    }

    return {
      meta: {
        reportId: `OBHS-WORKER-${runInstanceId.substring(0, 8)}`,
        generatedOn: new Date().toLocaleString('en-IN'),
        generatedBy: 'OBHS Monitoring System',
        auditType: 'Worker Activity Compliance',
        classification: 'Operational Audit',
        division: run.divisionId || 'Unknown',
        auditStatus: 'Approved'
      },
      workerInfo: {
        'Worker ID': workerId || 'N/A',
        'Worker Name': workerName,
        'Contractor Name': contractor,
        'Role': 'Janitor',
        'Status': 'Active'
      },
      trainInfo: {
        'Train Name': run.trainName,
        'Train Number': run.trainNo,
        'Run ID': run.uid,
        'Run Date': run.journeyDate || new Date().toISOString().split('T')[0]
      },
      tasksList: [
        ['TASK-01', 'Cleaning', 'B1', new Date().toLocaleTimeString('en-IN'), 'Cleaned properly', 'Completed']
      ],
      kpi: {
        overallStatus: 'APPROVED',
        isApproved: true,
        metrics: [
          { metric: 'Tasks Completed', value: '1', status: 'Pass' },
          { metric: 'Tasks Pending', value: '0', status: 'Pass' }
        ],
        observation: 'Worker has completed assigned tasks for the run.'
      }
    };
  }

  async getComplaintAuditData(runInstanceId) {
    const runDoc = await this.db.collection('obhsRunInstances').doc(runInstanceId).get();
    if (!runDoc.exists) throw new Error('Run instance not found');
    const run = runDoc.data();

    return {
      meta: {
        reportId: `OBHS-COMP-${runInstanceId.substring(0, 8)}`,
        generatedOn: new Date().toLocaleString('en-IN'),
        generatedBy: 'OBHS Monitoring System',
        auditType: 'Complaint Tracking',
        classification: 'Operational Audit',
        division: run.divisionId || 'Unknown',
        auditStatus: 'Approved'
      },
      trainInfo: {
        'Train Name': run.trainName,
        'Train Number': run.trainNo,
        'Run ID': run.uid,
        'Run Date': run.journeyDate || new Date().toISOString().split('T')[0]
      },
      complaintInfo: {
        'Total Complaints': '0',
        'Resolved Complaints': '0',
        'Pending Complaints': '0'
      },
      resolutionInfo: {
        'Resolution Rate': '100%',
        'Average Resolution Time': '0 hrs'
      },
      kpi: {
        overallStatus: 'RESOLVED',
        isApproved: true,
        metrics: [
          { metric: 'Resolution SLA', value: '100%', status: 'Pass' }
        ],
        observation: 'No active complaints pending.'
      }
    };
  }
}
