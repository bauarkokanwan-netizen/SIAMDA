function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SIAMDA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let users = ss.getSheetByName('Users');
  if (!users) {
    users = ss.insertSheet('Users');
    users.appendRow(['id','name','email','password','role','status']);
    users.appendRow(['ADM001','Administrator','admin@example.com','admin123','admin','active']);
    users.appendRow(['USR001','Demo User','user@example.com','user123','user','active']);
  }

  let attendance = ss.getSheetByName('Attendance');
  if (!attendance) {
    attendance = ss.insertSheet('Attendance');
    attendance.appendRow([
      'timestamp','user_id','name','type','latitude','longitude','status'
    ]);
  }
}

function login(email, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == email && data[i][3] == password) {
      return {
        success: true,
        user: {
          id: data[i][0],
          name: data[i][1],
          role: data[i][4]
        }
      };
    }
  }

  return { success: false };
}

function submitAttendance(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance');

  sheet.appendRow([
    new Date(),
    payload.userId,
    payload.name,
    payload.type,
    payload.latitude,
    payload.longitude,
    'SUCCESS'
  ]);

  return {
    success: true,
    message: 'Attendance saved'
  };
}
