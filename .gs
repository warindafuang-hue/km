// --- CONFIGURATION ---
const SHEET_ID = '1gBTpIqB7-4Ne2suMLaiEXUc04KPlD-QslELo-7y1CvQ'; 
const SHEET_NAME_TASK = 'Task_List';
const SHEET_NAME_USERS = 'Users';
const SHEET_NAME_LOG = 'Log';
const SHEET_NAME_MEETINGS = 'Meetings'; 

const LINE_CHANNEL_TOKEN = '/4iNwZVyvs7Cdu3FdGeItOvYHaVlyKMbtg4g7dRevfH+q67NZyZleMXUkVpTBACUtGO9aKsix78iH5C9kt/Oy5PULJ9kNjnc5QxRL/HrCThAgp72IYkBeKbNsrfSAZHC0Xfiq+RZgLQODDf5tEQghwdB04t89/1O/w1cDnyilFU='; // 🌟 ใส่ Token ของคุณตรงนี้

// --- 1. WEB APP (doGet) ---
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  if (page === 'add') {
    return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('มอบหมายงานใหม่').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'calendar') {
    return HtmlService.createTemplateFromFile('CalendarPage').evaluate().setTitle('ปฏิทินการประชุม').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Dashboard').evaluate().setTitle('Dashboard ติดตามงาน').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- 2. ฟังก์ชันจัดการข้อมูล ---

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_USERS);
  let userMap = {};
  if (userSheet) {
    const userData = userSheet.getDataRange().getValues();
    for (let i = 1; i < userData.length; i++) {
      const name = userData[i][0];
      const email = userData[i][3];
      if (email && name) userMap[email.trim().toLowerCase()] = name.trim();
    }
  }

  const sheet = ss.getSheetByName(SHEET_NAME_TASK);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const tasks = data.slice(1); 
  
  let stats = { total: tasks.length, pending: 0, inProgress: 0, completed: 0, delayed: 0 };
  let userStatsMap = {}; 
  let uniqueCategories = new Set(); 
  
  tasks.forEach(row => {
    const status = row[10]; 
    if (status === 'Pending') stats.pending++;
    else if (status === 'In Progress') stats.inProgress++;
    else if (status === 'Completed') stats.completed++;
    else if (status === 'Delayed') stats.delayed++;

    const email = row[5] ? row[5].toString().trim().toLowerCase() : '';
    const name = userMap[email] ? userMap[email] : (row[5] || 'ไม่ระบุ');
    let cat = row[4] || 'ส่วนกลาง';

    if (name !== '') {
      uniqueCategories.add(cat);
      if (!userStatsMap[name]) userStatsMap[name] = {};
      if (!userStatsMap[name][cat]) userStatsMap[name][cat] = 0;
      userStatsMap[name][cat]++;
    }
  });

  return {
    stats: stats,
    categories: Array.from(uniqueCategories),
    userStats: Object.keys(userStatsMap).map(k => ({ name: k, categories: userStatsMap[k] })),
    recentTasks: tasks.slice(-30).reverse().map(row => ({
      id: row[0], type: row[4] || 'ส่วนกลาง', name: row[2],
      assignee: userMap[row[5] ? row[5].toString().trim().toLowerCase() : ''] || row[5],
      dueDate: row[8] ? new Date(row[8]).toLocaleDateString('th-TH') : '-',
      status: row[10], progress: row[11] || 0, remark: row[12] || '-'
    }))
  };
}

function getMeetingsData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MEETINGS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  return data.slice(1).map(r => ({ id: r[0], title: r[1], date: r[2], startTime: r[3], endTime: r[4], location: r[5] }));
}

// 🌟 ฟังก์ชัน: ดึงรายชื่อพนักงานสำหรับไปแสดงเป็นตัวเลือก Checkbox ในหน้าปฏิทิน
function getUserListForPicker() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_USERS);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const names = [];
    for (let i = 1; i < data.length; i++) {
      const name = data[i][0]; // คอลัมน์ A: ชื่อ-นามสกุล
      if (name && name.toString().trim() !== "") {
        names.push(name.toString().trim());
      }
    }
    return names.sort(); // เรียงลำดับรายชื่อตามตัวอักษร
  } catch (e) {
    return [];
  }
}

function addMeetingFromWebApp(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MEETINGS);
    if (!sheet) return { success: false, message: "ไม่พบชีต Meetings" };
    
    sheet.appendRow([
      'MTG-'+Date.now(), 
      payload.title, 
      "'"+payload.date, 
      "'"+payload.startTime, 
      "'"+payload.endTime, 
      payload.location, 
      Session.getActiveUser().getEmail(),
      payload.attendees // คอลัมน์ H: รายชื่อผู้ร่วมประชุม
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "บันทึกการประชุมสำเร็จ!" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function notifyTodayMeetings() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const mtgSheet = ss.getSheetByName(SHEET_NAME_MEETINGS);
  const userSheet = ss.getSheetByName(SHEET_NAME_USERS);
  if (!mtgSheet || !userSheet) return;

  const userData = userSheet.getDataRange().getValues();
  let userLineMap = {}; 
  for (let i = 1; i < userData.length; i++) {
    const name = userData[i][0] ? userData[i][0].toString().trim() : '';
    const lineId = userData[i][4] ? userData[i][4].toString().trim() : '';
    if (name && lineId.startsWith('U')) {
      userLineMap[name] = lineId; 
    }
  }

  const mtgData = mtgSheet.getDataRange().getDisplayValues();
  const todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");

  let personalNotifications = {};

  for (let i = 1; i < mtgData.length; i++) {
    const mtgDate = mtgData[i][2]; 
    
    if (mtgDate === todayStr) {
       const title = mtgData[i][1];     
       const startTime = mtgData[i][3]; 
       const endTime = mtgData[i][4];   
       const location = mtgData[i][5];  
       const attendeesStr = mtgData[i][7] || ''; 

       if (attendeesStr !== '') {
         const attendees = attendeesStr.split(',').map(name => name.trim());
         const meetingMsg = `📌 หัวข้อ: ${title}\n⏰ เวลา: ${startTime} - ${endTime}\n📍 สถานที่: ${location}`;

         attendees.forEach(name => {
           const lineId = userLineMap[name];
           if (lineId) {
             if (!personalNotifications[lineId]) {
               personalNotifications[lineId] = [];
             }
             personalNotifications[lineId].push(meetingMsg);
           }
         });
       }
    }
  }

  for (const lineId in personalNotifications) {
    const allMessages = personalNotifications[lineId].join("\n\n-------------------\n\n");
    const finalMessage = `🔔📢 สวัสดีจ้า วันนี้คุณมีตารางนัดหมายประชุมดังนี้นะคะ:\n\n${allMessages}`;
    
    sendLinePush(lineId, finalMessage);
  }
}

function sendLinePush(toUserId, text) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN },
      contentType: 'application/json',
      payload: JSON.stringify({
        to: toUserId,
        messages: [{ type: 'text', text: text }]
      })
    });
  } catch(e) {
    console.log("Error sending Push Message: " + e);
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const event = data.events[0];
  if (event.type === 'message' && event.message.type === 'text') {
    const userId = event.source.userId;
    const replyToken = event.replyToken;
    const message = "🌟สวัสดีค่ะ! รหัส LINE ID ของคุณคือ:\n\n" + userId + "\n\nรบกวนคัดลอกรหัสนี้ส่งให้ผู้ดูแลระบบ หรือนำไปใส่ในชีต Users (คอลัมน์ E) เพื่อเปิดรับแจ้งเตือนครับ";
    replyMessage(replyToken, message);
  }
  return ContentService.createTextOutput(JSON.stringify({content: "ok"})).setMimeType(ContentService.MimeType.JSON);
}

function replyMessage(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN },
    contentType: 'application/json',
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] })
  });
}

function updateTaskFromWebApp(updateDataObject) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_TASK);
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน " + SHEET_NAME_TASK };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == updateDataObject.taskId) { 
        rowIndex = i + 1; 
        break; 
      }
    }
    
    if (rowIndex == -1) return { success: false, message: "ไม่พบรหัสงาน: " + updateDataObject.taskId };
    
    let progressValue = parseFloat(updateDataObject.progress);
    if (isNaN(progressValue)) progressValue = 0;
    
    if (updateDataObject.status === 'Completed') {
      progressValue = 100;
    }
    
    sheet.getRange(rowIndex, 11).setValue(updateDataObject.status);  
    sheet.getRange(rowIndex, 12).setValue(progressValue);           
    
    if (updateDataObject.remark !== undefined) {
      sheet.getRange(rowIndex, 13).setValue(updateDataObject.remark); 
    }
    
    SpreadsheetApp.flush(); 
    return { success: true, message: "อัปเดตสถานะงานรหัส " + updateDataObject.taskId + " สำเร็จ!" };
    
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

function uploadPDF(filename, base64Data) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, filename);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    throw new Error("อัปโหลดไฟล์ไม่สำเร็จ: " + e.toString());
  }
}

// ฟังก์ชันดึงรายชื่อพนักงานสำหรับหน้า "มอบหมายงานใหม่"
function getUserList() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_USERS);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        users.push({
          name: data[i][0].toString().trim(),
          email: data[i][3] ? data[i][3].toString().trim() : '' // ดึงอีเมลจากคอลัมน์ D
        });
      }
    }
    return users;
  } catch (e) {
    return [];
  }
}

// ฟังก์ชันบันทึกข้อมูลฟอร์ม "มอบหมายงานใหม่" 
function addNewTask(formObject) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_TASK);
    if (!sheet) return { success: false, message: "ไม่พบชีต Task_List" };
    
    const taskId = "TASK-" + Date.now();
    
    sheet.appendRow([
      taskId,
      new Date(),
      formObject.taskName,
      formObject.description,
      formObject.department,
      formObject.assignee, 
      formObject.startDate,
      formObject.priority,
      formObject.dueDate,
      formObject.docLink || "", 
      "Pending",
      0,
      ""
    ]);
    
    SpreadsheetApp.flush();
    return { success: true, message: "เพิ่มงานใหม่เรียบร้อยแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 🌟 ฟังก์ชันสำหรับลบการประชุมออกจากชีต Meetings (เพิ่มให้แล้ว) 🌟
function deleteMeetingFromWebApp(meetingId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MEETINGS);
    if (!sheet) return { success: false, message: "ไม่พบชีต Meetings" };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // วนลูปค้นหาแถวที่มีรหัส ID ตรงกัน (คอลัมน์ A)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == meetingId) {
        rowIndex = i + 1; // คืนค่าตำแหน่งแถวใน Excel/Sheets
        break;
      }
    }
    
    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex); // สั่งลบแถวนั้นทิ้งทันที
      SpreadsheetApp.flush();
      return { success: true, message: "🗑️ ลบตารางนัดหมายการประชุมเรียบร้อยแล้ว!" };
    } else {
      return { success: false, message: "ไม่พบรหัสการประชุมที่ต้องการลบ" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
