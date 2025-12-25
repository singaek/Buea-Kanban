// Sheet ID ของ Google Sheet
var SHEET_ID = "SHEET_ID";

// ID ของโฟลเดอร์ Google Drive หลัก
var DRIVE_FOLDER_ID = "DRIVE_FOLDER_ID";

// ชื่อชีตต่างๆ
var CLASSROOMS_SHEET_NAME = "Classrooms";
var TEACHERS_SHEET_NAME = "Teachers";
var HOMEWORK_SHEET_NAME = "Homework";
var SUBMISSIONS_SHEET_NAME = "Submissions";

// Admin Credentials
var ADMIN_USERNAME = "Lover";
var ADMIN_PASSWORD = "MS210555*";

function doPost(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var lock = LockService.getScriptLock();
  
  try {
    // ลดเวลารอ Lock ลงเล็กน้อยเพื่อความเร็ว แต่ยังคงความปลอดภัยข้อมูล
    if (action.includes('update') || action.includes('create') || action.includes('delete') || action.includes('add')) {
      lock.waitLock(20000); 
    }

    switch (action) {
      case 'login':
        return createTextOutput(handleLogin(e.parameter.username, e.parameter.password));
      case 'getClasses':
        return createTextOutput(getClasses(ss));
      case 'getHomeworkForStudent':
        return createTextOutput(getHomeworkForStudent(ss, e.parameter.classroom, e.parameter.studentId));
      case 'createHomework':
        return createTextOutput(createHomework(ss, e.parameter));
      case 'getTeacherHomework':
        return createTextOutput(getTeacherHomework(ss, e.parameter.username, e.parameter.classroom || null));
      case 'deleteHomework':
        return createTextOutput(archiveHomework(ss, e.parameter.homeworkId));
      case 'getStudentSubmissionsForGrading':
        return createTextOutput(getStudentSubmissionsForGrading(ss, e.parameter.homeworkId, e.parameter.classroom));
      case 'updateSubmission':
        return createTextOutput(updateSubmission(ss, e.parameter));
      case 'getAdminData':
        return createTextOutput(getAdminData(ss));
      case 'updateClassroomStudentCount':
        return createTextOutput(updateClassroomStudentCount(ss, e.parameter.classroomName, e.parameter.studentCount));
      case 'addTeacher':
        return createTextOutput(addTeacher(ss, e.parameter.username, e.parameter.password, e.parameter.fullName, e.parameter.subjects));
      case 'deleteTeacher':
        return createTextOutput(deleteTeacher(ss, e.parameter.username));
      default:
        return createTextOutput({ success: false, message: "Invalid action: " + action });
    }
  } catch (error) {
    Logger.log("Error in doPost for action " + action + ": " + error.message);
    return createTextOutput({ success: false, message: "Server error: " + error.message });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function createTextOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [];
    if (sheetName === CLASSROOMS_SHEET_NAME) headers = ["ClassroomName", "StudentCount"];
    else if (sheetName === TEACHERS_SHEET_NAME) headers = ["Username", "Password", "FullName", "Subjects"];
    else if (sheetName === HOMEWORK_SHEET_NAME) headers = ["Id", "Classroom", "Subject", "HomeworkName", "Details", "DueDate", "MaxScore", "OnlineSubmission", "TeacherUsername", "Status"];
    else if (sheetName === SUBMISSIONS_SHEET_NAME) headers = ["Id", "HomeworkId", "StudentId", "Classroom", "SubmissionLink", "Score", "Status", "Comments", "SubmissionDate"];
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    return [];
  }
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0];
  var data = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    data.push(obj);
  }
  return data;
}

function addRowToSheet(ss, sheetName, dataObject) {
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRowValues = headers.map(function(header) {
    return dataObject[header] !== undefined ? dataObject[header] : "";
  });
  sheet.appendRow(newRowValues);
  return true;
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

/**
 * อัปเดตข้อมูลการส่งงาน (ปรับปรุงให้เสถียรขึ้นสำหรับการส่งซ้ำ)
 */
function updateSubmission(ss, params) {
  try {
    var studentId = parseInt(params.studentId);
    var homeworkId = params.homeworkId;
    var classroom = params.classroom;
    
    var submissionLink = params.submissionLink; 
    var submissionDate = null;
    var newStatus = params.status;

    // --- ส่วนตรรกะการอัปโหลดไฟล์ ---
    if (params.fileData && params.fileName) {
      var homeworks = getSheetData(ss, HOMEWORK_SHEET_NAME);
      var homework = homeworks.find(function(hw) { return hw.Id == homeworkId });
      if (!homework) { throw new Error("ไม่พบข้อมูลการบ้าน ID: " + homeworkId); }
      
      var homeworkName = homework.HomeworkName.replace(/[/\\?%*:|"<>]/g, '-'); 
      var mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var classroomFolder = getOrCreateFolder(mainFolder, classroom);
      var homeworkFolder = getOrCreateFolder(classroomFolder, homeworkName);

      var dataUrl = params.fileData;
      var splitData = dataUrl.split(','); 
      var mimeType = params.mimeType || (splitData[0].match(/:(.*?);/) ? splitData[0].match(/:(.*?);/)[1] : 'application/octet-stream');
      var decodedData = Utilities.base64Decode(splitData[1]);
      var blob = Utilities.newBlob(decodedData, mimeType, params.fileName);

      var newFileName = "เลขที่" + studentId + "_" + params.fileName;
      
      // ลบไฟล์เก่าที่มีชื่อเดียวกัน เพื่อประหยัดพื้นที่และป้องกันความสับสน
      var oldFiles = homeworkFolder.getFilesByName(newFileName);
      while(oldFiles.hasNext()){
        oldFiles.next().setTrashed(true);
      }

      var newFile = homeworkFolder.createFile(blob);
      newFile.setName(newFileName); 

      submissionLink = newFile.getUrl();
      submissionDate = new Date().toISOString(); 
      
      // ถ้ามีการส่งไฟล์ บังคับให้สถานะเป็น Submitted ทันที (แก้บั๊กสถานะไม่เปลี่ยน)
      newStatus = 'Submitted';
    }

    // --- ส่วนบันทึกลง Google Sheet ---
    var submissionsSheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
    var headers = submissionsSheet.getRange(1, 1, 1, submissionsSheet.getLastColumn()).getValues()[0];
    var values = submissionsSheet.getDataRange().getValues();
    
    var rowIndexToUpdate = -1;
    var existingSubmissionData = null;

    for(var i = 1; i < values.length; i++) {
        if (values[i][headers.indexOf('HomeworkId')] == homeworkId &&
            values[i][headers.indexOf('StudentId')] == studentId &&
            values[i][headers.indexOf('Classroom')] == classroom) {
            rowIndexToUpdate = i + 1; 
            existingSubmissionData = values[i]; 
            break;
        }
    }

    if (rowIndexToUpdate !== -1) { 
        var rowData = existingSubmissionData;
        rowData[headers.indexOf('Score')] = (params.score !== undefined) ? params.score : rowData[headers.indexOf('Score')];
        rowData[headers.indexOf('Comments')] = (params.comments !== undefined) ? params.comments : rowData[headers.indexOf('Comments')];
        
        // อัปเดตสถานะ
        if (newStatus) {
           rowData[headers.indexOf('Status')] = newStatus;
        }
        
        if (submissionLink) {
          rowData[headers.indexOf('SubmissionLink')] = submissionLink;
          rowData[headers.indexOf('SubmissionDate')] = submissionDate;
        }
        
        submissionsSheet.getRange(rowIndexToUpdate, 1, 1, rowData.length).setValues([rowData]);

    } else { 
        var newSubmissionDataObject = {
          Id: Utilities.getUuid(),
          HomeworkId: homeworkId,
          StudentId: studentId,
          Classroom: classroom,
          SubmissionLink: submissionLink || '',
          Score: params.score || '',
          Status: newStatus || 'Submitted',
          Comments: params.comments || '',
          SubmissionDate: submissionDate || ''
        };
        addRowToSheet(ss, SUBMISSIONS_SHEET_NAME, newSubmissionDataObject);
    }

    return { success: true, message: "ดำเนินการเรียบร้อยแล้ว" };

  } catch (e) {
    Logger.log("Error in updateSubmission: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

// -----------------------------------------------------------------------------
//  Function อื่นๆ คงเดิม
// -----------------------------------------------------------------------------
function handleLogin(username, password) {
  try {
    if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
      return { success: true, role: 'admin', fullName: 'ผู้ดูแลระบบ', username: 'admin' };
    }
    var teachers = getSheetData(SpreadsheetApp.openById(SHEET_ID), TEACHERS_SHEET_NAME);
    var teacher = teachers.find(function(t) { return t.Username === username && t.Password === password; });
    
    if (teacher) {
        return { success: true, role: 'teacher', username: teacher.Username, fullName: teacher.FullName, subjects: teacher.Subjects };
    }
    return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function getClasses(ss) {
  var classrooms = getSheetData(ss, CLASSROOMS_SHEET_NAME);
  return { success: true, classrooms: classrooms };
}

function getHomeworkForStudent(ss, classroom, studentId) {
  var allHomework = getSheetData(ss, HOMEWORK_SHEET_NAME);
  var allSubmissions = getSheetData(ss, SUBMISSIONS_SHEET_NAME);
  
  var relevantHomework = allHomework.filter(function(hw) {
    return hw.Classroom === classroom && hw.Status === 'Active';
  });

  var studentHomework = relevantHomework.map(function(hw) {
    var studentSubmission = allSubmissions.find(function(sub) {
      return sub.HomeworkId == hw.Id && sub.StudentId == studentId && sub.Classroom === classroom;
    });

    var homeworkWithStatus = JSON.parse(JSON.stringify(hw));
    homeworkWithStatus.SubmissionStatus = studentSubmission ? studentSubmission.Status : 'Pending';
    homeworkWithStatus.Score = studentSubmission ? (studentSubmission.Score !== '' ? studentSubmission.Score : '') : '';
    homeworkWithStatus.Comments = studentSubmission ? studentSubmission.Comments : '';
    homeworkWithStatus.SubmissionLink = studentSubmission ? studentSubmission.SubmissionLink : '';
    homeworkWithStatus.SubmissionDate = studentSubmission ? studentSubmission.SubmissionDate : '';
    return homeworkWithStatus;
  });
  
  return { success: true, homework: studentHomework };
}

function createHomework(ss, params) {
  var newHomework = {
    Id: Utilities.getUuid(),
    Classroom: params.classroom,
    Subject: params.subject,
    HomeworkName: params.homeworkName,
    Details: params.details,
    DueDate: params.dueDate,
    MaxScore: parseInt(params.maxScore),
    OnlineSubmission: params.onlineSubmission === 'true',
    TeacherUsername: params.teacherUsername,
    Status: 'Active'
  };
  addRowToSheet(ss, HOMEWORK_SHEET_NAME, newHomework);
  return { success: true, message: "สร้างการบ้านเรียบร้อยแล้ว" };
}

function getTeacherHomework(ss, username, classroom) {
  var allHomework = getSheetData(ss, HOMEWORK_SHEET_NAME).filter(function(hw) { return hw.Status === 'Active' });
  var filteredHomework = [];

  if (username === 'admin') {
     filteredHomework = classroom ? allHomework.filter(function(hw) { return hw.Classroom === classroom; }) : allHomework;
  } else {
    filteredHomework = allHomework.filter(function(hw) {
      return hw.TeacherUsername === username && (!classroom || hw.Classroom === classroom);
    });
  }
  return { success: true, homework: filteredHomework };
}

function updateRowInSheet(ss, sheetName, keyColumn, keyToFind, newDataObject) {
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getDataRange().getValues();
  var keyColIndex = headers.indexOf(keyColumn);
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][keyColIndex] == keyToFind) {
      var rowToUpdate = values[i];
      headers.forEach(function(header, j) {
        if (newDataObject.hasOwnProperty(header)) {
          rowToUpdate[j] = newDataObject[header];
        }
      });
      sheet.getRange(i + 1, 1, 1, rowToUpdate.length).setValues([rowToUpdate]);
      return true;
    }
  }
  return false;
}

function archiveHomework(ss, homeworkId) {
  var updateResult = updateRowInSheet(ss, HOMEWORK_SHEET_NAME, "Id", homeworkId, { Status: 'Archived' });
  if (updateResult) return { success: true, message: "ลบการบ้านเรียบร้อยแล้ว" };
  else return { success: false, message: "ไม่พบการบ้าน" };
}

function getStudentSubmissionsForGrading(ss, homeworkId, classroom) {
  var classroomsData = getSheetData(ss, CLASSROOMS_SHEET_NAME);
  var targetClassroom = classroomsData.find(function(cls) { return cls.ClassroomName === classroom; });
  if (!targetClassroom) return { success: false, message: "ไม่พบห้องเรียน" };

  var studentCount = targetClassroom.StudentCount;
  var allStudentsInClass = [];
  for (var i = 1; i <= studentCount; i++) {
    allStudentsInClass.push({ StudentId: i, Score: '', Status: 'Pending', SubmissionLink: '', Comments: '', SubmissionDate: '' });
  }

  var relevantSubmissions = getSheetData(ss, SUBMISSIONS_SHEET_NAME).filter(function(sub) {
    return sub.HomeworkId == homeworkId && sub.Classroom === classroom;
  });

  var submissionMap = new Map();
  relevantSubmissions.forEach(function(sub) { submissionMap.set(sub.StudentId, sub); });

  allStudentsInClass.forEach(function(student) {
    var existingSub = submissionMap.get(student.StudentId);
    if (existingSub) {
      student.Score = existingSub.Score !== undefined ? existingSub.Score : '';
      student.Status = existingSub.Status || 'Pending';
      student.SubmissionLink = existingSub.SubmissionLink || '';
      student.Comments = existingSub.Comments || '';
      student.SubmissionDate = existingSub.SubmissionDate || '';
    }
  });

  return { success: true, students: allStudentsInClass };
}

function getAdminData(ss) {
    var classrooms = getSheetData(ss, CLASSROOMS_SHEET_NAME);
    var teachers = getSheetData(ss, TEACHERS_SHEET_NAME).map(function(t) {
        return { Username: t.Username, FullName: t.FullName, Subjects: t.Subjects };
    });
    return { success: true, classrooms: classrooms, teachers: teachers };
}

function updateClassroomStudentCount(ss, classroomName, studentCount) {
  var result = updateRowInSheet(ss, CLASSROOMS_SHEET_NAME, "ClassroomName", classroomName, { StudentCount: parseInt(studentCount) });
  if (result) return { success: true, message: "อัปเดตเรียบร้อย" };
  else return { success: false, message: "ไม่พบห้องเรียน" };
}

function addTeacher(ss, username, password, fullName, subjects) {
  var teachersData = getSheetData(ss, TEACHERS_SHEET_NAME);
  if (teachersData.some(function(t) { return t.Username === username; })) {
    return { success: false, message: "ชื่อผู้ใช้นี้มีอยู่แล้ว" };
  }
  addRowToSheet(ss, TEACHERS_SHEET_NAME, { Username: username, Password: password, FullName: fullName, Subjects: subjects });
  return { success: true, message: "เพิ่มครูเรียบร้อยแล้ว" };
}

function deleteRowFromSheet(ss, sheetName, keyColumn, keyToDelete) {
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getDataRange().getValues();
  var keyColIndex = headers.indexOf(keyColumn);
  
  for (var i = values.length - 1; i > 0; i--) { 
    if (values[i][keyColIndex] == keyToDelete) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function deleteTeacher(ss, username) {
  var result = deleteRowFromSheet(ss, TEACHERS_SHEET_NAME, "Username", username);
  if (result) return { success: true, message: "ลบครูเรียบร้อยแล้ว" };
  else return { success: false, message: "ไม่พบครู" };
}