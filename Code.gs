// ================= ================= =================
// سكريبت الإدارة والإشراف المتكامل مع النماذج
// ================= ================= =================

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ورقة المستخدمين والحسابات
  var usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["اسم المستخدم", "كلمة المرور", "الرتبة"]);
    usersSheet.appendRow(["الإدارة العامة", "123456", "Admin"]);
    usersSheet.appendRow(["فاطمة الطلحي", "123456", "Supervisor"]);
    usersSheet.appendRow(["فاطمة الثمالي", "123456", "Supervisor"]);
    usersSheet.appendRow(["مريم بالبيد", "123456", "Supervisor"]);
    usersSheet.appendRow(["مستورة الخديدي", "123456", "Supervisor"]);
  }
  
  // 2. الخطط اليومية
  var plansSheet = ss.getSheetByName("DailyPlans");
  if (!plansSheet) {
    plansSheet = ss.insertSheet("DailyPlans");
    plansSheet.appendRow(["تاريخ التسجيل", "اسم المشرفة", "اسم المعلمة", "الفترة", "اليوم", "التاريخ", "نوع الزيارة", "رابط الاستمارة"]);
  }

  // 3. أحوال الموظفات
  var empSheet = ss.getSheetByName("Employees");
  if (!empSheet) {
    empSheet = ss.insertSheet("Employees");
    empSheet.appendRow(["تاريخ التسجيل", "المشرفة", "اسم المعلمة", "المؤهل العلمي", "المؤهل في القرآن", "الفئة", "عدد الدورات", "الجنسية", "رقم الهاتف"]);
  }

  // 4. المرفقات العامة
  var attachSheet = ss.getSheetByName("Attachments");
  if (!attachSheet) {
    attachSheet = ss.insertSheet("Attachments");
    attachSheet.appendRow(["تاريخ التسجيل", "المشرفة", "عنوان المرفق", "رابط المرفق"]);
  }

  // 5. نماذج الإدارة للمشرفات (Admin Forms)
  var adminFormsSheet = ss.getSheetByName("AdminForms");
  if (!adminFormsSheet) {
    adminFormsSheet = ss.insertSheet("AdminForms");
    adminFormsSheet.appendRow(["تاريخ النشر", "عنوان النموذج", "وصف النموذج", "رابط النموذج"]);
  }

  var defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheetByName("ورقة1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

function doGet(e) {
  setupDatabase();
  var action = e ? e.parameter.action : "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action == "getDailyPlans") {
    var sheet = ss.getSheetByName("DailyPlans");
    var data = sheet.getDataRange().getValues();
    var plans = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        plans.push({
          timestamp: data[i][0], supervisor: data[i][1], teacher: data[i][2],
          period: data[i][3], day: data[i][4], date: data[i][5], visitType: data[i][6], formUrl: data[i][7]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(plans.reverse())).setMimeType(ContentService.MimeType.JSON);
  }

  if (action == "getEmployees") {
    var sheet = ss.getSheetByName("Employees");
    var data = sheet.getDataRange().getValues();
    var emp = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][2]) {
        emp.push({
          timestamp: data[i][0], supervisor: data[i][1], teacher: data[i][2],
          degree: data[i][3], quranDegree: data[i][4], category: data[i][5], courses: data[i][6],
          nationality: data[i][7], phone: data[i][8]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(emp.reverse())).setMimeType(ContentService.MimeType.JSON);
  }

  if (action == "getAttachments") {
    var sheet = ss.getSheetByName("Attachments");
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][2]) {
        list.push({ timestamp: data[i][0], supervisor: data[i][1], title: data[i][2], url: data[i][3] });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(list.reverse())).setMimeType(ContentService.MimeType.JSON);
  }

  if (action == "getAdminForms") {
    var sheet = ss.getSheetByName("AdminForms");
    var data = sheet.getDataRange().getValues();
    var forms = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        forms.push({ timestamp: data[i][0], title: data[i][1], description: data[i][2], url: data[i][3] });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(forms.reverse())).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  setupDatabase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  if (data.action == "saveDailyPlan") {
    var sheet = ss.getSheetByName("DailyPlans");
    sheet.appendRow([new Date().toLocaleString('ar-SA'), data.supervisor, data.teacher, data.period, data.day, data.date, data.visitType, data.formUrl]);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action == "saveEmployee") {
    var sheet = ss.getSheetByName("Employees");
    sheet.appendRow([new Date().toLocaleString('ar-SA'), data.supervisor, data.teacher, data.degree, data.quranDegree, data.category, data.courses, data.nationality, data.phone]);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action == "saveAttachment") {
    var sheet = ss.getSheetByName("Attachments");
    sheet.appendRow([new Date().toLocaleString('ar-SA'), data.supervisor, data.title, data.url]);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action == "saveAdminForm") {
    var sheet = ss.getSheetByName("AdminForms");
    sheet.appendRow([new Date().toLocaleString('ar-SA'), data.title, data.description, data.url]);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}
