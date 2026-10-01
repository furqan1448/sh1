// ================= ================= =================
// سكريبت قوقل الربط الحقيقي والمباشر 100%
// ================= ================= =================

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["اسم المشرفة", "كلمة المرور"]);
    usersSheet.appendRow(["فاطمة الطلحي", "123456"]);
    usersSheet.appendRow(["فاطمة الثمالي", "123456"]);
    usersSheet.appendRow(["مريم بالبيد", "123456"]);
    usersSheet.appendRow(["مستورة الخديدي", "123456"]);
  }
  
  var plansSheet = ss.getSheetByName("DailyPlans");
  if (!plansSheet) {
    plansSheet = ss.insertSheet("DailyPlans");
    plansSheet.appendRow(["تاريخ التسجيل", "اسم المشرفة", "اسم المعلمة", "الفترة", "اليوم", "التاريخ", "نوع الزيارة", "رابط الاستمارة"]);
  }
  
  var typesSheet = ss.getSheetByName("VisitTypes");
  if (!typesSheet) {
    typesSheet = ss.insertSheet("VisitTypes");
    typesSheet.appendRow(["نوع الزيارة"]);
    typesSheet.appendRow(["زيارة استطلاعية"]);
    typesSheet.appendRow(["زيارة تقويمية"]);
    typesSheet.appendRow(["زيارة توجيهية"]);
    typesSheet.appendRow(["زيارة متابعة"]);
  }

  var defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheetByName("ورقة1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

// 1. القراءة الحقيقية والمباشرة من قوقل شيت
function doGet(e) {
  setupDatabase();
  var action = e ? e.parameter.action : "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // جلب جميع الخطط الحقيقية المخزنة في الشيت لعرضها في الموقع
  if (action == "getDailyPlans") {
    var sheet = ss.getSheetByName("DailyPlans");
    var data = sheet.getDataRange().getValues();
    var plans = [];
    
    // التجاوز عن الصف الأول (العناوين)
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) { // التأكد من وجود اسم المشرفة
        plans.push({
          timestamp: data[i][0],
          supervisor: data[i][1],
          teacher: data[i][2],
          period: data[i][3],
          day: data[i][4],
          date: data[i][5],
          visitType: data[i][6],
          formUrl: data[i][7]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(plans.reverse())).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action == "login") {
    var user = e.parameter.user;
    var pass = e.parameter.pass;
    var usersSheet = ss.getSheetByName("Users");
    var usersData = usersSheet.getDataRange().getValues();
    
    for (var i = 1; i < usersData.length; i++) {
      if (usersData[i][0] == user && String(usersData[i][1]) == pass) {
        return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "كلمة المرور غير صحيحة"})).setMimeType(ContentService.MimeType.JSON);
  }
}

// 2. الكتابة والحفظ الحقيقي في قوقل شيت
function doPost(e) {
  setupDatabase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  if (data.action == "saveDailyPlan") {
    var sheet = ss.getSheetByName("DailyPlans");
    
    sheet.appendRow([
      new Date().toLocaleString('ar-SA'),
      data.supervisor,
      data.teacher,
      data.period,
      data.day,
      data.date,
      data.visitType,
      data.formUrl
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}
