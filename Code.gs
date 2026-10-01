// كود Google Apps Script الداعم لعملية تسجيل الدخول وقراءة أنواع الزيارة وحفظ الخطة اليومية

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. التحقق من تسجيل الدخول
  if (action == "login") {
    var user = e.parameter.user;
    var pass = e.parameter.pass;
    var usersSheet = ss.getSheetByName("Users");
    var usersData = usersSheet.getDataRange().getValues();
    
    for (var i = 1; i < usersData.length; i++) {
      if (usersData[i][0] == user && usersData[i][1] == pass) {
        return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "كلمة المرور غير صحيحة"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. جلب أنواع الزيارات من شيت VisitTypes
  if (action == "getVisitTypes") {
    var typeSheet = ss.getSheetByName("VisitTypes");
    if (!typeSheet) {
      return ContentService.createTextOutput(JSON.stringify(["زيارة استطلاعية", "زيارة تقويمية", "زيارة توجيهية", "زيارة متابعة"])).setMimeType(ContentService.MimeType.JSON);
    }
    var data = typeSheet.getDataRange().getValues();
    var types = [];
    for (var j = 1; j < data.length; j++) {
      if (data[j][0]) types.push(data[j][0]);
    }
    return ContentService.createTextOutput(JSON.stringify(types)).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  // 3. حفظ بيانات الخطة اليومية
  if (data.action == "saveDailyPlan") {
    var sheet = ss.getSheetByName("DailyPlans");
    if (!sheet) {
      sheet = ss.insertSheet("DailyPlans");
      sheet.appendRow(["تاريخ التسجيل", "اسم المشرفة", "اسم المعلمة", "الفترة", "اليوم", "التاريخ", "نوع الزيارة", "رابط الاستمارة"]);
    }
    
    sheet.appendRow([
      new Date(),
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
