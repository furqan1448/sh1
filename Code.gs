// كود Google Apps Script للتوصيل بقاعدة بيانات Google Sheets
// أضف هذا الكود في Extensions > Apps Script داخل شيت قوقل

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("Users");
  var visitsSheet = ss.getSheetByName("Visits");
  
  var data = JSON.parse(e.postData.contents);
  
  // التحقق من الرقم السري
  var usersData = usersSheet.getDataRange().getValues();
  var isAuthenticated = false;
  
  for (var i = 1; i < usersData.length; i++) {
    if (usersData[i][0] == data.supervisor && usersData[i][1] == data.password) {
      isAuthenticated = true;
      break;
    }
  }
  
  if (!isAuthenticated) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": "الرقم السري غير صحيح للمشرفة المحددة"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // إضافة البيانات في شيت الزيارات
  visitsSheet.appendRow([
    new Date(),
    data.supervisor,
    data.school,
    data.rating,
    data.notes
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({
    "status": "success",
    "message": "تم تسجيل الزيارة بنجاح"
  })).setMimeType(ContentService.MimeType.JSON);
}
