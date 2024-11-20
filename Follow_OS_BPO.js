// Kiểm tra chấm công OS và BPO

function Follow_OS_BPO() {
  // Khai báo sheet lấy dữ liệu
  var directorySpr = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gd662-yEDRBjoeffKRtKS9N0A2usNM7p_fAABl8j0tw/edit");
  var directorySheet_in = directorySpr.getSheetByName('Checkin');
  var directorySheet_out = directorySpr.getSheetByName('Checkout');

  // Khai báo sheet chứa dữ liệu
  var finderSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var finderSheet = finderSpreadsheet.getSheetByName('Follow_OS&BPO');

  
  var directoryData_in = directorySheet_in.getRange("A3:K").getValues();
  var directoryData_out = directorySheet_out.getRange("A3:K").getValues(); 
  var foundContacts_in = [];
  var foundContacts_out = [];
  var hubName = finderSheet.getRange("B1").getValue();
  var today = finderSheet.getRange("B2").getValue();
  var checkToday = formattedDate(today);

  for (var i = 1; i < directoryData_in.length; i++) {    
    if (directoryData_in[i][6] == hubName && formattedDate(directoryData_in[i][0]) == checkToday) {
      foundContacts_in.push([directoryData_in[i][0], directoryData_in[i][1], directoryData_in[i][2], directoryData_in[i][8], directoryData_in[i][9]]);
    }
  }

  for (var i = 1; i < directoryData_out.length; i++) {    
    if (directoryData_out[i][5] == hubName && formattedDate(directoryData_out[i][0]) == checkToday) {
      foundContacts_out.push([directoryData_out[i][0], directoryData_out[i][1], directoryData_out[i][2], directoryData_out[i][7], directoryData_out[i][8]]);
    }
  }

  // Xóa dữ liệu cũ
  refreshSheet("1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A","Follow_OS&BPO","A6:K")

  // Hiển thị dữ liệu
  const timeNow = new Date();
  finderSheet.getRange("E1").setValue("Last update: " + timeNow.toLocaleDateString('vi-VN')+ ' ' + timeNow.toLocaleTimeString('vi-VN'));

  if (foundContacts_in.length > 0) {
    finderSheet.getRange(6, 1, foundContacts_in.length, foundContacts_in[0].length).setValues(foundContacts_in);
  } else {
    finderSheet.getRange("B6").setValue("Chưa có thông tin check-in");
  }

  if (foundContacts_out.length > 0) {
    finderSheet.getRange(6, 7, foundContacts_out.length, foundContacts_out[0].length).setValues(foundContacts_out);
  } else {
    finderSheet.getRange("H6").setValue("Chưa có thông tin check-out");
  }
}


// Theo dõi kế hoạch booking OS 

function PlanOS() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1SYvlNV6tUxRZ2OwOfrbilWLaNt8rfGMh5HHfZORE_dw/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("Weekly OS");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Follow_OS&BPO")

  var hub_Name = "50-HCM Tan Binh/Bach Dang";
  var data = source_Sheet.getRange("B2:AB").getValues();
  const today = new Date();
  var week = getFirstDayOfWeek(today);
  const daynow = week.setHours(0, 0, 0, 0);
  var data_plan = [];

  for (var row = 0; row < data.length; row++) {
    if (data[row][0] != 0 && data[row][1] == hub_Name && timestamp(data[row][2]) >= timestamp(daynow)) {
      data_plan.push(data[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Follow_OS&BPO','N3:AN')

  const timenow = new Date();
  if (data_plan.length > 0) {
    destination_Sheet.getRange(3, 14, data_plan.length, data_plan[0].length).setValues(data_plan);
    destination_Sheet.getRange("M1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("M1").setValue("Không truy vấn được dữ liệu");
  }
}
