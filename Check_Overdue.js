// Kiểm tra quá hạn đơn hàng

function Check_Overdue() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1PVTzcLIzSHejfTyWUES4O2EZBP5tTfvLNU6y7H6ohIs/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("Tan Binh/Bach Dang");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Check_Overdue")

  var data_A77 = source_Sheet.getRange("A:N").getValues();
  var data = [];
  data.push(data_A77[0]);
  var hub_Name = "50-HCM Tan Binh/Bach Dang Hub";

  for (var row = 0; row < data_A77.length; row++) {
    if(data_A77[row][2].trim() == hub_Name) {
      data.push(data_A77[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Check_Overdue','A3:O')

  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}
