// Kiểm tra đơn nội bộ 247

function Check_OrderSPX() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1R-DfZaTKOiZLGj3u-Lza2kxkvGC0YEQmqAn2JAHJs3A/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("Đơn nội bộ (Ops/HR/IT/CS/DOP)");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Check_OrderSPX")

  var data_A77 = source_Sheet.getRange("B5:AM").getValues();
  var data = [];
  var hub_Code = 296;

  for (var row = 0; row < data_A77.length; row++) {
    if(data_A77[row][6] == hub_Code || data_A77[row][14] == hub_Code) {
      data.push(data_A77[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Check_OrderSPX','B5:AM')

  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(4, 2, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}
