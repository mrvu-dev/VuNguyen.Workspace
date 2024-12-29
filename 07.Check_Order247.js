// Kiểm tra đơn nội bộ 247

function Check_Order247() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1WTi74y4xjsPzLDd5cbNFjzL5HV2gP930KfZ5QoWK-Lk/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("Dịch vụ 247");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Check_Order247")

  var data_A77 = source_Sheet.getRange("B5:X").getValues();
  var data = [];
  var hub_Code = 296;

  for (var row = 0; row < data_A77.length; row++) {
    if(data_A77[row][7] == hub_Code || data_A77[row][14] == hub_Code) {
      data.push(data_A77[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Check_Order247','B4:X')

  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(4, 2, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}
