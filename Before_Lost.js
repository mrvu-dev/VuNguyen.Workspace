// Theo dõi phản hồi đơn hàng trước khi cập nhật Lost hành trình

function Before_Lost() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/15MjG4w-i3YbNyssP_WuuQBcwpasWj4YJwU81mJKOq6Q/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("HCM");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Before_Lost")

  var hub_Name = "50-HCM Tan Binh/Bach Dang";
  var data_HCM = source_Sheet.getRange("A:J").getValues();
  var data = [];
  data.push(data_HCM[0]);

  for (var row = 0; row < data_HCM.length; row++) {
    if (data_HCM[row][3] == hub_Name) {
      data.push(data_HCM[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Before_Lost','A3:K')

  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}
