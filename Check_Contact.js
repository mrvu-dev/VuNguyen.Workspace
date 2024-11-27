// Kiểm tra thông tin liên hệ các team khác

function check_Contact() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1WdgqEJal7KKEnch0D4GmhtpdA3Z7dyeI9KcFJFNz6yg/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("Sheet1");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Check_Contact")

  var hub_Name = destination_Sheet.getRange("B1").getValue();
  var contact_Type = "Đối soát";
  var directory_Data = source_Sheet.getRange("B:E").getValues();
  var data_Contacts = [];
  
  for (var row = 1; row < directory_Data.length; row++) {
    if (directory_Data[row][0] == hub_Name && directory_Data[row][2] == contact_Type) {
      data_Contacts.push(directory_Data[row]);
    }
  }

  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','Check_Contact','A4:D')

  if (data_Contacts.length > 0) {
    destination_Sheet.getRange(4, 1, data_Contacts.length, data_Contacts[0].length).setValues(data_Contacts);
  } else {
    destination_Sheet.getRange("A4").setValue("Không truy vấn được dữ liệu");
  }
}