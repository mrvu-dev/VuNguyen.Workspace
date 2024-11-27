// Follow thông tin rider

function Data_Rider() {
    // Sheet lấy data
    var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1SYvlNV6tUxRZ2OwOfrbilWLaNt8rfGMh5HHfZORE_dw/edit");
    var source_Sheet = source_Spreadsheet.getSheetByName("Rider BĐ");
  
    // Sheet chứa data
    var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
    var destination_Sheet = destination_Spreadsheet.getSheetByName("Data_Rider");
    
    var source_data = source_Sheet.getRange("C2:H").getValues();
    var dataRider = new Array();
  
    for (var row = 0; row < source_data.length; row++) {
      if (source_data[row][0] != "") {
        dataRider.push([source_data[row][2], source_data[row][0], source_data[row][1], source_data[row][5]]);
      }
    }
  
    try {
      destination_Sheet.getRange("B3:E").clearContent();
    } catch {
      SpreadsheetApp.getActiveSheet().toast("Không có dữ liệu để xóa", "Done !", 3);
    }
    
    const timeNow = new Date();
    if (dataRider.length > 0) {
      destination_Sheet.getRange(3, 2, dataRider.length, dataRider[0].length).setValues(dataRider);
      destination_Sheet.getRange("A1").setValue("Late Update: " + timeNow.toLocaleDateString("vi-VN") + ' ' + timeNow.toLocaleTimeString('vi-VN'));
    } else {
      destination_Sheet.getRange("A1").setValue("Sheet nguồn gặp lỗi");
    }
}
  