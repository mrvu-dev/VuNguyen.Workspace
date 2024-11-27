// Cập nhật lịch làm việc 8 ngày tiếp theo

function LLV() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1SlXiDuY1U9pv5iW3brkvmg_Ng0hwz9lxqjdM4X_rYww/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("LLV 2024");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("LLV");

  var SPX_ID = destination_Sheet.getRange("C3:C22").getValues();
  var SPX_IDs = [];
  
  for (var i = 0; i < SPX_ID.length; i++){
    SPX_IDs.push(SPX_ID[i][0]);
  }  

  // Đọc toàn bộ sheet nguồn [source_Sheet] và lưu vào [directory_Data]
  var directory_Data = source_Sheet.getRange(1, 1, source_Sheet.getLastRow(), source_Sheet.getLastColumn()).getValues();


  // Khai báo và xử lý dữ liệu
  var data_LLV = [];
  const day = new Date();
  
  for (var row = 0; row < directory_Data.length; row++) {
    var n = SPX_IDs.indexOf(directory_Data[row][3].trim());
    if (n != -1) {
      for (var col = 0; col < directory_Data.length; col++) {
        if (formattedDate(directory_Data[1][col]) == formattedDate(day)) {
          data_LLV.push([SPX_IDs[n], directory_Data[row][col], directory_Data[row][col+1], directory_Data[row][col+2], directory_Data[row][col+3], directory_Data[row][col+4], directory_Data[row][col+5], directory_Data[row][col+6], directory_Data[row][col+7]]);
          break; 
        }
      }
    }    
  }  

  var data_LLVs = [];
  for (i = 0; i < data_LLV.length; i++){
    var n = SPX_IDs.indexOf(data_LLV[i][0]);
    data_LLVs.push([data_LLV[n][1], data_LLV[n][2], data_LLV[n][3], data_LLV[n][4], data_LLV[n][5], data_LLV[n][6], data_LLV[n][7], data_LLV[n][8]]);
  }


  // Hiển thị mảng data_LLs vào sheet
  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','LLV','G3:N');
  if (data_LLVs.length > 0) {
    destination_Sheet.getRange(3, 7, data_LLVs.length, data_LLVs[0].length).setValues(data_LLVs);
    destination_Sheet.getRange('A1').setValue('Last Update: ' + day.toLocaleDateString('vi-VN')+ ' ' + day.toLocaleTimeString('vi-VN'))
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}