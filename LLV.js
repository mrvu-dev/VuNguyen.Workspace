// Cập nhật lịch làm việc 8 ngày tiếp theo

function LLV() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1SlXiDuY1U9pv5iW3brkvmg_Ng0hwz9lxqjdM4X_rYww/edit");
  var source_Sheet = source_Spreadsheet.getSheetByName("LLV 2025");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("LLV");

  var SPX_ID = destination_Sheet.getRange("B3:B").getValues();
  var SPX_IDs = [];
  
  for (var i = 0; i < SPX_ID.length; i++){
    SPX_IDs.push(SPX_ID[i][0]);
  }  

  // Đọc toàn bộ sheet nguồn [source_Sheet] và lưu vào [directory_Data]
  // var directory_Data = source_Sheet.getRange(1, 1, source_Sheet.getLastRow(), source_Sheet.getLastColumn()).getValues();
  var directory_Data = source_Sheet.getDataRange().getValues();
  var numberColumn = source_Sheet.getLastColumn();

  // Khai báo và xử lý dữ liệu
  var data_LLV = [];
  const day = new Date();
  
  for (var row = 0; row < directory_Data.length; row++) {
    var n = SPX_IDs.indexOf(directory_Data[row][2].trim());
    if (n != -1) {
      // Logger.log(directory_Data[row][2]);
      for (var col = 0; col < numberColumn; col++) {
        if (formattedDate(directory_Data[1][col]) == formattedDate(day)) {
          data_LLV.push([SPX_IDs[n], directory_Data[row][col], directory_Data[row][col+1], directory_Data[row][col+2], directory_Data[row][col+3], directory_Data[row][col+4], directory_Data[row][col+5], directory_Data[row][col+6], directory_Data[row][col+7], directory_Data[row][col+8], directory_Data[row][col+9], directory_Data[row][col+10], directory_Data[row][col+11], directory_Data[row][col+12], directory_Data[row][col+13], directory_Data[row][col+14], directory_Data[row][col+15]]);
          break; 
        }
      }
    }    
  } 

  Logger.log(data_LLV);

  // Lấy id của data_LLV lưu vào một mảng riêng.
  var id_data_LLV = new Array();
  for (row = 0; row < data_LLV.length; row++) {
    id_data_LLV.push(data_LLV[row][0]);
  }

  // Tìm vị trí id nhân viên tại sheet ứng với vị trí id nhân viên tại gsheet nguồn và lưu tuần tự vào.
  var data_LLVs = [];
  for (i = 0; i < data_LLV.length; i++){
    var n = id_data_LLV.indexOf(SPX_IDs[i]);
    // Logger.log("Vị trí " + SPX_IDs[i] + " trong data_LLV là: " + n);
    try {
      data_LLVs.push([data_LLV[n][1], data_LLV[n][2], data_LLV[n][3], data_LLV[n][4], data_LLV[n][5], data_LLV[n][6], data_LLV[n][7], data_LLV[n][8], data_LLV[n][9], data_LLV[n][10], data_LLV[n][11], data_LLV[n][12], data_LLV[n][13], data_LLV[n][14]]);
    } catch {
      continue;
    } 
  }

  // Hiển thị mảng data_LLs vào sheet
  refreshSheet('1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A','LLV','H3:U');
  if (data_LLVs.length > 0) {
    destination_Sheet.getRange(3, 8, data_LLVs.length, data_LLVs[0].length).setValues(data_LLVs);
    destination_Sheet.getRange('A1').setValue('Last Update: ' + day.toLocaleDateString('vi-VN')+ ' ' + day.toLocaleTimeString('vi-VN'))
  } else {
    destination_Sheet.getRange("A1").setValue("Không truy vấn được dữ liệu");
  }
}