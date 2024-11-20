// Kiểm tra kế hoạch điều xe lấy hàng Bạch Đằng - SOC

function Plan_Truck() {
  // Sheet lấy data
  var source_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1K-7Rrr_2TjCQmOeGAkys3Pl51VGrtExqD94ciqixGhM/edit");

  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Plan_Truck")
  hubName = "Tan Binh/Bach Dang";


  // --------------5. Plan Daily_HCM SOC--------------
  var source_Sheet_HCM = source_Spreadsheet.getSheetByName("5. Plan Daily_HCM SOC");
  var source_data_HCM = source_Sheet_HCM.getRange("B4:J").getValues();
  var data_HCM =[];
  data_HCM.push(source_data_HCM[0]);

  for(var i = 0; i < source_data_HCM.length; i++) {
    if(source_data_HCM[i][6].trim() == hubName || source_data_HCM[i][7].trim() == hubName) {
      data_HCM.push(source_data_HCM[i]);
    }
  }

  // --------------6. Plan Daily_BINH DUONG SOC--------------
  var source_Sheet_BD = source_Spreadsheet.getSheetByName("6. Plan Daily_BINH DUONG SOC");
  var source_data_BD = source_Sheet_BD.getRange("B3:J").getValues();
  var data_BD =[];

  for(var i = 0; i < source_data_BD.length; i++) {
    if(source_data_BD[i][6].trim() == hubName || source_data_BD[i][7].trim() == hubName) {
      data_BD.push(source_data_BD[i]);
    }
  }

  // --------------7. Plan Daily_SW SOC--------------
  var source_Sheet_SW = source_Spreadsheet.getSheetByName("7. Plan Daily_SW SOC");
  var source_data_SW = source_Sheet_SW.getRange("B3:J").getValues();
  var data_SW =[];

  for(var i = 0; i < source_data_SW.length; i++) {
    if(source_data_SW[i][6].trim() == hubName || source_data_SW[i][7].trim() == hubName) {
      data_SW.push(source_data_SW[i]);
    }
  }


  // --------------Show data--------------
  var data = new Array();
  for(row = 0; row < data_HCM.length; row++) {
    data.push(data_HCM[row]);
  }
  for(row = 0; row < data_BD.length; row++) {
    data.push(data_BD[row]);
  }
  for(row = 0; row < data_SW.length; row++) {
    data.push(data_SW[row]);
  } 

  try {    
    destination_Sheet.getRange("B3:J" + destination_Sheet.getLastRow()).clearContent();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "Done !", 3);
  }


  const timeNow = new Date();
  if (data.length > 1) {    
    destination_Sheet.getRange("A1").setValue("Last update: " + timeNow.toLocaleDateString('vi-VN')+ ' ' + timeNow.toLocaleTimeString('vi-VN'));
    destination_Sheet.getRange(2, 2, data.length, data[0].length).setValues(data);
  } else {
    destination_Sheet.getRange("A1").setValue("Team LH chưa cập nhật lịch mới, check link nguồn");
  }
  
}
