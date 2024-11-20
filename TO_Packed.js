function TO_Packed() {
  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("TO_Packed");

  const day = new Date();
  day.setHours(23, 59, 59, 59);
  const get15Days = get15DaysAgo();
  get15Days.setHours(0, 0, 0, 0);  

  var start_time = Math.floor(timestamp(get15Days)/1000);
  var end_time = Math.floor(timestamp(day)/1000);
  var url = "https://spx.shopee.vn/api/in-station/general_to/outbound/search?pageno=1&count=1000&status=2&ctime=" + start_time + "," + end_time;
  Logger.log(url);

  var infor_Sheet = destination_Spreadsheet.getSheetByName("README");
  fms_user_id = infor_Sheet.getRange("H3").getValue();
  fms_user_skey = infor_Sheet.getRange("H2").getValue();
  fms_display_name = infor_Sheet.getRange("H1").getValue();

  var options = {
    "method": "get",
    "headers": {
    "Cookie": "fms_user_id=" + fms_user_id + "; fms_user_skey=" + fms_user_skey + "; fms_display_name=" + fms_display_name + "; spx_st = 1; spx_cid = VN"
    },
    "muteHttpExceptions": true
  };

  // Tạo kiểu dữ liệu cần lấy
  function parserProduct(json) {
    var d = {};
    d['to_number'] = json.to_number;
    d['quantity'] = json.quantity;
    d['dest_station_name'] = json.dest_station_name;
    d['high_value'] = json.high_value;
    d['operator'] = json.operator;
    d['transfer_direction'] = json.transfer_direction;
    return d;
  }

  var result = [];

  // Send request
  try {
    var response = UrlFetchApp.fetch(url, options);

    // check status
    if (response.getResponseCode() == 200) {
      var data = JSON.parse(response.getContentText());
      var lengthData = 200;    

      // continue data processing
      for (var i = 0; i < lengthData; i++) {
        var iJSON = data.data.list[i];
        if (iJSON.current_station_id == 915 || iJSON.current_station_name == "50-HCM Tan Binh/Bach Dang Hub") {
          result.push(parserProduct(iJSON));
        }        
      } 

    } else {
      destination_Sheet.getRange("A1").setValue("fms_user_skey đã thay đổi");
    }
  } catch (error) {
    destination_Sheet.getRange("A1").setValue("Lỗi Code");
  }

  var data = [];

  for (var i = 0; i < result.length; i++) {
    var row = [
      result[i].to_number,
      result[i].quantity,
      result[i].dest_station_name,
      result[i].high_value,
      result[i].operator,
      result[i].transfer_direction
    ];
    data.push(row);
  }

  try {
    destination_Sheet.getRange("A3:F").clearContent();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "Done !", 3);
  }

  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("fms_user_skey đã thay đổi");
  }
}
