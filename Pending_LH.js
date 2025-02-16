// Lấy dữ liệu TO packing/packed

function TO_Management() {
  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Pending_LH");

  const day = new Date();
  day.setHours(23, 59, 59, 59);
  const get15Days = get15DaysAgo();
  get15Days.setHours(0, 0, 0, 0);  

  var start_time = Math.floor(timestamp(get15Days)/1000);
  var end_time = Math.floor(timestamp(day)/1000);
  var url_packing = "https://spx.shopee.vn/api/in-station/general_to/outbound/search?pageno=1&count=1000&status=1&ctime=" + start_time + "," + end_time; //TO packing
  var url_packed = "https://spx.shopee.vn/api/in-station/general_to/outbound/search?pageno=1&count=1000&status=2&ctime=" + start_time + "," + end_time; //TO packed

  var options = options_GET();

  // Tạo kiểu dữ liệu cần lấy
  function parserProduct(json) {
    var d = {};
    d['to_number'] = json.to_number;
    d['quantity'] = json.quantity;
    d['dest_station_name'] = json.dest_station_name;
    d['high_value'] = check_high_value(json.high_value);
    d['operator'] = json.operator;
    d['transfer_direction'] = json.transfer_direction;
    d['complete_time'] = convertTimestamp(json.complete_time);
    return d;
  }

  var result = [];

  // Send request
  try {
    var response = UrlFetchApp.fetch(url_packed, options);
    var response_packing = UrlFetchApp.fetch(url_packing, options);

    // check status
    if (response.getResponseCode() == 200) {
      
      Logger.log("url_packing: " + url_packing);
      Logger.log("url_packed: " + url_packed);

      var data = JSON.parse(response.getContentText());
      var data_packing = JSON.parse(response_packing.getContentText());

      var total_packing = data_packing.data.total;
      var total_packed = data.data.total;
      Logger.log("total_packing: " + total_packing);
      Logger.log("total_packed: " + total_packed);

      // continue data processing
      for (var i = 0; i < total_packed; i++) {
        var iJSON = data.data.list[i];        
        if (iJSON.current_station_id == 915) {
          result.push(parserProduct(iJSON));
        }
      }
      for (var j = 0; j < total_packing; j++) {
        var iJSON_packing = data_packing.data.list[j];
        if (iJSON_packing.current_station_id == 915) {
          result.push(parserProduct(iJSON_packing));
        }
      }

    } else {
      destination_Sheet.getRange("A1").setValue("PIC đã thay đổi khóa truy cập");
      return;
    }
  } catch (error) {
    destination_Sheet.getRange("A3").setValue("Lỗi Code - PIC Vũ Nguyễn");
  }
  
  var data = [];

  for (var i = 0; i < result.length; i++) {
    var row = [
      result[i].to_number,
      result[i].quantity,
      result[i].dest_station_name,
      result[i].high_value,
      result[i].operator,
      result[i].transfer_direction,
      result[i].complete_time
    ];
    data.push(row);
  }

  try {
    destination_Sheet.getRange("A3:G").clearContent();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "Done !", 3);
  }

  Logger.log("Dữ liệu thu được: " + data.length);
  const timenow = new Date();
  if (data.length > 0) {
    destination_Sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
    destination_Sheet.getRange("A1").setValue("Last update: " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN'));
  } else {
    destination_Sheet.getRange("A1").setValue("Không có TO chờ lên tải");
  }

}

// Lấy nhanh dữ liệu FMHub_Received hiện tại của Hub
function FMHub_Received() {
  // Sheet chứa data
  var destination_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var destination_Sheet = destination_Spreadsheet.getSheetByName("Pending_LH");

  var data = {
    "order_status":"42",
    "count":2000,
    "current_station_ids":"915",
    "page_no":1
  };
  const options = options_POST(data);

  // Dữ liệu cần lấy
  function parserProduct(json) {
    var d = {};
    d['shipment_id'] = json.shipment_id;
    d['sort_code_name'] = json.sort_code_name;
    d['high_value'] = check_high_value(json.high_value);
    d['next_station_name'] = json.next_station_name;
    return d;
  }

  var result_dataFM = [];
  
  var url_api = "https://spx.shopee.vn/api/fleet_order/order/tracking_list/search";
  var response = UrlFetchApp.fetch(url_api, options);
  var data_fmreceived = JSON.parse(response.getContentText());

  if(response.getResponseCode() == 200) {
    if (data_fmreceived != null) {
      Logger.log("url_api FMHub_Received: " + url_api);
      // Logger.log("Số trang:" + data_fmreceived.data.page_no);
      // Logger.log("Số đơn: " + data_fmreceived.data.list.length);
      for (var n = 0; n < data_fmreceived.data.list.length; n++) {
        result_dataFM.push(parserProduct(data_fmreceived.data.list[n]));
      }
    } else {
        Logger.log("Không dữ liệu FMHub_Received")
    }
  } else {
    Logger.log("getResponseCode: " + response.getResponseCode());
  }

  var show_data = [];

  for (var i = 0; i < result_dataFM.length; i++) {
    var row = [
      result_dataFM[i].shipment_id,
      result_dataFM[i].sort_code_name,
      result_dataFM[i].high_value,
      result_dataFM[i].next_station_name,
    ];
    show_data.push(row);
  }

  try {
    destination_Sheet.getRange("N3:Q").clearContent();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "Done !", 3);
  }

  Logger.log("Toltal FMHub_Received: " + show_data.length);
  const timenow = new Date();
  if (show_data.length > 0) {
    destination_Sheet.getRange(3, 14, show_data.length, show_data[0].length).setValues(show_data);
    destination_Sheet.getRange("N1").setValue("FMHub_Received, " + timenow.toLocaleDateString('vi-VN')+ ' ' + timenow.toLocaleTimeString('vi-VN') + ", total: " + show_data.length);
  } else {
    destination_Sheet.getRange("N3").setValue("Không có đơn FMHub_Received");
  }

}