function refresh_Data_LH_Complete() {
  refreshSheet("1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A","Data_LH_Complete","A2:G");
}

function Data_LH_Complete() {

  // Lấy id tất cả LT lưu vào một mảng.
  var myAppSpr = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
  var mySheet = myAppSpr.getSheetByName("Data_LH_Complete");
  var value_mySheet = mySheet.getRange("A2:A").getValues();
  const options = options_API();

  var id_LT = new Array();
  var url1 = "https://spx.shopee.vn/api/admin/transportation/trip/history/list?trip_number=";
  var url2 = "https://spx.shopee.vn/api/admin/transportation/trip/list?trip_number=";
  for (var i = 0; i < value_mySheet.length; i++) {
    if(value_mySheet[i][0] != "") {
      var response = UrlFetchApp.fetch(url1 + value_mySheet[i][0], options);
      var data = JSON.parse(response.getContentText());
      var check_data = data.data.total;
      switch (check_data != 0) {
        case true:
          var iJSON = data.data.list[0];
          id_LT.push([iJSON.id, iJSON.trip_number]);
          break;
        case false:
          response = UrlFetchApp.fetch(url2 + value_mySheet[i][0], options);
          data = JSON.parse(response.getContentText());
          var iJSON = data.data.list[0];
          id_LT.push([iJSON.id, iJSON.trip_number]);
          break;          
      };
    }
  }


  // Khai báo thông tin cần lấy
  function parser_LT(LT, json_loading) {
    var d = {};
    d['LH_Trip'] = LT;
    d['to_number'] = json_loading.to_number;
    d['scan_number'] = json_loading.scan_number;
    d['to_parcel_quantity'] = json_loading.to_parcel_quantity;
    d['unloaded_station_name'] = json_loading.unloaded_station_name;
    return d;
  }

  // Lấy dữ liệu theo từng id_LT
  var url_get_dataLT = "https://spx.shopee.vn/api/admin/transportation/trip/loading/list?trip_id=";
  var data_LH = new Array();
  var IDs = 0;
  do{
    var link = url_get_dataLT + id_LT[IDs][0] + "&pageno=1&count=5000&loaded_sequence_number=1&type=outbound";
    var response_data = UrlFetchApp.fetch(link, options);
    var data_IDs = JSON.parse(response_data.getContentText());
    if(response_data.getResponseCode() == 200){
      var IDsJSON = data_IDs.data.list;
    } else {
      Logger.log("API call failed id: " + id_LT[IDs][0]);
    }  
    for (var j = 0; j < IDsJSON.length; j++) {
      data_LH.push(parser_LT(id_LT[IDs][1], IDsJSON[j]));

    }
    IDs++;
  }while(IDs < id_LT.length);


  // Sắp xếp lại dữ liệu theo thứ tự mong muốn
  var data = [];
  for (var i = 0; i < data_LH.length; i++) {
    var row = [
      data_LH[i].LH_Trip,
      data_LH[i].to_number,
      data_LH[i].scan_number,
      data_LH[i].to_parcel_quantity,
      data_LH[i].unloaded_station_name
    ];
    data.push(row);
  }


  try {
    mySheet.getRange("C2:G").clearContent();
  }catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "DONE!", 3);
  }

  if (data.length > 0) {
    mySheet.getRange(2, 3, data.length, data[0].length).setValues(data);
  } else {
    mySheet.getRange("C2").setValue("Truy vấn dữ liệu không thành công");
  }
}
