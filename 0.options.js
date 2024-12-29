// Chứa thông tin khi thực hiện truy vấn
function options_API() {
  var mySpr = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cl8dsvV923vkJelkGHP9y2P860O1ZC3jvaYsqBxKHpk/edit");
  var mySheet = mySpr.getSheetByName("API");

  fms_user_id = mySheet.getRange("B3").getValue();
  fms_user_skey = mySheet.getRange("B2").getValue();
  fms_display_name = mySheet.getRange("B1").getValue();

  var options = {
    "method": "get",
    "headers": {
      "Cookie": "fms_user_id=" + fms_user_id + "; fms_user_skey=" + fms_user_skey + "; fms_display_name=" + fms_display_name + "; spx_st = 1; spx_cid = VN"
    },
    "muteHttpExceptions": true
  };
  return options;
}
