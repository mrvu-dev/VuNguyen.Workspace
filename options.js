// Chứa thông tin khi thực hiện truy vấn
function options_API() {
    var mySpr = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-YHC2Nvv9s97CfB2ShgKwO9J2xMFrZ-PHR_t5IiV8-A/edit");
    var mySheet = mySpr.getSheetByName("README");
  
    fms_user_id = mySheet.getRange("C3").getValue();
    fms_user_skey = mySheet.getRange("C2").getValue();
    fms_display_name = mySheet.getRange("C1").getValue();
  
    var options = {
      "method": "get",
      "headers": {
        "Cookie": "fms_user_id=" + fms_user_id + "; fms_user_skey=" + fms_user_skey + "; fms_display_name=" + fms_display_name + "; spx_st = 1; spx_cid = VN"
      },
      "muteHttpExceptions": true
    };
    return options;
  }
  