// Xóa nhanh dữ liệu.
function refreshSheet(sheetId, sheetName, range) {
  var myApp = SpreadsheetApp.openById(sheetId);
  var mySheet = myApp.getSheetByName(sheetName);

  try {
    mySheet.getRange(range + mySheet.getLastRow()).clearContent();
  } catch(error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Không có dữ liệu để xóa", "Done !", 3);
  }
}

// Cộng thêm ngày.
function addDays(date, days) {
  const newDate = new Date(date); // Tạo một bản sao của ngày gốc
  newDate.setDate(date.getDate() + days); // Cộng thêm số ngày
  return newDate;
}

function formattedDate(time) {
  time = new Date(time);
  var day = time.getDate();
  var month = time.getMonth() + 1; //Tháng trong JS bắt đầu từ 0
  var year = time.getFullYear();
  time = `${day}/${month}/${year}`;
  return time;
}

function layngay(time) {
  var getString = formattedDate(time)
  const kq = parseInt(getString.slice(0, 2));
  return kq;
}


function checkCode() {
  const day = new Date();
  today = addDays(day,8);
  Logger.log("Ngay: " + today.toLocaleDateString('vi-VN')+ ' ' + today.toLocaleTimeString('vi-VN'));
}
