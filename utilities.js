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

// Định dạng lại ngày tháng.
function formattedDate(time) {
  time = new Date(time);
  var day = time.getDate();
  var month = time.getMonth() + 1; //Tháng trong JS bắt đầu từ 0
  var year = time.getFullYear();
  time = `${day}/${month}/${year}`;
  return time;
}

// Lấy ngày.
function layngay(time) {
  var getString = formattedDate(time)
  const kq = parseInt(getString.slice(0, 2));
  return kq;
}

// Chuyển đổi date sang timestamp.
function timestamp(time) {
  time = new Date(time);
  time.getTime();
  return time;
}

// Lấy ngày đầu tiên của tuần hiện tại.
function getFirstDayOfWeek(date = new Date()) {
  const dayOfWeek = date.getDay(); // Lấy thứ trong tuần (0 - Chủ Nhật, 1 - Thứ Hai, ..., 6 - Thứ Bảy)
  const diff = date.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1); // Tính số ngày cần trừ để đến thứ Hai
  return new Date(date.setDate(diff));
}

// Lấy 15 ngày trước kể từ ngày hiện tại
function get15DaysAgo() {
    const today = new Date();
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(today.getDate() - 15);
    return thirtyDaysAgo;
}
