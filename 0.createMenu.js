function createMenu() {
  const ui = SpreadsheetApp.getUi()
    ui.createMenu('PIC Tool')
      .addSubMenu(ui.createMenu('Cập nhật dữ liệu')
        .addItem('LLV: Lịch làm việc', 'LLV')
        .addItem('PlanOS: Kế hoạch booking OS', 'PlanOS')
        .addItem('Before_Lost: Phản hồi hành trình đơn hàng', 'Before_Lost')
        .addItem('Check_Overdue: Kiểm tra đơn quá hạn', 'Check_Overdue')
        .addItem('Check_OrderSPX: Kiểm tra đơn nội bộ SPX', 'Check_OrderSPX')
        .addItem('Check_Order247: Kiểm tra đơn nội bộ 247', 'Check_Order247')
        .addItem('Plan_Truck: Lịch xe tải lấy hàng', 'Plan_Truck')
        .addItem('Data_Rider: Danh sách rider', 'Data_Rider')
      )

      .addSubMenu(ui.createMenu('Pending_LH')
        .addItem('TO Pending LH: Cập nhật list TO', 'TO_Management')
        .addItem('FMHub_Received: Cập nhật list FMHub_Received', 'FMHub_Received')
        // .addItem('Update data Sheet: Cập nhật toàn bộ data sheet', 'Pending_LH')
      )
      
      .addSubMenu(ui.createMenu('Follow_Reverse')
        .addItem('Xóa dữ liệu cũ', 'refresh_Push_Reverse')
        .addItem('Cập nhật pending_assign', 'craw_pending_assign')
      )

      .addSubMenu(ui.createMenu('Data_LH')
        .addItem("Xóa dữ liệu cũ", "refresh_Data_LH")
        .addItem('Export data', 'Data_LH'))
  .addToUi();
}


function onOpen() {
  createMenu();
}