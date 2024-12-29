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
        .addItem('TO_Packed: TO Packed chờ lên tải', 'TO_Packed')
        .addItem('Data_Rider: Danh sách rider', 'Data_Rider')   
      )
      
      .addSubMenu(ui.createMenu('Follow_Reverse')
        .addItem('Xóa dữ liệu cũ', 'refresh_Push_Reverse')
        .addItem('Cập nhật [pending_assign]', 'craw_pending_assign')
        .addItem('Check onhold time - pending project', 'craw_pending_assign')
      )

      .addSubMenu(ui.createMenu('Data_LH_Complete')
        .addItem("Xóa dữ liệu cũ", "refresh_Data_LH_Complete")
        .addItem('Export data', 'Data_LH_Complete'))
  .addToUi();
}


function onOpen() {
  createMenu();
}