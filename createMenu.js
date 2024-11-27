function createMenu() {
  const ui = SpreadsheetApp.getUi()
    ui.createMenu('PIC Tool')
      .addSubMenu(ui.createMenu('LLV')
        .addItem('Cập nhật dữ liệu', 'LLV'))
      .addSubMenu(ui.createMenu('Follow_OS&BPO')
        .addItem('Cập nhật PlanOS', 'PlanOS'))
      .addSubMenu(ui.createMenu('Follow_Reverse')
        .addItem('Xóa dữ liệu cũ', 'refresh_Push_Reverse')
        .addItem('Cập nhật [pending_assign]', 'craw_pending_assign'))
      .addSubMenu(ui.createMenu('Before_Lost')
        .addItem('Cập nhật dữ liệu', 'Before_Lost'))      
      .addSubMenu(ui.createMenu('Check_Overdue')
        .addItem('Cập nhật dữ liệu', 'Check_Overdue'))
      .addSubMenu(ui.createMenu('Check_OrderSPX')
        .addItem('Cập nhật dữ liệu', 'Check_OrderSPX'))
      .addSubMenu(ui.createMenu('Check_Order247')
        .addItem('Cập nhật dữ liệu', 'Check_Order247'))
      .addSubMenu(ui.createMenu('Plan_Truck')
        .addItem('Cập nhật dữ liệu', 'Plan_Truck'))
      .addSubMenu(ui.createMenu('TO_Packed')
        .addItem('Cập nhật dữ liệu', 'TO_Packed'))
      .addSubMenu(ui.createMenu('Data_Rider')
        .addItem('Cập nhật dữ liệu', 'Data_Rider'))
  .addToUi();
}

function onOpen() {
  createMenu();
}