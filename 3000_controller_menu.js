function onOpen() {
  const ui = SpreadsheetApp.getUi();
  buildFmMenu(ui);
  buildCommonMenu(ui);
}