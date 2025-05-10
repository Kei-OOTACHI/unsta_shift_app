function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🤖自動化ツール")
    .addSubMenu(buildFmMenu(ui))
    .addSubMenu(buildGanttMenu(ui))
    .addSubMenu(buildMemberMenu(ui))
    .addSubMenu(buildCommonMenu(ui))
    .addToUi();
}