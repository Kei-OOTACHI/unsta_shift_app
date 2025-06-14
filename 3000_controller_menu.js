/**
 * スプレッドシートを開いたときに実行される関数
 * メニューを構築し、UIに追加します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("🤖自動化ツール");
  
  // サブメニューを追加
  const subMenus = [
    buildFmMenu(ui),
    buildGanttMenu(ui),
    buildMemberMenu(ui),
    buildShiftDataMergerMenu(ui),
    buildCommonMenu(ui)
  ];
  
  subMenus.forEach(subMenu => menu.addSubMenu(subMenu));
  menu.addToUi();
}