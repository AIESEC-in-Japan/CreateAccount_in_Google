function onOpen() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var menu = [
    {name: "アカウントを作成(最後の1行のみ)",functionName: "main"},
    {name: "アカウントを作成(現在のシートの全ての列)",functionName: "bulkReinsert"},
    {name: "アカウントの存在をチェック(最後から指定した行まで)",functionName: "errorCheck"}
  ];

  sheet.addMenu("CreateAccount", menu);
}
