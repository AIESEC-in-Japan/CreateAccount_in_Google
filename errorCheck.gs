function errorCheck () {

  var ui = SpreadsheetApp.getUi()
  var response = ui.prompt('何行目までチェックを行いますか?')

  //var ROW_END = 665
  var ROW_END = Number(response.getResponseText())
  if (typeof(ROW_END) === "number") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    var lastRow = sheet.getLastRow()
    var lastColumn  = 5

    for (var i = lastRow; i >= ROW_END; i--) {
      var rowRange = sheet.getRange(i, 4, 1, 2)
      var rowData = rowRange.getValues()[0]
      var userKey = rowData[0] + '.' + rowData[1] + '@aiesec.jp'
      var user
      Logger.log(userKey)

      try {
        user = AdminDirectory.Users.get(userKey)
        rowRange.setBackground('green')
      } catch (e) {
        Logger.log(userKey)
        rowRange.setBackground('orange')
      }
    }
  }
}


function bulkReinsert () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var range = sheet.getDataRange()
  var data = range.getValues()

  data.forEach(function (rowData) {
    try {
      insertUser(rowData)
    } catch (e) {
      Logger.log(e)
      Logger.log(rowData[4] + rowData[5])
    }
  })
}
