function diaFolga() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getRange("D9").getValue()
  return ss
}
function folga() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getRange("B12:B42").getValues()
  var sheet = SpreadsheetApp.getActiveSheet()
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var range = ss.toString()
  var range1 = [{}]
  range1 = range.split(",")
  var first = range1.indexOf(diaFolga())
  var check = first
  while (check < range1.length) {
    var obj = [["XXX", "XXX", "XXX", "XXX", "FOLGA"]]
    var color = sheet.getRange(check + 12, 1, 1, 7).setBackgroundRGB(165, 165, 165)
    var newValue = sheet.getRange(check + 12, 3, 1, 5).setValues(obj)
    check += 7
  }
  sheet1.setActiveSheet(sheet1.getSheets()[0])
}

function getMounth() {
  let ui = SpreadsheetApp.getUi()
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let activeSheet = sheet.getActiveSheet()
  let g9 = sheet.getActiveSheet().getRange("G9").getValue()
  switch (g9) {
    case 4:
    case 6:
    case 9: 
    case 11:
      activeSheet.showRows(40, 3)
      activeSheet.hideRows(42, 1)
      break;
    case 2:
      activeSheet.hideRows(40, 3)
      break
    default:
      activeSheet.showRows(40, 3)
      break
  }



  /*  if(sheet.getActiveSheet().getRange("G9").getValue() == 4 || 6 || 9 || 11) {
     activeSheet.showRows(40, 3)
     activeSheet.hideRows(42, 6)
 
   } else if (sheet.getActiveSheet().getRange("G9").getValue() === 1 || 3 || 5 || 7 || 8 || 10 || 12) {
     activeSheet.showRows(40, 3)
   } if (sheet.getActiveSheet().getRange("G9").getValue() === 2) {
     activeSheet.hideRows(40, 3)
   } */
}

function getSaturday(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let days = SpreadsheetApp.getActiveSheet().getRange("B12:B42").getValues()
  let ui = SpreadsheetApp.getUi()
  let parsedDays = days.toString()
  let newDays = [{}]
  newDays = parsedDays.split(",")
  let first = newDays.indexOf("SÁBADO")
  let check = first
  let saturday = sheet.getActiveSheet().getRange("A9").getValue()
  while(check < newDays.length){
  let merge = sheet.getActiveSheet().getRange(check + 12, 4, 1, 2)
  if(saturday == "08:00 AS 12:00" ){
    merge.mergeAcross()
    merge.setBackgroundRGB(165, 165, 165)
    merge.setValue("SÁBADO")
  }
   check += 7
  }
}

function umerge(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let days = SpreadsheetApp.getActiveSheet().getRange("B12:B42").getValues()
  let ui = SpreadsheetApp.getUi()
  let parsedDays = days.toString()
  let newDays = [{}]
  newDays = parsedDays.split(",")
  let first = newDays.indexOf("SÁBADO")
  let check = first
  let saturday = sheet.getActiveSheet().getRange("A9").getValue()
  while(check < newDays.length){
  let merge = sheet.getActiveSheet().getRange(check + 12, 4, 1, 2)
  if(saturday == "08:00 AS 12:00" ){
    merge.breakApart()
    merge.setBorder(true,true,true,true,true,true)
  }
   check += 7
  }
}

function isAol(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let ui = SpreadsheetApp.getUi()
  let org = sheet.getActiveSheet().getRange("a3").getValue().toString().substring(0,3)
  
  ui.alert(org)
}

function folhas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var selection = SpreadsheetApp.getActiveSpreadsheet()
  var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i = 1; i < sheetnum.length; i++) {
    selection.setActiveSheet(ss.getSheets()[i])
    getMounth()
    getSaturday()
    folga()
  }

}

function isMtb(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let ui = SpreadsheetApp.getUi()
    if(sheet.getActiveSheet().getSheetName() == "MTB"){
      return true
    }
}
function clearAll() {
  let ui = SpreadsheetApp.getUi()
  var sss = SpreadsheetApp.getActiveSpreadsheet()
  var selection = SpreadsheetApp.getActiveSpreadsheet()
  var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  let response = ui.alert("Tem certeza que deseja remover as folgas?", ui.ButtonSet.YES_NO)
  
  if (response == ui.Button.YES) {
    for (var i = 1; i < sheetnum.length; i++){
      selection.setActiveSheet(sss.getSheets()[i])
      if(isMtb() != true){
      umerge()
      var ss = SpreadsheetApp.getActiveSheet().getRange(12, 3, 31, 5).clearContent()
      var ss = SpreadsheetApp.getActiveSheet().getRange(12, 1, 31, 7).setBackground("white")
      }
    }
    sss.setActiveSheet(sss.getSheets()[0])
  }
  else {
    ui.alert("Você cancelou a operação!")
  }
}

function teste(){
  let ui = SpreadsheetApp.getUi()
  if(isMtb() == true) ui.alert("True")
  else ui.alert("False")
}
function duplicate() {
  var i = 0
  while (i < 15) {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    ss.duplicateActiveSheet()
    i++
  }
}
function rename() {
  var sss = SpreadsheetApp.getActiveSpreadsheet()
  var selection = SpreadsheetApp.getActiveSpreadsheet()
  var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i = 1; i < sheetnum.length; i++) {
    selection.setActiveSheet(sss.getSheets()[i])
    var funcionario = sss.setActiveSheet(sss.getSheets()[i]).getRange(5, 1).getValue().toString().substr(11, Number.MAX_VALUE) //SpreadsheetApp.getActiveSpreadsheet().getRange("A5").getValue()
    // @ts-ignore
    var rename = SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet(funcionario + " " + i)
  }
}

function remove() {
  var sss = SpreadsheetApp.getActiveSpreadsheet()
  var selection = SpreadsheetApp.getActiveSpreadsheet()
  var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i = 3; i < sheetnum.length; i++) {
    selection.setActiveSheet(sss.getSheets()[i])
    var funcionario = SpreadsheetApp.getActiveSpreadsheet().getRange("A5").getValue()
    var rename = SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet()
  }
}

function getSunday() {
  let ui = SpreadsheetApp.getUi()
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let name = sheet.getRange("A5").getValue()
  let days = sheet.getRange("B12:B42").getValues()
  let range = days.toString()
  let rangeDays = [{}]
  rangeDays = range.split(",")
  let isSunday = rangeDays.indexOf("DOMINGO")
  let check = isSunday
  let folgas = []
  let folgas1 = folgas.toString()

  while (check < rangeDays.length) {
    let folga = sheet.getActiveSheet().getRange(check + 12, 1, 1, 7)
    if (folga.getBackground() == "#d3d3d3") {
      folgas.push(folga.getValue())
    }
    check += 7
  }
  sheet.setActiveSheet(sheet.getSheets()[0])
  let i = 0
  let obj = [name, [folgas1]]

  let append = sheet.getActiveSheet().getRange(3 + i, 9, 1, 2)
  let blankCell = append.isBlank()
  ui.alert(blankCell + " " + folgas + " " + name)
  if (blankCell == false) {
    i += 1
    ui.alert(i)
  }
  append.setValues([obj])

}

function summary() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let ui = SpreadsheetApp.getUi()
  let activeSheet = sheet.getActiveSheet()
  let url = SpreadsheetApp.getActiveSpreadsheet().getUrl()
  for (let j = 1; j < sheet.getSheets().length; j++) {
    let getNameFunc = sheet.setActiveSheet(sheet.getSheets()[j]).getRange(5, 1).getValue().toString().substr(11, Number.MAX_VALUE)
    let getFolgaName = sheet.setActiveSheet(sheet.getSheets()[j]).getRange(9, 4).getValue().toString()
    let getFolga = sheet.setActiveSheet(sheet.getSheets()[j]).getRange(9, 4).getA1Notation()
    let sheetName = sheet.getSheets()[j].getName()
    let sheetId = sheet.setActiveSheet(sheet.getSheets()[j]).getSheetId()
    let linkFolga = SpreadsheetApp.newRichTextValue()
      .setText(getFolgaName)
      .setLinkUrl(`${url}#gid=${sheetId}`)
      .build()

    sheet.setActiveSheet(sheet.getSheets()[0])

    for (let i = 0; i < 120; i++) {
      let isEmpty = activeSheet.getRange(2 + i, 7).isBlank()
      if (isEmpty == true) {
        activeSheet.getRange(2 + i, 7).setValue(getNameFunc)
        //activeSheet.getRange(2 + i, 8).setFormula(`=('${sheetName}'!${getFolga})`)
        activeSheet.getRange(2 + i, 8).setRichTextValue(linkFolga)



        break
      }
    }
  }

}

function removeSheet() {
  var sss = SpreadsheetApp.getActiveSpreadsheet()
  var selection = SpreadsheetApp.getActiveSpreadsheet()
  var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i = 38; i < sheetnum.length; i++) {
    selection.deleteSheet(sss.getSheets()[i])

  }
}

function burro() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  for (let j = 1; j < sheet.getSheets().length; j++) {
    sheet.setActiveSheet(sheet.getSheets()[j]).getRange("G9").setFormula("=(Inicia!A3)")
  }
}




