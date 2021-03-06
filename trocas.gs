function num(num1) // essa função vai coletar o codigo de barras, porem convertido em string para validação
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var num1 = sheet.getRange("A3").getValues().toString()
  return num1;
  
}

function valor()
{
  var ss = ss = SpreadsheetApp.getActiveSpreadsheet();
  var valor1 = ss.getRange("C3").getValue();
  return valor1;
  
}
function name() 
{
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var name1 = ss.getRange("B3").getValues().toString();
  return name1;
}
function date(){
  var formattedDate = Utilities.formatDate(new Date(), "GMT-3", "dd-MM-yyyy HH:mm ")
  return formattedDate

}
function append() // essa função contem toda a regra de negocio 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange("A5:A100").getValues(); // aqui vamos pegar um range para comparação com a celula A3
  var str = range.toString(); // convertemos esse o range para string para poder fazer a validação 
  var range1 = [{}]
  range1 = str.split(",");
  var check = range1.indexOf(num()); // aqui vamos checar qual o index do valor da celula A3 caso ela nao exista no index o valor a ser retornado e -1
  var cellempty = SpreadsheetApp.getActiveSheet().getRange("C3").isBlank() // checa se existe algo digitado na celula C3
  var nomEmpty = SpreadsheetApp.getActiveSheet().getRange("A3").isBlank() // repete o mesmo processo 
  if(cellempty == true ) // caso o valor de C3 seja null  vai impedir que seja adicionado algo na planilha
  {
    var a = SpreadsheetApp.getUi().alert("A quantidade não pode estar vazia"); 
  }
  else if(range1.indexOf(num()) != -1 && name() != "NAO CADASTRADO" && nomEmpty == false) // essa parte e responsavel pela validação caso o valor de A3 seja encontrado no range
  {
          var a = SpreadsheetApp.getUi();
          var cellrange = SpreadsheetApp.getActiveSheet().getRange(5 + check, 3).getA1Notation();// essa formula vai pegar o valor de check que é o index do valor de A3
          var cellvalue = SpreadsheetApp.getActiveSheet().getRange(cellrange).getValue();        //  e somar mais 5(porquê os codigos são adicionados a partir de A5)
          var sheet = SpreadsheetApp.getActiveSheet().getRange(cellrange);
          var newvalue = cellvalue + valor()
          var cellDate = SpreadsheetApp.getActiveSheet().getRange(5 + check, 4).getA1Notation();
          var sheet1 = SpreadsheetApp.getActiveSheet().getRange(cellDate);
          var yesno = SpreadsheetApp.getUi()
          var response = yesno.alert("Você esta adiconando o produto: " + name() + " Quantidade: " + valor() + "\nCONFIRMA?",yesno.ButtonSet.YES_NO)
          if(response == yesno.Button.YES)
          {
          sheet.setValue(cellvalue + valor()) // aqui vai somar antigo com a nova quantidade
          sheet1.setValue(date()) // aqui vai somar antigo com a nova quantidade
          
            
    a.alert("Produto ja adicionado: " + name() + " voce adicionou: " + valor() + " o valor foi alterado de: " + cellvalue + " para: " + newvalue);
          var c = SpreadsheetApp.getActiveSheet().getRange("A3").clearContent();
          var b = SpreadsheetApp.getActiveSheet().getRange("C3").clearContent();
          var d = SpreadsheetApp.getActiveSheet().setActiveSelection("A3")
          }
    else
    {
      yesno.alert("Você cancelou a operação");
    }
  }
  else if(nomEmpty == true){
    var a = SpreadsheetApp.getUi().alert("Você precisa digitar um codigo de barras");
  }
  else if(name() == "NAO CADASTRADO"){
    var a = SpreadsheetApp.getUi().alert("CODIGO DE BARRAS NÃO EXISTE");
  }
  
  
  else{
     var yesno = SpreadsheetApp.getUi()
          var response = yesno.alert("Você esta adiconando o produto: " + name() + " Quantidade: " + valor() + "\nCONFIRMA?",yesno.ButtonSet.YES_NO)
          if(response == yesno.Button.YES){
    var add = ss.appendRow([num(),name(),valor(),date()]) // caso não existe o codigo no index, ele vai adicionar na proxima linha livre todas as informações 
    var a = SpreadsheetApp.getUi().alert("O produto adicionado: " + name() + " Quantidade: " + valor());
    var c = SpreadsheetApp.getActiveSheet().getRange("A3").clearContent();
    var b = SpreadsheetApp.getActiveSheet().getRange("C3").clearContent();
    var d = SpreadsheetApp.getActiveSheet().setActiveSelection("A3")
          }
    else
    {
      yesno.alert("Você cancelou a operação");
    }
}
 
    }

  

function buscar() {


    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var selection = SpreadsheetApp.getActiveSpreadsheet()
    var sheetnum = SpreadsheetApp.getActiveSpreadsheet().getSheets()

    var i = 0
    var num = SpreadsheetApp.getActiveSpreadsheet().getRange("C2").getValues().toString()


    for (var i = 1; i < sheetnum.length; i++) {

        selection.setActiveSheet(ss.getSheets()[i])
        var ui = SpreadsheetApp.getUi()
        var sss = SpreadsheetApp.getActiveSpreadsheet();
        var range = sss.getRange("A1:A100").getValues(); // aqui vamos pegar um range para comparação com a celula A3
        var str = range.toString(); // convertemos esse o range para string para poder fazer a validação 
        var range1 = [{}]
        range1 = str.split(",");

        var check = range1.indexOf(num)
        var namesheet = selection.getActiveSheet().getName()
        if (range1.indexOf(num) != -1) {
            var qnt = SpreadsheetApp.getActiveSheet().getRange(check + 1, 3).getValue()
            var desc = SpreadsheetApp.getActiveSheet().getRange(check + 1, 2).getValue()
            //ui.alert('Achei uma igual na planilha ' + namesheet + ' na celula ' + (1 + check) + ' Quantidade: ' + qnt )
            if (range1.indexOf(num) > 0) {
              selection.setActiveSheet(ss.getSheets()[0]) //.getRange("D2").setValue(desc)
              selection.appendRow(['', desc, qnt, namesheet, check + 1])
              
                
                //selection.getRange("E2").setValue(qnt)
                //selection.getRange("F2").setValue(namesheet)
            }
            selection.getRange("G2").setValue(check + 1)
            
        }
        else {

        }
        selection.setActiveSheet(ss.getSheets()[i])
    }


} 

function freeze(){
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let range = ss.getRange("A1:E2").canEdit(0)
}
function sortSheets () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = [];
  var sheets = ss.getSheets();
   
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  
  sheetNameArray.sort();
    
  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
}
