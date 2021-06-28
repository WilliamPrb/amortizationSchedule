 /* 
Sistemas de amortização ESPM
William Perboni
*/


function buttonExe() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
var amortization = ss.getSheetByName("Amortizacao");
var operar = amortization.getRange("B6").getValue();


switch (operar) {
  case "Misto":
    Misto();
    break;
  case "SAC":
    TableSAC()
    break;
  case "Price":
    TablePrice()
    break;
}

 }





function Misto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var amortization = ss.getSheetByName("Amortizacao");
  var interest = amortization.getRange("D6").getValue();
  var time = amortization.getRange("E6").getValue();
  var value = amortization.getRange("C6").getValue();
  amortization.getRange("B15:H180").clear();
  amortization.getRange("B10:D11").setValue("Tabela Mista");

  var finalAmortization = value / time;
  var balance = value; 
  var k = (interest*((1+interest)**time))/(((1+interest)**time)-1); 
  var principal = value * k; 
  var valorMisto = value; 


  for(var i=1; i<=time; i++) {
    
    var nextRow = amortization.getLastRow() +1; 

    // Juros 
    var finalInterest = balance * interest; 
    var jurosMisto = valorMisto * interest;
    amortization.getRange(nextRow,3).setValue(jurosMisto);

      // Amortização
    var amortizationFinal = finalAmortization;
     //Pagamento 
    var sistemaSac = ((1/2)*(principal+amortizationFinal+finalInterest));
    amortization.getRange(nextRow,5).setValue(sistemaSac);

    // Amortização 2

    var amortizationMisto = sistemaSac - jurosMisto; 
    amortization.getRange(nextRow,4).setValue(amortizationMisto); 

  
   


  
     //Saldo Devedor 
    balance = balance - amortizationFinal;
    valorMisto = valorMisto - amortizationMisto;
    amortization.getRange(nextRow,6).setValue(valorMisto);

    amortization.getRange(nextRow,2).setValue(i);
      

   }

}


function TableSAC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var amortization = ss.getSheetByName("Amortizacao");
  var interest = amortization.getRange("D6").getValue();
  var time = amortization.getRange("E6").getValue();
  var value = amortization.getRange("C6").getValue();
  amortization.getRange("B15:H180").clear();
  amortization.getRange("B10:D11").setValue("Tabela SAC");

  var finalAmortization = value / time;
  var balance = value; 

  for(var i=1; i<=time; i++) {
    
    var nextRow = amortization.getLastRow() +1; 

    // Juros 
    var finalInterest = balance * interest; 
    amortization.getRange(nextRow,3).setValue(finalInterest);

      // Amortização
    var amortizationFinal = finalAmortization;
    amortization.getRange(nextRow,4).setValue(amortizationFinal); 

  
    //Pagamento 
  amortization.getRange(nextRow,5).setValue(amortizationFinal+finalInterest);


  
     //Saldo Devedor 
    balance = balance - amortizationFinal;
    amortization.getRange(nextRow,6).setValue(balance);

    amortization.getRange(nextRow,2).setValue(i);
      

   }


}


function TablePrice() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var amortization = ss.getSheetByName("Amortizacao");
  var interest = amortization.getRange("D6").getValue();
  var time = amortization.getRange("E6").getValue();
  var value = amortization.getRange("C6").getValue();
  amortization.getRange("B15:H180").clear();
  amortization.getRange("B10:D11").setValue("Tabela Price");



  var k = (interest*((1+interest)**time))/(((1+interest)**time)-1); 
  var principal = value * k; 
  var balance = value;


    for(var i=1; i<=time; i++) {
    
    var nextRow = amortization.getLastRow() +1; 

    // Juros 
    var finalInterest = balance * interest; 
    amortization.getRange(nextRow,3).setValue(finalInterest);

      // Amortização
    var amortizationFinal = principal - finalInterest;
    amortization.getRange(nextRow,4).setValue(amortizationFinal); 

  
    //Pagamento 
  amortization.getRange(nextRow,5).setValue(principal);


  
     //Saldo Devedor 
    balance = balance - amortizationFinal;
    amortization.getRange(nextRow,6).setValue(balance);


    


    amortization.getRange(nextRow,2).setValue(i);
      

   }


  return principal; 
}

function CalcX() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var amortization = ss.getSheetByName("Amortizacao");
  amortization.getRange("B15:H180").clear();

  var time = amortization.getRange("E6").getValue();
  var lastRow = amortization.getLastRow() +1; 

   for(var i=1; i<=time; i = i+1) {
      var nextRow = amortization.getLastRow() +1; 
      var printRange = amortization.getRange(nextRow,2).setValue(i);
   }

  }


