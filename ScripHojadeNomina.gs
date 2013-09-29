//CONSTANTES

var SPREADSHEET_ID="";   //ID de la hoja de Nomina
var TEMPLATE_ID="";      //ID de la plantilla, los campos a completar en la plantilla deben estar entre %DATO%
var HEADER_SIZE=4;                                                   //Las 4 filas antes de los datos de los empleados, este puede variar, 
var FOOTER_SIZE=2;                                                   //Las 2 filas que siguen despues de los datos de las personas, si se desea cambiar, que sea el valor numerico



//FUNCION QUE SE EJECUTA
function doGet() {
  
  //Soporte para tranformar datos a JSON
  var datos = myFunction();
  return ContentService.createTextOutput(JSON.stringify(datos)).setMimeType(ContentService.MimeType.JSON);
}

function myFunction() {
  
  
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  var limiteEmpleados=numRows-(FOOTER_SIZE); 
  Logger.log("EL mumero de empleados es: "+(numRows-(FOOTER_SIZE+HEADER_SIZE)).toString());
  var personas= []; 
  for(var i=HEADER_SIZE; i<limiteEmpleados; i++){ 
     personas[i-HEADER_SIZE] =values[i]; 
  }

 var nominaPersona = [];
 for (i in personas) {
    
    var row           = personas[i];
    var nombre        = row[0];  
    var salarioDiario = row[1];    
    var diasLaborados = row[2]; 
    var sueldo        = row[3];  
    var gratificacion = row[4]; 
    var bono          = row[5];     
    var ingresos      = row[6]; 
    var isr           = row[7];     
    var imss          = row[8]; 
    var deduciones    = row[9];     
    var neto          = row[10];
    var curp          = row[11];
    var rfc           = row[12];
    var numeroImss    = row[13];
    
    var nomina= new GenerarNomima(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13]);
    Logger.log(nomina);
    Logger.log("\n");
    nominaPersona[i] = nomina;
  }
   
   return nominaPersona;
}

function GenerarNomima(nombre, salarioDiario, diasLaborados, sueldo, gratificacion, bono, ingresos, isr, imss, deduciones, neto, curp, rfc, numeroImss){
  
  this.nombre = nombre;
  this.salarioDiario = salarioDiario;
  this.diasLaborados = diasLaborados;
  this.sueldo = sueldo;
  this.gratificacion = gratificacion;
  this.bono = bono;
  this.ingresos = ingresos;
  this.isr = isr;
  this.imss = imss;
  this.deduciones = deduciones;
  this.neto = neto;
  this.curp = curp;
  this.rfc  = rfc;
  this.numeroImss= numeroImss;
  
  //Solo 2 decimales
   
  salarioDiario=(Math.floor(salarioDiario*100))/100;
  sueldo=(Math.floor(sueldo*100))/100;
  gratificacion=(Math.floor(gratificacion*100))/100;
  bono=(Math.floor(bono*100))/100;
  ingresos=(Math.floor(ingresos*100))/100;
  isr=(Math.floor(isr*100))/100;
  isr=(Math.floor(isr*100))/100;
  imss=(Math.floor(imss*100))/100;
  deduciones=(Math.floor(deduciones*100))/100;
  neto=(Math.floor(neto*100))/100;
  
  //Crea nombre del Mes por el atributo DATE().GETMOUNTH
  
  var now = new Date();
  var mes = "";
  
  if((now.getMonth())==0){
    mes="Enero";
  }else if((now.getMonth())==1){
      mes="Febrero";
  }else if((now.getMonth())==2){
   mes="Marzo";
  }else if((now.getMonth())==3){
    mes="Abril";
  }else if((now.getMonth())==4){
    mes="Mayo";
  }else if((now.getMonth())==5){
    mes="Junio";
  }else if((now.getMonth())==6){
    mes="Julio";
  }else if((now.getMonth())==7){
    mes="Agosto"
  }else if((now.getMonth())==8){
    mes="Septiembre";
  }else if((now.getMonth())==9){
    mes="Octubre";  
  }else if((now.getMonth())==10){
    mes="Noviembre";
  }else{
    mes="Diciembre";
  }
    

  var fecha = (now.getDate().toString()+"/ "+mes+"/ "+now.getFullYear().toString());
    
  var docId = DocsList.getFileById(TEMPLATE_ID).makeCopy().getId();
  var doc = DocumentApp.openById(docId);
  doc.setName("Recibo de Nomina "+nombre+"  Mes de "+mes);
  var body = doc.getActiveSection();
 
  body.replaceText("%nombre%", nombre);
  body.replaceText("%rfc%", rfc);
  body.replaceText("%numeroImss%", numeroImss);
  body.replaceText("%curp%", curp);
  body.replaceText("%date%", fecha);
  body.replaceText("%diasLaborados%", diasLaborados);
  body.replaceText("%salarioDiario%",  salarioDiario);
  body.replaceText("%sueldo%", sueldo);
  body.replaceText("%isr%", isr);
  body.replaceText("%imss%", imss);
  body.replaceText("%deduciones%", deduciones);
  body.replaceText("%neto%", neto);
  
  doc.saveAndClose();
  
  //Una pausa para el servidor, de no ser asi se satura y manda error
  Utilities.sleep(3000);
  
  //Soporte para mandar los datos a GMAIL
  //GmailApp.sendEmail("tucorreo@gmail.com", "nomina de: "+nombre, "La nomina de: "+nombre+"se encuentra en la siguiente liga " + doc.getUrl()+"\n \n  Gracias!!");
 
}


//Crea el menu para generar la nomina, se agrego un trigger para mostrarlo una vez que se abra el spreadsheet
function createMenu(){
  var menuEntries = [
     { 
         name : "GenerarNomina", functionName : "doGet" },
         null,
     ];
     SpreadsheetApp.getActiveSpreadsheet().addMenu( "Crea la nomina de Viajez.com ", menuEntries );
}
  







