//Variables globales
//var rutaWeb =  google.script.url();
var rutaWeb = "https://script.google.com/a/macros/sedic.com.co/s/AKfycbyEdqUoNHWlqWMIilTy5HVsnTmpppSrqnbivMzDo8TSmp7psgkiIWFfuFOFDTgtee3y/exec"
//var rutaWeb = "https://script.google.com/macros/s/AKfycbz2aHx3QLV7MkP8h7_LXlSsxsgDWq4XH3t0Xbb0vek/dev";
var ssId = "11q-sSgjnXKFy1xTCFIvpEvZaL16iBtAEa7TdPaq0lDs";
var ssDb = "BD";
 

//Obtenemos correo y nombre para mostrar en el front
function userId() {

  var correo = Session.getActiveUser().getEmail();
  var nombre = ContactsApp.getContact(correo).getFullName();


  var usuario = {
    nombre: nombre,
    correo: correo
  }
  return (usuario)
}

/*
function doGet(e) {

  var userData = userDatos()
  var politicas = checkPoliticas()
  if(!politicas[0]){
    
    var page = 'Home';
  }
  else if (userData.cedula != '') {
    var page = e.parameter.p || 'Home';
  } else {
    var page = 'registro';
  }

  return HtmlService.createTemplateFromFile(page).evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
    .setTitle('DATATALENTO')
    .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');

}
*/

function include(filename) {

  return HtmlService.createHtmlOutputFromFile(filename).getContent();

}

function userDatos(state) {
  //obtenemos correo y nombre
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName('Usuarios');
  //var db2 = ss.getSheetByName(configuracion);

  var data = db.getDataRange().getValues();
 // var data2 = db.getDataRange().getValues();


  var datosUser = ShowValues();



  var correo = datosUser.email
  var nombre = datosUser.name
  var cedula = ""

  //Obtenemos la cedula del registro
  for (i = 1; i < data.length; i++) {
    if (data[i][4] == correo) {
      var cedula = data[i][0]
    }
  }

  //devolvemos un objeto con la información personal
  var arrayUser = {
    correo: correo,
    nombre: nombre,
    cedula: cedula
  }

  return arrayUser
}

function findDataRegistro() {
  var tabla = 'Registro'
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var data = db.getDataRange().getValues();
  var user = userDatos()
  var status = false
  var datos = []
  var info = {
    datos: datos,
    status: status
  }

  var keyCol = 0

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < data[0].length; i++) {
    if (data[0][i] == "Cédula") {
      keyCol = i
      break
    }
  }

  for (i = 1; i < data.length; i++) {

    if (data[i][keyCol] == user.cedula) {
      status = true
      datos = data[i]
      //transformamos la fecha en año-mes-dia para que el input la reconozca
      var addedTime = Utilities.formatDate(new Date(data[i][5]), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      datos[5] = addedTime
      info = {
        datos: datos,
        status: status
      }
      break
    }
  }

  return JSON.stringify(info)
}

function findData2(tabla) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var data = db.getDataRange().getValues();
  var user = userDatos()
  var status = false
  var datos = []
  var info = {
    datos: datos,
    status: status
  }
  var nMax = 0
  var keyCol = 0
  var idCol = 0
  var colFecha = []

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < data[0].length; i++) {
    if (data[0][i] == "Cédula") {
      keyCol = i
      break
    }
  }

  //Buscamos las columnas que tengan Fecha
  for (i=0;i< data[0].length;i++){
    var fechaE= data[0][i].slice(0,5)
    if(fechaE == "Fecha"){
    colFecha.push(i)
    }
  }

  //Buscamos en la tabla la columna que lleva el id
  for (i = 0; i < data[0].length; i++) {
    if (data[0][i] == "Id") {
      idCol = i
      break
    }
  }

  //buscamos el id más grande en la tabla asociado a la cedula
  for (i = 1; i < data.length; i++) {
    if (data[i][keyCol] == user.cedula && data[i][idCol] > nMax) {
      nMax = data[i][idCol]
    }
  }

  //buscamos las personas relacionados a la cedula
  for (i = 1; i < data.length; i++) {

    if (data[i][keyCol] == user.cedula && data[i][idCol] == nMax) {
      status = true

      for(j=0;j<colFecha.length;j++){
        
        var fechaC = data[i][colFecha[j]]
       data[i][colFecha[j]] = fechas(fechaC)
       
      }

      datos.push(data[i])
      info = {
        datos: datos,
        status: status
      }
    }
  }
  return JSON.stringify(info)
}


/**
 * @function findDataAllMapa
 * Carga TODAS las filas de "Mapa de Conocimiento" para la cédula del usuario.
 * No filtra por Id.
 */
function findDataAllMapa() {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Mapa de Conocimiento");
  var data = db.getDataRange().getValues();
  var user = userDatos();
  var status = false;
  var datos = [];
  var keyCol = 0;

  for (var i = 0; i < data[0].length; i++) {
    if (data[0][i] == "Cédula") {
      keyCol = i;
      break;
    }
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][keyCol] == user.cedula) {
      status = true;
      datos.push(data[i]);
    }
  }

  var info = { datos: datos, status: status };
  Logger.log("findDataAllMapa: encontradas " + datos.length + " filas");
  return JSON.stringify(info);
}


/**
 * @function appendMapaRow
 * Agrega UNA SOLA fila a "Mapa de Conocimiento". No borra nada.
 */
function appendMapaRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Mapa de Conocimiento");
  var lr = db.getLastRow() + 1;
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendMapaRow: agregada fila en posición " + lr);
}


/**
 * @function eliminarMapaRow
 * Elimina una fila específica por cédula + tipo + conocimiento.
 */
function eliminarMapaRow(cedula, tipo, conocimiento) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Mapa de Conocimiento");
  var data = db.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == cedula &&
        data[i][1].toString().trim() == tipo.toString().trim() &&
        data[i][2].toString().trim() == conocimiento.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarMapaRow: eliminada fila " + (i + 1));
      return true;
    }
  }
  return false;
}


function findData(tabla) {

  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var data = db.getDataRange().getValues();
  var user = userDatos()
  var status = false
  var datos = []
  var info = {
    datos: datos,
    status: status
  }
  var keyCol = 0

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < data[0].length; i++) {
    if (data[0][i] == "Cédula") {
      keyCol = i
      break
    }
  }

  //buscamos la info relacionados a la cedula
  for (i = 1; i < data.length; i++) {

    if (data[i][keyCol] == user.cedula) {
      status = true
      datos.push(data[i])
      info = {
        datos: datos,
        status: status
      }
    }
  }
  Logger.log(info)
  return JSON.stringify(info)
}

function postInDB(data, tabla, ind, state) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var dataTable = db.getDataRange().getValues();
  var lr = db.getLastRow() + 1;
  var keyCol = 0

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < dataTable[0].length; i++) {
    if (dataTable[0][i] == "Cédula") {
      keyCol = i
      break
    }
  }

  if (state == true) {
    //Buscamos en la tabla si ya se realizo el registro previamente
    for (i = 1; i < dataTable.length; i++) {
      if (dataTable[i][keyCol] == data[ind]) {
        lr = i + 1
        break
      }
    }
  }


  if (typeof data[0] != "object") {
    data = [data]

  }

  db.getRange(lr, 1, data.length, data[0].length).setValues(data)
}

//Json con las ciudades de colombia
function getJSON() {
  var aUrl = "https://raw.githubusercontent.com/marcovega/colombia-json/master/colombia.min.json";
  var response = UrlFetchApp.fetch(aUrl); // get feed
  var dataAll = JSON.parse(response); //

  for (item in dataAll) {
    var departamentos = {
      departamento: dataAll[item].departamento,
      indice: item
    }
  }

  Logger.log(dataAll.find(item => item.departamento == 'Santander'))

}

//Se crean listas que obtiene datos de la hoja maestros
function listas(lista) {

  var ss = SpreadsheetApp.openById(ssId);
  var maestros = ss.getSheetByName('Maestros');

  //Numero de elementos por cada lista el +2 es para traer los valores directamente
  var nEspecialidad = maestros.getRange("A2").getValue() + 2;
  var nTipo = maestros.getRange("B2").getValue() + 2;


  //listas
  var especialidad = maestros.getRange("A3:A" + nEspecialidad).getValues();
  var tipo = maestros.getRange("B3:B" + nTipo).getValues();


  //retorno
  switch (lista) {
    case "especialidad":
      return especialidad
      break;
    case "tipo":
      return tipo
      break;
  }
}

function politicas() {
  var user = ShowValues();
  var fecha = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy/MM/dd');
  var arrayPoliticas = [[user.email, user.name, 'Aceptado', fecha]]

  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName('Tratamiento de datos');
  var dataTable = db.getDataRange().getValues();
  var lr = db.getLastRow()+1;
  var statusPoliticas = false

  //Buscamos en la tabla si ya se realizo el registro previamente
  for (i = 1; i < dataTable.length; i++) {
    if (dataTable[i][0] == arrayPoliticas[0][0]) {
      statusPoliticas = true
      break
    }
  }

  if (typeof arrayPoliticas[0] != "object") {
    arrayPoliticas = [arrayPoliticas]

  }

  if(!statusPoliticas){
    db.getRange(lr, 1, arrayPoliticas.length, arrayPoliticas[0].length).setValues(arrayPoliticas)
  }
}

function checkPoliticas(){
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName('Tratamiento de datos');
  var data = db.getDataRange().getValues();
  var user = ShowValues()
  var status = false
  var datos = [status, rutaWeb]

  //buscamos la info relacionados a la cedula
  for (i = 1; i < data.length; i++) {

    if (data[i][0] == user.email) {
      status = true
      break
    }
  }
  var datos = [status, rutaWeb]
  return datos
}

function testDiagnostico() {
  try {
    // Intentamos conectar con la hoja activa
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const nombre = libro.getName();
    
    Logger.log("✅ ÉXITO: El script está conectado a: " + nombre);
    Logger.log("🆔 ID de la Hoja: " + libro.getId());
  } catch (e) {
    Logger.log("❌ ERROR: Este script no está vinculado correctamente.");
    Logger.log("Detalle: " + e.toString());
  }
}


