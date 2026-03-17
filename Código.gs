//Variables globales
//var rutaWeb =  google.script.url();
var rutaWeb = "https://script.google.com/a/macros/sedic.com.co/s/AKfycbyZw-cnwlWNcFyqx7sCkvLAw6SiCgPOOKZv4quxK6BqIgRbKWwDEGgt3Iz3GqYsaszkZA/exec"
//var rutaWeb = "https://script.google.com/macros/s/AKfycbz2aHx3QLV7MkP8h7_LXlSsxsgDWq4XH3t0Xbb0vek/dev";
var ssId = "1Q7KH2rEwvxJubf2UKf2zdJPvI8m2c093S_rxJr3juHY";
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


/**
 * @function checkAcceso
 * Verifica si el usuario actual tiene rol de "admin" en la columna "Acceso" de la hoja Usuarios.
 * @returns {boolean} true si el usuario es admin, false si no
 */
function checkAcceso() {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName('Usuarios');
  var data = db.getDataRange().getValues();
  var user = ShowValues();

  // Buscar columna "Acceso"
  var accesoCol = -1;
  for (var i = 0; i < data[0].length; i++) {
    if (data[0][i].toString().trim().toLowerCase() == "acceso") {
      accesoCol = i;
      break;
    }
  }

  if (accesoCol == -1) {
    Logger.log("checkAcceso: columna 'Acceso' no encontrada en Usuarios");
    return false;
  }

  // Buscar email del usuario actual
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] == user.email) {
      var acceso = data[i][accesoCol].toString().trim().toLowerCase();
      Logger.log("checkAcceso: usuario " + user.email + " tiene acceso=" + acceso);
      return acceso === "admin";
    }
  }

  Logger.log("checkAcceso: usuario " + user.email + " no encontrado");
  return false;
}


/**
 * @function buscarTalento
 * Carga TODOS los datos de Registro, Experiencia y Formación,
 * los cruza por cédula y retorna un JSON consolidado para búsqueda client-side.
 * Solo accesible por admins (validado en el frontend).
 * @returns {string} JSON con estructura { personas: [...], filtros: {...} }
 */
function buscarTalento() {
  var ss = SpreadsheetApp.openById(ssId);
  
  // --- Leer Registro ---
  var dbRegistro = ss.getSheetByName('Registro');
  var dataRegistro = dbRegistro.getDataRange().getValues();
  
  // --- Leer Experiencia ---
  var dbExp = ss.getSheetByName('Experiencia');
  var dataExp = dbExp.getDataRange().getValues();
  
  // --- Leer Formación ---
  var dbForm = ss.getSheetByName('Formación');
  var dataForm = dbForm.getDataRange().getValues();
  
  // --- Leer Información Laboral ---
  var dbInfoLab = ss.getSheetByName('Información Laboral');
  var dataInfoLab = dbInfoLab ? dbInfoLab.getDataRange().getValues() : [];
  
  // --- Construir mapa de personas por cédula ---
  var personasMap = {};
  var clasificacionesSet = {};
  var tiposProyectoSet = {};
  var nivelesFormacionSet = {};
  var sectoresSet = {};
  
  // Procesar Registro (col 0=Fecha, 1=Nombre, 2=Correo, 3=Cédula)
  // Buscar columnas dinámicamente
  var regCedulaCol = 0;
  for (var i = 0; i < dataRegistro[0].length; i++) {
    if (dataRegistro[0][i] == "Cédula") { regCedulaCol = i; break; }
  }
  
  for (var i = 1; i < dataRegistro.length; i++) {
    var cedula = dataRegistro[i][regCedulaCol].toString().trim();
    if (cedula == "") continue;
    
    personasMap[cedula] = {
      cedula: cedula,
      nombre: dataRegistro[i][1] || "",
      correo: dataRegistro[i][2] || "",
      experiencias: [],
      formaciones: [],
      infoLaboral: []
    };
  }
  
  // Procesar Experiencia
  // Columnas: Fecha(0), Cédula(1), Empresa(2), Cargo(3), FechaIn(4), FechaFin(5), 
  //           Objeto(6), Funciones(7), Adjunto(8), Sector(9), Clasificación(10), TipoProy(11), Dedicación(12)
  var expCedulaCol = 0;
  if (dataExp.length > 0) {
    for (var i = 0; i < dataExp[0].length; i++) {
      if (dataExp[0][i] == "Cédula") { expCedulaCol = i; break; }
    }
  }
  
  for (var i = 1; i < dataExp.length; i++) {
    var cedula = dataExp[i][expCedulaCol].toString().trim();
    if (cedula == "" || !personasMap[cedula]) continue;
    
    var fechaIn = dataExp[i][4];
    var fechaFin = dataExp[i][5];
    
    // Formatear fechas
    if (fechaIn instanceof Date) {
      fechaIn = Utilities.formatDate(fechaIn, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    }
    if (fechaFin instanceof Date) {
      fechaFin = Utilities.formatDate(fechaFin, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    }
    
    var sector = (dataExp[i][9] || "").toString().trim();
    var clasificacion = (dataExp[i][10] || "").toString().trim();
    var tipoProy = (dataExp[i][11] || "").toString().trim();
    var dedicacion = (dataExp[i][12] || "").toString().trim();
    
    if (sector) sectoresSet[sector] = true;
    if (clasificacion) clasificacionesSet[clasificacion] = true;
    if (tipoProy) tiposProyectoSet[tipoProy] = true;
    
    personasMap[cedula].experiencias.push({
      empresa: (dataExp[i][2] || "").toString(),
      cargo: (dataExp[i][3] || "").toString(),
      fechaIn: fechaIn.toString(),
      fechaFin: fechaFin.toString(),
      objeto: (dataExp[i][6] || "").toString(),
      funciones: (dataExp[i][7] || "").toString(),
      sector: sector,
      clasificacion: clasificacion,
      tipoproy: tipoProy,
      dedicacion: dedicacion
    });
  }
  
  // Procesar Formación
  // Columnas: Fecha(0), Cédula(1), NivelEsc(2), Universidad(3), Titulo(4), Estado(5), FechaGraduacion(6)
  var formCedulaCol = 0;
  if (dataForm.length > 0) {
    for (var i = 0; i < dataForm[0].length; i++) {
      if (dataForm[0][i] == "Cédula") { formCedulaCol = i; break; }
    }
  }
  
  for (var i = 1; i < dataForm.length; i++) {
    var cedula = dataForm[i][formCedulaCol].toString().trim();
    if (cedula == "" || !personasMap[cedula]) continue;
    
    var nivel = (dataForm[i][2] || "").toString().trim();
    if (nivel) nivelesFormacionSet[nivel] = true;
    
    personasMap[cedula].formaciones.push({
      nivel: nivel,
      universidad: (dataForm[i][3] || "").toString(),
      titulo: (dataForm[i][4] || "").toString(),
      estado: (dataForm[i][5] || "").toString()
    });
  }
  
  // Procesar Información Laboral
  if (dataInfoLab.length > 0) {
    var infoLabCedulaCol = 0;
    for (var i = 0; i < dataInfoLab[0].length; i++) {
      if (dataInfoLab[0][i] == "Cédula") { infoLabCedulaCol = i; break; }
    }
    
    for (var i = 1; i < dataInfoLab.length; i++) {
      var cedula = dataInfoLab[i][infoLabCedulaCol].toString().trim();
      if (cedula == "" || !personasMap[cedula]) continue;
      
      personasMap[cedula].infoLaboral.push({
        especialidad: (dataInfoLab[i][2] || "").toString(),
        tipo: (dataInfoLab[i][3] || "").toString()
      });
    }
  }
  
  // Convertir mapa a array
  var personas = [];
  for (var key in personasMap) {
    personas.push(personasMap[key]);
  }
  
  var resultado = {
    personas: personas,
    filtros: {
      sectores: Object.keys(sectoresSet).sort(),
      clasificaciones: Object.keys(clasificacionesSet).sort(),
      tiposProyecto: Object.keys(tiposProyectoSet).sort(),
      nivelesFormacion: Object.keys(nivelesFormacionSet).sort()
    }
  };
  
  Logger.log("buscarTalento: " + personas.length + " personas cargadas");
  return JSON.stringify(resultado);
}

