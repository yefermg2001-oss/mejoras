/**
 * @function findDataRegistro 
 * Carga los datos de la tabla Registro y lego los datos de la tabla que envié el parametro op
 * @param {string} op
 */
function findDataRegistro2(op) {
  //var op = "Experiencia"
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName('Registro');
  var dataRegistro = db.getDataRange().getValues();
  var info2 = {}


  switch (op) {

    case "Formación":
      info2 = findDataAllFormacion();
      Logger.log('caso formacion ' + info2.datos)
      break;

    case "Experiencia":
      info2 = findDataAllExperiencia();
      Logger.log('caso experiencia ' + info2.datos)
      break;
  }


  var user = userDatos()
  var status = false
  var datos = []
  var info = {
    user: user,
    datos: datos,
    status: status
  }

  var keyCol = 0

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < dataRegistro[0].length; i++) {
    if (dataRegistro[0][i] == "Cédula") {

      keyCol = i

      break
    }
  }

  for (i = 1; i < dataRegistro.length; i++) {

    if (dataRegistro[i][keyCol] == user.cedula) {

      

      status = true
      datos = dataRegistro[i]

      info = {
        datos: datos,
        status: status,

      }
      break
    }
  }

  if (info.status == true && info2.status == true) {

     info2.global = op
     Logger.log('opcion 1')
    return JSON.stringify(info2)

  
  } else if(info.status == true && info2.status == false){

    info.global = "Registro"
    Logger.log('opcion 2')
    return JSON.stringify(info)

 } else if (info.status == false) {

    info = {}
     Logger.log('opcion 3')
    return JSON.stringify(info)

}
}




 function findData3(tabla) {
      //tabla = "Formación"
      var ss = SpreadsheetApp.openById(ssId);
      var db = ss.getSheetByName(tabla);
      var dataInfo = db.getDataRange().getValues();
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
  var idFecha = []
  var colFecha = []

  

  //Buscamos en la tabla la columna que lleva la key
  for (i = 0; i < dataInfo[0].length; i++) {
    if (dataInfo[0][i] == "Cédula") {
      keyCol = i
      break
    }
  }

  //Buscamos en la tabla la columna que lleva el id
  for (i = 0; i < dataInfo[0].length; i++) {
    if (dataInfo[0][i] == "Id") {
      idCol = i
      break
    }
  }
  
  //buscamos el id más grande en la tabla
  for (i = 1; i < dataInfo.length; i++) {
    if (dataInfo[i][keyCol] == user.cedula && dataInfo[i][idCol] > nMax) {
      nMax = dataInfo[i][idCol]
    }
  }


for (i=0;i< dataInfo[0].length;i++){
 var fechaE= dataInfo[0][i].slice(0,5)
       if(fechaE == "Fecha"){
         //idFecha.push(dataInfo[0][i])
         colFecha.push(i)
        

  }
}

       

  
 

  //buscamos las personas relacionados a la cedula
  for (i = 0; i < dataInfo.length; i++) {

     //console.log(fechas(dataInfo[i][colFecha[i]]))

    if (dataInfo[i][keyCol] == user.cedula && dataInfo[i][idCol] == nMax) {
      status = true

      for(j=0;j<colFecha.length;j++){
        
        var fechaC = dataInfo[i][colFecha[j]]
       dataInfo[i][colFecha[j]] = fechas(fechaC.toString())
       
      }

      datos.push(dataInfo[i])

      info = {
        user: user,
        datos: datos,
        status: status,
        
      }
    }
  }
 Logger.log(info.datos)
  return info;


}


/**
 * @function findDataAllExperiencia
 * Carga TODAS las filas de Experiencia para la cédula del usuario.
 * No filtra por Id, así se cargan todas las experiencias independientemente
 * de cuándo fueron guardadas.
 */
function findDataAllExperiencia() {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Experiencia");
  var dataInfo = db.getDataRange().getValues();
  var user = userDatos();
  var status = false;
  var datos = [];
  var info = { datos: datos, status: status };
  var keyCol = 0;
  var colFecha = [];

  // Buscar columna Cédula
  for (var i = 0; i < dataInfo[0].length; i++) {
    if (dataInfo[0][i] == "Cédula") {
      keyCol = i;
      break;
    }
  }

  // Buscar columnas de fecha
  for (var i = 0; i < dataInfo[0].length; i++) {
    var fechaE = dataInfo[0][i].slice(0, 5);
    if (fechaE == "Fecha") {
      colFecha.push(i);
    }
  }

  // Cargar TODAS las filas del usuario (sin filtro de Id)
  for (var i = 1; i < dataInfo.length; i++) {
    if (dataInfo[i][keyCol] == user.cedula) {
      status = true;

      // Formatear fechas
      for (var j = 0; j < colFecha.length; j++) {
        var fechaC = dataInfo[i][colFecha[j]];
        dataInfo[i][colFecha[j]] = fechas(fechaC.toString());
      }

      datos.push(dataInfo[i]);
    }
  }

  info = {
    user: user,
    datos: datos,
    status: status
  };

  Logger.log("findDataAllExperiencia: encontradas " + datos.length + " filas para " + user.cedula);
  return info;
}

/**
 * @function fechas
 * Asigna formato a las fechas tipo date
 * @param {date} fecha
 * @returns {date} fechaRes retorna la fecha con el formato definido segun la condición
 */
function fechas(fecha) {
Logger.log('esta es la fecha '+fecha)

var fecha1 = new Date(fecha)

  if (fecha == "") {
    var fechaRes = "";
    return fechaRes;
  } else {
    var fechaRes = Utilities.formatDate(fecha1, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
   
    return fechaRes;
  }
}



/**
 * @function postInDB
 * Envia la información a la base de datos
 * @param {object} dataInfo
 * @param {string} tabla
 * @param {number} ind
 * @param {boolean} state
*/

function postInDB2(dataInfo, tabla, ind, state) {
  
  
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var dataTable = db.getDataRange().getValues();
  var lr = db.getLastRow() + 1;
  var keyCol = 0
 Logger.log('Estos son los datos enviados '+dataInfo+' fila  '+lr)
  
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
      if (dataTable[i][keyCol] == dataInfo[ind]) {
        lr = i + 1
        break
      }
    }
  }



  
    if (typeof dataInfo[0] != "object") {
      dataInfo = [dataInfo]
  
    }


    
  // Logger.log('Esta es la dataInfo '+dataInfo+"  "+lr)
  db.getRange(lr, 1, dataInfo.length, dataInfo[0].length).setValues(dataInfo)
  /*var id = 0;
  for (i in dataInfo) {
    id = i
    db.appendRow([
      new Date(),
      dataInfo[i].cedula,
      dataInfo[i].nivelEsc,
      dataInfo[i].universidad,
      dataInfo[i].titulo,
      dataInfo[i].estado, dataInfo.fechaGraduacion,
      dataInfo[i].adjunto,
      dataInfo[i].tarjeta,
      dataInfo[i].numTarjeta,
      dataInfo[i].seccional,
      dataInfo[i].expedicion,
      id]);

  }
  */

}


/**
 * @function appendExperienciaRow
 * Agrega UNA SOLA fila de experiencia a la hoja. No borra nada.
 * @param {Array} rowData - Array plano con los valores de una fila
 */
function appendExperienciaRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Experiencia");
  var lr = db.getLastRow() + 1;
  
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendExperienciaRow: agregada fila en posición " + lr);
}


/**
 * @function eliminarExperienciaRow
 * Elimina una fila específica de la hoja Experiencia buscando por cédula + empresa + cargo + fechaInicio
 * @param {string} cedula
 * @param {string} empresa
 * @param {string} cargo
 * @param {string} fechaIn
 */
function eliminarExperienciaRow(cedula, empresa, cargo, fechaIn) {
  // Verificar que el usuario solo borre sus propios datos
  if (!_verificarCedulaPropia(cedula)) {
    Logger.log("SEGURIDAD: Intento no autorizado de borrar experiencia de cédula " + cedula);
    return false;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Experiencia");
  var dataTable = db.getDataRange().getValues();
  
  // Encontrar columnas clave
  var keyCol = 0, empCol = 2, cargoCol = 3, fechaCol = 4;
  for (var i = 0; i < dataTable[0].length; i++) {
    if (dataTable[0][i] == "Cédula") keyCol = i;
  }
  
  // Buscar y borrar la fila (de abajo hacia arriba)
  for (var i = dataTable.length - 1; i >= 1; i--) {
    var rowCedula = dataTable[i][keyCol];
    var rowEmpresa = dataTable[i][empCol];
    var rowCargo = dataTable[i][cargoCol];
    var rowFechaIn = dataTable[i][fechaCol];
    
    // Formatear fecha de la hoja para comparar
    if (rowFechaIn instanceof Date) {
      rowFechaIn = Utilities.formatDate(rowFechaIn, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    }
    
    if (rowCedula == cedula && 
        rowEmpresa.toString().trim() == empresa.toString().trim() && 
        rowCargo.toString().trim() == cargo.toString().trim() && 
        rowFechaIn.toString().trim() == fechaIn.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarExperienciaRow: borrada fila " + (i + 1) + " para " + cedula + " - " + empresa);
      return true;
    }
  }
  
  Logger.log("eliminarExperienciaRow: no se encontró fila para " + cedula + " - " + empresa);
  return false;
}

/**
 * @function findDataAllFormacion
 * Carga TODAS las filas de Formación para la cédula del usuario.
 */
function findDataAllFormacion() {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Formación");
  var dataInfo = db.getDataRange().getValues();
  var user = userDatos();
  var status = false;
  var datos = [];
  var info = { datos: datos, status: status };
  var keyCol = 0;
  var colFecha = [];

  // Buscar columna Cédula
  for (var i = 0; i < dataInfo[0].length; i++) {
    if (dataInfo[0][i] == "Cédula") {
      keyCol = i;
      break;
    }
  }

  // Buscar columnas de fecha
  for (var i = 0; i < dataInfo[0].length; i++) {
    var fechaE = dataInfo[0][i].toString().slice(0, 5);
    if (fechaE == "Fecha") {
      colFecha.push(i);
    }
  }

  // Cargar TODAS las filas del usuario
  for (var i = 1; i < dataInfo.length; i++) {
    if (dataInfo[i][keyCol] == user.cedula) {
      status = true;

      for (var j = 0; j < colFecha.length; j++) {
        var fechaC = dataInfo[i][colFecha[j]];
        dataInfo[i][colFecha[j]] = fechas(fechaC.toString());
      }

      datos.push(dataInfo[i]);
    }
  }

  info = {
    user: user,
    datos: datos,
    status: status
  };

  Logger.log("findDataAllFormacion: encontradas " + datos.length + " filas para " + user.cedula);
  return info;
}

/**
 * @function appendFormacionRow
 * Agrega UNA SOLA fila de formación a la hoja.
 * @param {Array} rowData
 */
function appendFormacionRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Formación");
  var lr = db.getLastRow() + 1;
  
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendFormacionRow: agregada fila en posición " + lr);
}

/**
 * @function eliminarFormacionRow
 * Elimina una fila de la hoja Formación por cédula + nivelEsc + titulo
 * @param {string} cedula
 * @param {string} nivelEsc
 * @param {string} titulo
 */
function eliminarFormacionRow(cedula, nivelEsc, titulo) {
  // Verificar que el usuario solo borre sus propios datos
  if (!_verificarCedulaPropia(cedula)) {
    Logger.log("SEGURIDAD: Intento no autorizado de borrar formación de cédula " + cedula);
    return false;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Formación");
  var dataTable = db.getDataRange().getValues();
  
  var keyCol = 0;
  for (var i = 0; i < dataTable[0].length; i++) {
    if (dataTable[0][i] == "Cédula") { keyCol = i; break; }
  }
  
  // nivelEsc está en col 2, titulo en col 4
  for (var i = dataTable.length - 1; i >= 1; i--) {
    if (dataTable[i][keyCol] == cedula && 
        dataTable[i][2].toString().trim() == nivelEsc.toString().trim() && 
        dataTable[i][4].toString().trim() == titulo.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarFormacionRow: borrada fila " + (i + 1));
      return true;
    }
  }
  
  Logger.log("eliminarFormacionRow: no se encontró fila para " + cedula + " - " + titulo);
  return false;
}

/**
 * @function appendPersonaRow
 * Agrega UNA SOLA fila de Persona a Cargo a la hoja.
 * @param {Array} rowData
 */
function appendPersonaRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Personas a cargo");
  var lr = db.getLastRow() + 1;
  
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendPersonaRow: agregada fila en posición " + lr);
}

/**
 * @function eliminarPersonaRow
 * Elimina una fila de Personas a Cargo por cédula + nombre + parentesco
 * @param {string} cedula
 * @param {string} nombre
 * @param {string} parentesco
 */
function eliminarPersonaRow(cedula, nombre, parentesco) {
  // Verificar que el usuario solo borre sus propios datos
  if (!_verificarCedulaPropia(cedula)) {
    Logger.log("SEGURIDAD: Intento no autorizado de borrar persona a cargo de cédula " + cedula);
    return false;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Personas a cargo");
  var dataTable = db.getDataRange().getValues();
  
  var keyCol = 0;
  // Cédula en col 0
  
  for (var i = dataTable.length - 1; i >= 1; i--) {
    if (dataTable[i][0] == cedula && 
        dataTable[i][1].toString().trim() == nombre.toString().trim() && 
        dataTable[i][2].toString().trim() == parentesco.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarPersonaRow: borrada fila " + (i + 1));
      return true;
    }
  }
  return false;
}

/**
 * @function appendHijoRow
 * Agrega UNA SOLA fila de Hijo a la hoja.
 * @param {Array} rowData
 */
function appendHijoRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Hijos");
  var lr = db.getLastRow() + 1;
  
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendHijoRow: agregada fila en posición " + lr);
}

/**
 * @function eliminarHijoRow
 * Elimina una fila de Hijos por cédula + nombre + fecha
 * @param {string} cedula
 * @param {string} nombre
 * @param {string} fecha
 */
function eliminarHijoRow(cedula, nombre, fecha) {
  // Verificar que el usuario solo borre sus propios datos
  if (!_verificarCedulaPropia(cedula)) {
    Logger.log("SEGURIDAD: Intento no autorizado de borrar hijo de cédula " + cedula);
    return false;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Hijos");
  var dataTable = db.getDataRange().getValues();
  
  // Formatear fecha si llega como string YYYY-MM-DD
  // En Sheet la fecha puede ser objeto Date o string.
  
  for (var i = dataTable.length - 1; i >= 1; i--) {
    var rowFecha = dataTable[i][2];
    if (rowFecha instanceof Date) {
      rowFecha = Utilities.formatDate(rowFecha, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    }
    
    if (dataTable[i][0] == cedula && 
        dataTable[i][1].toString().trim() == nombre.toString().trim() && 
        rowFecha.toString().trim() == fecha.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarHijoRow: borrada fila " + (i + 1));
      return true;
    }
  }
  return false;
}

/**
 * @function findDataVinculados
 * Carga TODAS las filas de la tabla especificada (Hijos o Personas a cargo)
 * para la cédula del usuario, sin filtrar por ID máximo.
 * @param {string} tabla - Nombre de la hoja ("Hijos" o "Personas a cargo" o "Experiencia" o "Formación")
 */
function findDataVinculados(tabla) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName(tabla);
  var data = db.getDataRange().getValues();
  var user = userDatos();
  var status = false;
  var datos = [];
  var info = { datos: datos, status: status };
  
  var keyCol = 0;
  var colFecha = [];

  // Buscar columna Cédula
  if (data.length > 0) {
      for (var i = 0; i < data[0].length; i++) {
        if (data[0][i] == "Cédula") {
          keyCol = i;
          break;
        }
      }

      // Buscar columnas Fecha
      for (var i = 0; i < data[0].length; i++) {
        var fechaE = data[0][i].toString().slice(0, 5);
        if (fechaE == "Fecha") {
          colFecha.push(i);
        }
      }
  }

  // Filtrar por cédula
  for (var i = 1; i < data.length; i++) {
    if (data[i][keyCol] == user.cedula) {
      status = true;
      
      // Formatear fechas
      for (var j = 0; j < colFecha.length; j++) {
        var fechaC = data[i][colFecha[j]];
        data[i][colFecha[j]] = fechas(fechaC.toString());
      }
      
      datos.push(data[i]);
    }
  }
  
  if (datos.length > 0) {
      status = true;
  }

  info = {
    datos: datos,
    status: status
  };
  
  return JSON.stringify(info);
}


/**
 * @function findDataAllInfoLab
 * Carga TODAS las filas de Información Laboral del usuario actual (sin filtro de Id)
 * Replica el patrón de findDataAllExperiencia
 */
function findDataAllInfoLab() {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Información Laboral");
  var dataInfo = db.getDataRange().getValues();
  var user = userDatos();
  var status = false;
  var datos = [];
  var info = { datos: datos, status: status };
  var keyCol = 0;

  // Buscar columna Cédula
  for (var i = 0; i < dataInfo[0].length; i++) {
    if (dataInfo[0][i] == "Cédula") {
      keyCol = i;
      break;
    }
  }

  // Cargar TODAS las filas del usuario (sin filtro de Id)
  for (var i = 1; i < dataInfo.length; i++) {
    if (dataInfo[i][keyCol].toString().trim() == user.cedula.toString().trim()) {
      status = true;
      datos.push(dataInfo[i]);
    }
  }

  info = {
    datos: datos,
    status: status
  };

  Logger.log("findDataAllInfoLab: encontradas " + datos.length + " filas para " + user.cedula);
  return JSON.stringify(info);
}


/**
 * @function appendInfoLabRow
 * Agrega UNA SOLA fila de Información Laboral a la hoja. No borra nada.
 * @param {Array} rowData - Array plano con los valores de una fila
 */
function appendInfoLabRow(rowData) {
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Información Laboral");
  var lr = db.getLastRow() + 1;
  
  if (typeof rowData[0] != "object") {
    rowData = [rowData];
  }
  
  db.getRange(lr, 1, rowData.length, rowData[0].length).setValues(rowData);
  Logger.log("appendInfoLabRow: agregada fila en posición " + lr);
}


/**
 * @function eliminarInfoLabRow
 * Elimina una fila específica de la hoja Información Laboral buscando por cédula + especialidad + tipo
 * @param {string} cedula
 * @param {string} especialidad
 * @param {string} tipo
 */
function eliminarInfoLabRow(cedula, especialidad, tipo) {
  // Verificar que el usuario solo borre sus propios datos
  if (!_verificarCedulaPropia(cedula)) {
    Logger.log("SEGURIDAD: Intento no autorizado de borrar info laboral de cédula " + cedula);
    return false;
  }
  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Información Laboral");
  var dataTable = db.getDataRange().getValues();
  
  // Columnas de la hoja: A=Cédula(0), B=General(1), C=Especialidad(2), D=Tipo(3)
  var keyCol = 0, espCol = 2, tipoCol = 3;
  for (var i = 0; i < dataTable[0].length; i++) {
    if (dataTable[0][i] == "Cédula") keyCol = i;
    if (dataTable[0][i] == "Especialidad") espCol = i;
    if (dataTable[0][i] == "Tipo") tipoCol = i;
  }
  
  // Buscar y borrar la fila (de abajo hacia arriba)
  for (var i = dataTable.length - 1; i >= 1; i--) {
    var rowCedula = dataTable[i][keyCol];
    var rowEsp = dataTable[i][espCol];
    var rowTipo = dataTable[i][tipoCol];
    
    if (rowCedula.toString().trim() == cedula.toString().trim() && 
        rowEsp.toString().trim() == especialidad.toString().trim() && 
        rowTipo.toString().trim() == tipo.toString().trim()) {
      db.deleteRow(i + 1);
      Logger.log("eliminarInfoLabRow: borrada fila " + (i + 1) + " para " + cedula + " - " + especialidad);
      return true;
    }
  }
  
  Logger.log("eliminarInfoLabRow: no se encontró fila para " + cedula + " - " + especialidad);
  return false;
}
