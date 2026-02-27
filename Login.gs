var configuracion = "Usuarios";
var datos = {};

//Esta funcion inicia la llamada de la pagina y recibe peticiones de tipo GET
function doGet(e) {

  var page = e.parameter.p || 'index';
  PutValues("", "", "");

  var html = HtmlService.createTemplateFromFile(page);
  html.mensaje = "";

  return html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
    .setTitle('DATATALENTO')
    .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');

}
//Close función doGet


//Esta funcion recibe peticiones y parametros de tipo POST
function doPost(e) {

  var ss = SpreadsheetApp.openById(ssId);
  var db = ss.getSheetByName("Registro");
  var data = db.getDataRange().getValues();


  Logger.log(JSON.stringify(e));

  var user = e.parameter.username || " ";
  var password = e.parameter.password || " ";

  Logger.log(typeof (user));
  Logger.log(typeof (password));

  datos = checkLogin(user, password);
  Logger.log("estado del check " + datos.sw);


  if (datos.sw) {

    //Enviamos los datos a la memoria cache
    PutValues(datos.correo, datos.nombre, datos.contrasena);
    ShowValues();

    var page = e.parameter.btnLogin || 'Home';

    var html = HtmlService.createTemplateFromFile(page);
    html.mensaje = "";
    return html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
      .setTitle('DATATALENTO')
      .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');


  } else if (e.parameter.btnLogin != "") {





    var cacheVal = ShowValues();
    var datos = checkLogin(cacheVal.email, cacheVal.pass);
    Logger.log("El valor dentro del if es " + datos.sw);
    var politicas = checkPoliticas()

    if (datos.sw && e.parameter.btnLogin == "index") {
      PutValues("", "", "");
      var page = 'index';

      var html = HtmlService.createTemplateFromFile(page);
      html.mensaje = "";
      return html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
        .setTitle('DATATALENTO')
        .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');


    } else if (datos.sw) {
      var userData = userDatos()
     
       if(userData.cedula != ""){
         var page = e.parameter.btnLogin;
       }else{
         var page = "registro";
       }
       


      //var page = e.parameter.btnLogin;

      var html = HtmlService.createTemplateFromFile(page);
      html.mensaje = "";
      return html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
        .setTitle('DATATALENTO')
        .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');

    } else {

      var page = 'index';
      PutValues("", "", "");

      var html = HtmlService.createTemplateFromFile(page);
      html.mensaje = "";
      return html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
        .setTitle('DATATALENTO')
        .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');


    }

  } else if (datos.sw == false && user == " " || password == " ") {

    var page = 'index';
    PutValues("", "", "");


    var html = HtmlService.createTemplateFromFile(page);
    html.mensaje = "El campo de usuario o contraseña se encuentra vacío";
    return html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
      .setTitle('DATATALENTO')
      .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');


  } else {

    var page = 'index';
    PutValues("", "", "");

    var html = HtmlService.createTemplateFromFile(page);
    html.mensaje = "Usuario o contraseña incorrectos";
    return html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
      .setTitle('DATATALENTO')
      .setFaviconUrl('https://drive.google.com/uc?id=11VmVU-VhcrAi-_GjaxM048OZoL-2WpeM#.ico');

  }
}
//Close función POST






//Esta funcion valida si el usuario y la contraseña ingresados en el input son iguales a los de la base de datos
function checkLogin(user, password) {

  var spreadsheet = SpreadsheetApp.openById(ssId);
  var sheeta = spreadsheet.getSheetByName(configuracion);
  var LastR = sheeta.getLastRow();
  var datos = sheeta.getDataRange().getValues();
  var longitud = sheeta.getRange("A2:A" + LastR).getValues();
  var sw = false;

  for (var i = 0; i <= longitud.length; i++) {



    let apellido1 = datos[i][1];
    let apellido2 = datos[i][2];
    let nom = datos[i][3];
    let correo = datos[i][4];
    let contrasena = datos[i][5];

    let nombre = nom + " " + apellido1 + " " + apellido2

    var obj = {

      nombre: nombre,
      correo: correo,
      contrasena: contrasena,
      sw: sw

    };



    if (correo != "" && user != "" && password != "" && contrasena != "") {
      //if (usuario.toUpperCase() == user.toUpperCase() && password == contrasena) {
      console.log(correo)
      console.log(user)
      console.log(password)
      console.log(contrasena)
      if (correo == user && password == contrasena) {

        sw = true;
        break;
      }
    } else {

      obj.sw = false;
    }
  }//Close for

  if (sw == true) {


    obj.sw = sw;

    return obj;

  } else {

    obj.sw = sw;
    return obj;
  }
}
//Close checkLogin



//Esta función envia los datos obtenidos por parametros y los guarda en la memoria cahce
function PutValues(email, name, password) {

  var cache = CacheService.getUserCache();
  // var cache = CacheService.getScriptCache();
  //Se guardan los datos en cache durante 4 horas
  cache.put("EMAIL", email, 14400);
  cache.put("NAME", name, 14400);
  cache.put("PASS", password, 14400);


}
//Close putValues



//Esta función obtiene los datos guardados en la memoria cahce
function ShowValues() {
  var cache = CacheService.getUserCache();
  //var cache = CacheService.getScriptCache();


  var email = cache.get("EMAIL");
  var name = cache.get("NAME");
  var pass = cache.get("PASS");

  Logger.log(email + " : " + name + " : " + pass);

  var objVal = {
    email: email,
    name: name,
    pass: pass,

  };
  return objVal;
}
//Close ShowValues




