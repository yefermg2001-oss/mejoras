var platilla = "1uwK7yw5_qezX1TXufpC8wun9m6USjTfc20v8UDY8V_A";
var destinoID = "1eGmykRU_OIlb4K9hPUfVSwEcTCXcOP9b"
var ssId = "1Q7KH2rEwvxJubf2UKf2zdJPvI8m2c093S_rxJr3juHY";
var ssDB = "Registro";
var ssDB2 = 'Formación';
var ssDB3 = 'Mapa de Conocimiento';
var ssDB4 = 'Experiencia';
var ssDB5 = 'Información Laboral';

function generarReporte(id) {
//.
 id = "1037648553"

  //Bases de datos
  var ss = SpreadsheetApp.openById(ssId);
  var regist = ss.getSheetByName(ssDB).getDataRange().getValues();
  var formac = ss.getSheetByName(ssDB2).getDataRange().getValues();
  var mapconocimiento = ss.getSheetByName(ssDB3).getDataRange().getValues();
  var experi = ss.getSheetByName(ssDB4).getDataRange().getValues();
  //var asistentes = ss.getSheetByName(asistentesDB).getDataRange().getValues();

  //Plantilla
  var nombre = id
  var template = DriveApp.getFileById(platilla);

  //Recorremos los formularios
  for (a = 1; a < regist.length; a++) {

    //Almacenamos los datos en variables
    var idForm = regist[a][3];

    //Si el id del formulario corresponde realizamos el informe
    if (id == idForm) {
      var cedula = regist[a][3];
      var fecha = regist[a][0];
      var nombre = regist[a][1];
      var estado = regist[a][22];
      var telefono = regist[a][21];
      var correo = regist[a][2];
      var preguntas = [];
      var preguntasmapco = [];
      var preguntasexpe = [];
      var np = 1;
      var mc= 1;
      var exp= 1;
      var fila = a + 1;

      //Carpeta
      var destino = DriveApp.getFolderById(destinoID)

      //Buscamos en el formac los las preguntas asociadas
      for (j = 1; j < formac.length; j++) {
        var fkForm = formac[j][1];

        //Si coinciden las llaves almacenamos los datos de la pregunta en un array
        if (idForm == fkForm) {
          preguntas.push([np, formac[j][2], formac[j][3], formac[j][4], formac[j][5],formac[j][6]]);
          np += 1
        }
      }
      for (j = 1; j < mapconocimiento.length; j++) {
        var fkForm = mapconocimiento[j][0];

        //Si coinciden las llaves almacenamos los datos de la pregunta en un array
        if (idForm == fkForm) {
          preguntasmapco.push([mc, mapconocimiento[j][1], mapconocimiento[j][2], mapconocimiento[j][3]]);
          mc += 1
        }
      }
       for (j = 1; j < experi.length; j++) {
        var fkForm = experi[j][1];

        //Si coinciden las llaves almacenamos los datos de la pregunta en un array
        if (idForm == fkForm) {
          preguntasexpe.push([exp,experi[j][1], experi[j][2], experi[j][3], experi[j][4], experi[j][5], experi[j][6], experi[j][7], experi[j][10]]);
          exp += 1
        }
      }
      //Buscamos en la lista de asistentes asociadas
    /*  for (j = 1; j < asistentes.length; j++) {
        var fkForm = asistentes[j][1];

        //Si coinciden las llaves almacenamos los datos de la pregunta en un array
        if (idForm == fkForm) {
          listaAsistentes.push([na, asistentes[j][2], asistentes[j][3], asistentes[j][4], asistentes[j][5], asistentes[j][6]]);
          na += 1
        }
      } */


      //Creamos una copia de la plantilla
      var copia = template.makeCopy(nombre, destino)
      var copyUrl = copia.getUrl();
      var copyID = copia.getId();
      var reporte = SpreadsheetApp.openById(copyID);

      //Almacenamos la url en la base de datos
      var celda = ss.getSheetByName(ssDB).getRange("AE" + fila)
      var richValue = SpreadsheetApp.newRichTextValue()
        .setText("Ver Reporte")
        .setLinkUrl(copyUrl)
        .build();
      celda.setRichTextValue(richValue);


      //LLevamos los datos generales
      reporte.getRange("F2").setValue(nombre);
      reporte.getRange("B10").setValue(estado);
      reporte.getRange("B13").setValue(nombre);
      reporte.getRange("B16").setValue(id);
      reporte.getRange("B19").setValue(telefono);
      reporte.getRange("B22").setValue(correo);
      
    
      //llevamos los datos de los preguntas
      for (p = 0; p < preguntas.length; p++) {
        reporte.getRange("E" + (9 + p)).setValue(preguntas[p][2])
        reporte.getRange("G" + (9 + p)).setValue(preguntas[p][3])
        reporte.getRange("I" + (9 + p)).setValue(preguntas[p][4])
        reporte.getRange("J" + (9 + p)).setValue(preguntas[p][5])
      }

         //llevamos los datos de los preguntas
      for (p = 0; p < preguntasmapco.length; p++) {
        reporte.getRange("E" + (29 + p)).setValue(preguntasmapco[p][1])
        reporte.getRange("G" + (29 + p)).setValue(preguntasmapco[p][2])
        reporte.getRange("I" + (29 + p)).setValue(preguntasmapco[p][3])
      }
          //llevamos los datos de los preguntas
      for (p = 0; p < preguntasexpe.length; p++) {
        reporte.getRange("B" + (69 + p)).setValue(preguntasexpe[p][2])
        reporte.getRange("D" + (69 + p)).setValue(preguntasexpe[p][3])
        reporte.getRange("E" + (69 + p)).setValue(preguntasexpe[p][4])
        reporte.getRange("F" + (69 + p)).setValue(preguntasexpe[p][5])
        reporte.getRange("H" + (69 + p)).setValue(preguntasexpe[p][6])
        reporte.getRange("I" + (69 + p)).setValue(preguntasexpe[p][7])
        reporte.getRange("k" + (69 + p)).setValue(preguntasexpe[p][8])
      
      }

     // reporte.getRange("A37:F" + (36 + listaAsistentes.length)).setValues(listaAsistentes)

      //Igresamos las imagenes y cambiamos el tamaño de las celdas
   //   let image1 =  SpreadsheetApp.newCellImage().setSourceUrl(foto1).build();
     // let image2 = SpreadsheetApp.newCellImage().setSourceUrl(foto2).build();
      //let image4 = SpreadsheetApp.newCellImage().setSourceUrl(firma_Enc).build();
     // let image5 = SpreadsheetApp.newCellImage().setSourceUrl(firma_int).build();
      
      //LLevamos las imagenes a la hoja de calculo
    //  try {
      //  reporte.getRange("F12").setValue(image1);
     // }
     // catch (e) {
       // console.log(e)
     // }
     // try {
       // reporte.getRange("F29").setValue(image2);
      //}
     // catch (e) {
       // console.log(e)
     // }
      
    //  try {
     
//   reporte.getRange("A49").setValue(image4);
  ////  catch (e) {
      //  console.log(e)
      //}
     // try {
       // reporte.getRange("D49").setValue(image5);
      //}
      //catch (e) {
       // console.log(e)
      //}

      //Modificamos el alto de la celda
      //reporte.setRowHeight(73, 170);

    }
  }
}