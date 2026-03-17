/**
 * =============================================
 * GENERADOR AUTOMÁTICO DE HOJAS DE VIDA
 * =============================================
 * 
 * Genera una hoja de vida en Google Docs a partir de la plantilla,
 * llenándola con los datos de Registro, Formación y Experiencia.
 * Copia los certificados a la carpeta del candidato.
 */

var PLANTILLA_HV_ID = "1DiqKaHYPPPfugr1bx7qTgcUSB7gcAToD";
var CARPETA_HV_ID = "1G7oYj2_nEhQugM0QaICz-VCml9guLKF5";

/**
 * @function generarHojaDeVida
 * Genera la hoja de vida completa de una persona.
 * @param {string} cedula - Cédula de la persona
 * @returns {string} JSON con { success: true, url: "..." } o { success: false, error: "..." }
 */
function generarHojaDeVida(cedula) {
  try {
    cedula = cedula.toString().trim();
    var ss = SpreadsheetApp.openById(ssId);

    // =============================================
    // 1. LEER DATOS
    // =============================================

    // Registro
    var dataReg = ss.getSheetByName('Registro').getDataRange().getValues();
    var regCedulaCol = _findCol(dataReg[0], "Cédula");
    var personaRow = null;
    for (var i = 1; i < dataReg.length; i++) {
      if (dataReg[i][regCedulaCol].toString().trim() == cedula) {
        personaRow = dataReg[i];
        break;
      }
    }
    if (!personaRow) {
      return JSON.stringify({ success: false, error: "No se encontró la persona con cédula " + cedula });
    }

    // Formación
    var dataForm = ss.getSheetByName('Formación').getDataRange().getValues();
    var formCedulaCol = _findCol(dataForm[0], "Cédula");
    var formaciones = [];
    for (var i = 1; i < dataForm.length; i++) {
      if (dataForm[i][formCedulaCol].toString().trim() == cedula) {
        formaciones.push({
          nivel: (dataForm[i][2] || "").toString(),
          universidad: (dataForm[i][3] || "").toString(),
          titulo: (dataForm[i][4] || "").toString(),
          estado: (dataForm[i][5] || "").toString(),
          fechaGrado: _formatFecha(dataForm[i][6]),
          adjunto: (dataForm[i][7] || "").toString(),
          tarjeta: (dataForm[i][8] || "").toString(),
          numTarjeta: (dataForm[i][9] || "").toString(),
          seccional: (dataForm[i][10] || "").toString(),
          expedicion: _formatFecha(dataForm[i][11])
        });
      }
    }

    // Experiencia
    var dataExp = ss.getSheetByName('Experiencia').getDataRange().getValues();
    var expCedulaCol = _findCol(dataExp[0], "Cédula");
    var experiencias = [];
    for (var i = 1; i < dataExp.length; i++) {
      if (dataExp[i][expCedulaCol].toString().trim() == cedula) {
        experiencias.push({
          empresa: (dataExp[i][2] || "").toString(),
          cargo: (dataExp[i][3] || "").toString(),
          fechaIn: _formatFecha(dataExp[i][4]),
          fechaFin: _formatFechaOActual(dataExp[i][5]),
          objeto: (dataExp[i][6] || "").toString(),
          funciones: (dataExp[i][7] || "").toString(),
          adjunto: (dataExp[i][8] || "").toString(),
          sector: (dataExp[i][9] || "").toString(),
          clasificacion: (dataExp[i][10] || "").toString(),
          tipoproy: (dataExp[i][11] || "").toString(),
          dedicacion: (dataExp[i][12] || "").toString()
        });
      }
    }

    // Información Laboral
    var dbInfoLab = ss.getSheetByName('Información Laboral');
    var infoLaboral = [];
    if (dbInfoLab) {
      var dataIL = dbInfoLab.getDataRange().getValues();
      var ilCedulaCol = _findCol(dataIL[0], "Cédula");
      for (var i = 1; i < dataIL.length; i++) {
        if (dataIL[i][ilCedulaCol].toString().trim() == cedula) {
          infoLaboral.push(dataIL[i]);
        }
      }
    }

    var nombrePersona = (personaRow[1] || "Sin nombre").toString().trim();

    // =============================================
    // 2. CREAR CARPETA Y COPIAR PLANTILLA
    // =============================================

    var carpetaPadre = DriveApp.getFolderById(CARPETA_HV_ID);
    var nombreCarpeta = "HV_" + nombrePersona.replace(/\s+/g, '_') + "_" + cedula;

    // Buscar si ya existe la carpeta
    var carpetaPersona = null;
    var carpetas = carpetaPadre.getFoldersByName(nombreCarpeta);
    if (carpetas.hasNext()) {
      carpetaPersona = carpetas.next();
      // Vaciar la carpeta existente
      var archivos = carpetaPersona.getFiles();
      while (archivos.hasNext()) {
        archivos.next().setTrashed(true);
      }
    } else {
      carpetaPersona = carpetaPadre.createFolder(nombreCarpeta);
    }

    // Copiar la plantilla
    var plantilla = DriveApp.getFileById(PLANTILLA_HV_ID);
    var copia = plantilla.makeCopy("Hoja_de_Vida_" + nombrePersona, carpetaPersona);
    var docId = copia.getId();

    // =============================================
    // 3. LLENAR DATOS PERSONALES (replaceText)
    // =============================================

    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();

    body.replaceText("\\{\\{NOMBRE\\}\\}", nombrePersona);
    body.replaceText("\\{\\{CEDULA\\}\\}", cedula);
    body.replaceText("\\{\\{CIUDAD_NAC\\}\\}", (personaRow[4] || "").toString());
    body.replaceText("\\{\\{CORREO\\}\\}", (personaRow[2] || "").toString());
    body.replaceText("\\{\\{TELEFONO\\}\\}", (personaRow[20] || "").toString().trim());
    body.replaceText("\\{\\{MUNICIPIO\\}\\}", (personaRow[8] || "").toString());
    body.replaceText("\\{\\{DEPARTAMENTO\\}\\}", "");

    // Fecha de vinculación: experiencia más antigua con SEDIC
    var fechaVinc = "";
    experiencias.forEach(function (exp) {
      if (exp.empresa.toUpperCase().indexOf("SEDIC") >= 0) {
        if (!fechaVinc || exp.fechaIn < fechaVinc) {
          fechaVinc = exp.fechaIn;
        }
      }
    });
    body.replaceText("\\{\\{FECHA_VINCULACION\\}\\}", fechaVinc);

    // Matrícula: buscar en formaciones la primera que tenga tarjeta profesional
    var matricula = "", seccional = "", fechaMatricula = "";
    for (var f = 0; f < formaciones.length; f++) {
      if (formaciones[f].numTarjeta) {
        matricula = formaciones[f].numTarjeta;
        seccional = formaciones[f].seccional;
        fechaMatricula = formaciones[f].expedicion;
        break;
      }
    }
    body.replaceText("\\{\\{MATRICULA\\}\\}", matricula);
    body.replaceText("\\{\\{SECCIONAL\\}\\}", seccional);
    body.replaceText("\\{\\{FECHA_MATRICULA\\}\\}", fechaMatricula);

    // =============================================
    // 4. NIVEL DE EDUCACIÓN (marcar con X)
    // =============================================

    var nivelMax = _getNivelMaximo(formaciones);
    // Reemplazar la celda correspondiente con X
    var tablas = body.getTables();
    if (tablas.length > 1) {
      var tablaDatos = tablas[1]; // Table 1 = datos personales
      // Row 1: PROFESIONAL (col 7), TECNOLOGO (col 9), TÉCNICO (col 12)
      // Row 2: BACHILLER (col 7), NINGUNA (col 9)
      try {
        if (nivelMax === "PROFESIONAL" || nivelMax === "MAESTRÍA" || nivelMax === "ESPECIALIZACIÓN" ||
          nivelMax === "DOCTORADO" || nivelMax === "POSGRADO") {
          tablaDatos.getCell(1, 8).setText("X");
        } else if (nivelMax === "TECNÓLOGO" || nivelMax === "TECNOLÓGICO") {
          tablaDatos.getCell(1, 10).setText("X");
        } else if (nivelMax === "TÉCNICO") {
          tablaDatos.getCell(1, 12).setText("X");
        } else if (nivelMax === "BACHILLER") {
          tablaDatos.getCell(2, 8).setText("X");
        } else {
          tablaDatos.getCell(2, 10).setText("X"); // Ninguna de las anteriores
        }
      } catch (e) {
        Logger.log("Error al marcar nivel educación: " + e.toString());
      }
    }

    // =============================================
    // 5. FORMACIÓN ACADÉMICA + INFO LABORAL (Table 1, Row 10)
    // =============================================

    if (tablas.length > 1) {
      var tablaGrande = tablas[1];
      var filaTemplate = 10; // Row 10 es la fila con los marcadores

      // Llenar la primera fila con la primera formación (si hay)
      if (formaciones.length > 0) {
        _llenarFilaFormacion(tablaGrande, filaTemplate, formaciones[0]);
      }

      // Para formaciones adicionales, insertar nuevas filas
      for (var f = 1; f < formaciones.length; f++) {
        var newRow = tablaGrande.insertTableRow(filaTemplate + f);
        // Copiar estructura de celdas de la fila template
        var templateRow = tablaGrande.getRow(filaTemplate);
        for (var c = 0; c < templateRow.getNumCells(); c++) {
          if (c >= newRow.getNumCells()) {
            newRow.appendTableCell("");
          }
        }
        _llenarFilaFormacion(tablaGrande, filaTemplate + f, formaciones[f]);
      }

      // Llenar la primera fila de info laboral (experiencias, lado derecho)
      if (experiencias.length > 0) {
        _llenarFilaInfoLaboral(tablaGrande, filaTemplate, experiencias[0]);
      }
      for (var e = 1; e < experiencias.length && e < formaciones.length; e++) {
        _llenarFilaInfoLaboral(tablaGrande, filaTemplate + e, experiencias[e]);
      }
      // Si hay más experiencias que formaciones, agregar filas adicionales
      for (var e = Math.max(formaciones.length, 1); e < experiencias.length; e++) {
        var rowIdx = filaTemplate + e;
        if (rowIdx >= tablaGrande.getNumRows()) {
          var newRow = tablaGrande.appendTableRow();
          for (var c = 0; c < tablaGrande.getRow(filaTemplate).getNumCells(); c++) {
            if (c >= newRow.getNumCells()) {
              newRow.appendTableCell("");
            }
          }
        }
        _llenarFilaInfoLaboral(tablaGrande, rowIdx, experiencias[e]);
      }
    }

    // =============================================
    // 6. EXPERIENCIA DETALLADA (Table 3: Objeto + Funciones)
    // =============================================

    // Buscar la tabla de objeto/funciones
    var tablaExp = null;
    for (var t = 0; t < tablas.length; t++) {
      try {
        var firstCell = tablas[t].getCell(0, 0).getText();
        if (firstCell.indexOf("OBJETO DEL CONTRATO") >= 0) {
          tablaExp = tablas[t];
          break;
        }
      } catch (e) { }
    }

    if (tablaExp && experiencias.length > 0) {
      // La plantilla tiene filas 2, 3, 4 como template (3 filas por experiencia)
      // Fila 2: Objeto + CARGO
      // Fila 3: Objeto + % DEDICACIÓN
      // Fila 4: Objeto + FUNCIONES

      // Llenar primera experiencia en las filas template
      _llenarExpDetalle(tablaExp, 2, experiencias[0]);

      // Para experiencias adicionales, insertar 3 filas por cada una
      for (var e = 1; e < experiencias.length; e++) {
        var baseRow = 2 + (e * 3);
        for (var r = 0; r < 3; r++) {
          var newRow = tablaExp.insertTableRow(baseRow + r);
          for (var c = 0; c < 4; c++) {
            if (c >= newRow.getNumCells()) {
              newRow.appendTableCell("");
            }
          }
        }
        _llenarExpDetalle(tablaExp, baseRow, experiencias[e]);
      }
    }

    // =============================================
    // 7. COPIAR CERTIFICADOS
    // =============================================

    // Certificados de formación
    for (var f = 0; f < formaciones.length; f++) {
      _copiarCertificado(formaciones[f].adjunto, carpetaPersona,
        "Cert_Formacion_" + (f + 1) + "_" + formaciones[f].titulo);
    }

    // Certificados de experiencia
    for (var e = 0; e < experiencias.length; e++) {
      _copiarCertificado(experiencias[e].adjunto, carpetaPersona,
        "Cert_Experiencia_" + (e + 1) + "_" + experiencias[e].empresa);
    }

    // =============================================
    // 8. GUARDAR Y RETORNAR
    // =============================================

    doc.saveAndClose();

    // Exportar como PDF también
    var pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
    pdfBlob.setName("Hoja_de_Vida_" + nombrePersona + ".pdf");
    carpetaPersona.createFile(pdfBlob);

    return JSON.stringify({
      success: true,
      url: carpetaPersona.getUrl(),
      docUrl: copia.getUrl(),
      nombre: nombrePersona
    });

  } catch (error) {
    Logger.log("Error generarHojaDeVida: " + error.toString());
    return JSON.stringify({
      success: false,
      error: error.toString()
    });
  }
}


// =============================================
// FUNCIONES AUXILIARES
// =============================================

/**
 * Busca la columna con el nombre indicado
 */
function _findCol(headerRow, nombre) {
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i].toString().trim() == nombre) return i;
  }
  return 0;
}

/**
 * Formatea una fecha a DD/MM/YYYY
 */
function _formatFecha(valor) {
  if (!valor) return "";
  if (valor instanceof Date) {
    var d = valor.getDate().toString().padStart(2, '0');
    var m = (valor.getMonth() + 1).toString().padStart(2, '0');
    var y = valor.getFullYear();
    return d + "/" + m + "/" + y;
  }
  return valor.toString();
}

/**
 * Formatea una fecha o devuelve "La fecha" si está vacía (experiencia actual)
 */
function _formatFechaOActual(valor) {
  if (!valor || valor.toString().trim() == "") return "La fecha";
  return _formatFecha(valor);
}

/**
 * Determina el nivel máximo de educación
 */
function _getNivelMaximo(formaciones) {
  var jerarquia = {
    "DOCTORADO": 7,
    "MAESTRÍA": 6,
    "POSGRADO": 5,
    "ESPECIALIZACIÓN": 5,
    "PROFESIONAL": 4,
    "TECNÓLOGO": 3,
    "TECNOLÓGICO": 3,
    "TÉCNICO": 2,
    "BACHILLER": 1
  };

  var maxNivel = "";
  var maxValor = 0;
  formaciones.forEach(function (f) {
    var nivel = f.nivel.toUpperCase().trim();
    var valor = jerarquia[nivel] || 0;
    if (valor > maxValor) {
      maxValor = valor;
      maxNivel = nivel;
    }
  });
  return maxNivel;
}

/**
 * Llena una fila de la tabla con datos de formación (lado izquierdo)
 */
function _llenarFilaFormacion(tabla, rowIdx, formacion) {
  try {
    var row = tabla.getRow(rowIdx);
    row.getCell(0).setText(formacion.universidad);  // INSTITUCIÓN
    row.getCell(4).setText(formacion.fechaGrado);    // FECHA DE GRADO
    row.getCell(5).setText("");                       // DURACIÓN (calculada)
    row.getCell(6).setText(formacion.titulo);         // TÍTULO
  } catch (e) {
    Logger.log("Error llenando formación fila " + rowIdx + ": " + e.toString());
  }
}

/**
 * Llena una fila de la tabla con datos de info laboral (lado derecho)
 */
function _llenarFilaInfoLaboral(tabla, rowIdx, experiencia) {
  try {
    var row = tabla.getRow(rowIdx);
    var empresaCargo = experiencia.empresa + " / " + experiencia.cargo;
    row.getCell(7).setText(empresaCargo);      // EMPRESA / CARGO
    row.getCell(9).setText(experiencia.fechaIn); // DE
    row.getCell(12).setText(experiencia.fechaFin); // A
  } catch (e) {
    Logger.log("Error llenando info laboral fila " + rowIdx + ": " + e.toString());
  }
}

/**
 * Llena 3 filas de la tabla de experiencia detallada
 * Fila 0: Objeto + CARGO
 * Fila 1: Objeto + % DEDICACIÓN
 * Fila 2: Objeto + FUNCIONES
 */
function _llenarExpDetalle(tabla, baseRow, experiencia) {
  try {
    var objTexto = experiencia.objeto + "\n" + experiencia.empresa;

    // Fila 1: Cargo
    var row1 = tabla.getRow(baseRow);
    row1.getCell(0).setText(objTexto);
    row1.getCell(1).setText("CARGO: " + experiencia.cargo);
    row1.getCell(2).setText(experiencia.fechaIn);
    row1.getCell(3).setText(experiencia.fechaFin);

    // Fila 2: Dedicación
    var row2 = tabla.getRow(baseRow + 1);
    row2.getCell(0).setText(objTexto);
    row2.getCell(1).setText("% DEDICACIÓN: " + _formatDedicacion(experiencia.dedicacion));
    row2.getCell(2).setText(experiencia.fechaIn);
    row2.getCell(3).setText(experiencia.fechaFin);

    // Fila 3: Funciones
    var row3 = tabla.getRow(baseRow + 2);
    row3.getCell(0).setText(objTexto);
    row3.getCell(1).setText("FUNCIONES: " + experiencia.funciones);
    row3.getCell(2).setText(experiencia.fechaIn);
    row3.getCell(3).setText(experiencia.fechaFin);

  } catch (e) {
    Logger.log("Error llenando exp detalle fila " + baseRow + ": " + e.toString());
  }
}

/**
 * Formatea dedicación como porcentaje
 */
function _formatDedicacion(valor) {
  if (!valor) return "100%";
  var str = valor.toString().trim();
  if (str.indexOf("%") >= 0) return str;
  var num = parseFloat(str);
  if (isNaN(num)) return "100%";
  if (num <= 1) return Math.round(num * 100) + "%";
  return Math.round(num) + "%";
}

/**
 * Copia un certificado desde su URL al folder destino
 */
function _copiarCertificado(url, carpetaDestino, nombreBase) {
  if (!url || url.toString().trim() == "") return;
  try {
    var fileId = _extraerFileId(url);
    if (!fileId) return;
    var archivo = DriveApp.getFileById(fileId);
    // Limpiar nombre
    var nombreLimpio = nombreBase.replace(/[^a-zA-Z0-9áéíóúñÁÉÍÓÚÑ\s_\-]/g, '').substring(0, 80);
    var extension = archivo.getName().split('.').pop();
    archivo.makeCopy(nombreLimpio + "." + extension, carpetaDestino);
  } catch (e) {
    Logger.log("No se pudo copiar certificado: " + url + " - " + e.toString());
  }
}

/**
 * Extrae el ID de un archivo de Google Drive desde su URL
 */
function _extraerFileId(url) {
  if (!url) return null;
  url = url.toString();
  // Formato: https://drive.google.com/file/d/XXXXX/view
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // Formato: ?id=XXXXX
  match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // Si es solo un ID
  if (url.match(/^[a-zA-Z0-9_-]{10,}$/)) return url;
  return null;
}
