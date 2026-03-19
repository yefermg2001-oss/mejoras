/**
 * =============================================
 * GENERADOR AUTOMÁTICO DE HOJAS DE VIDA
 * =============================================
 * 
 * Genera una hoja de vida en Google Docs a partir de la plantilla,
 * llenándola con los datos de Registro, Formación y Experiencia.
 * Copia los certificados a la carpeta del candidato.
 */

var PLANTILLA_HV_ID = "1-ilRj2bjR4jK3IfHXwvUb9J5aetk-EKY9JWDvfEOpqk";
var CARPETA_HV_ID = "1G7oYj2_nEhQugM0QaICz-VCml9guLKF5";

/**
 * @function generarHojaDeVida
 * Genera la hoja de vida completa de una persona.
 * @param {string} cedula - Cédula de la persona
 * @returns {string} JSON con { success: true, url: "..." } o { success: false, error: "..." }
 */
function generarHojaDeVida(cedula) {
  try {
    // ↓↓ VERIFICACIÓN DE SEGURIDAD — Solo admins ↓↓
    if (!_esAdmin()) {
      Logger.log("SEGURIDAD: Usuario no autorizado intentó generar HV para cédula " + cedula);
      return JSON.stringify({ success: false, error: "No tiene permisos para generar hojas de vida" });
    }
    // ↑↑ FIN VERIFICACIÓN ↑↑

    cedula = cedula.toString().trim();
    var ss = SpreadsheetApp.openById(ssId);

    // =============================================
    // 1. LEER DATOS
    // =============================================

    // Registro - con búsqueda dinámica de columnas por nombre de encabezado
    var dataReg = ss.getSheetByName('Registro').getDataRange().getValues();
    var regHeaders = dataReg[0];
    var regCedulaCol = _findCol(regHeaders, "Cédula");
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

    // Mapear columnas del Registro por nombre de encabezado
    var _regCol = function(nombre) {
      for (var i = 0; i < regHeaders.length; i++) {
        if (regHeaders[i].toString().trim().toLowerCase().indexOf(nombre.toLowerCase()) >= 0) return i;
      }
      return -1;
    };

    // Log para debug: imprimir todos los encabezados del Registro
    Logger.log("=== ENCABEZADOS REGISTRO ===");
    for (var h = 0; h < regHeaders.length; h++) {
      Logger.log("Col " + h + ": " + regHeaders[h]);
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

    // Obtener nombre dinámicamente
    var colNombre = _regCol("nombre");
    var nombrePersona = colNombre >= 0 ? (personaRow[colNombre] || "Sin nombre").toString().trim() : "Sin nombre";

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
    // 3. LLENAR DATOS PERSONALES (búsqueda dinámica por encabezado)
    // =============================================

    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();

    // Actualizar fecha en el encabezado del documento
    var header = doc.getHeader();
    if (header) {
      header.replaceText("FECHA:.*", "FECHA: " + _formatFecha(new Date()));
    }

    body.replaceText("\\{\\{NOMBRE\\}\\}", nombrePersona);
    body.replaceText("\\{\\{CEDULA\\}\\}", cedula);

    // Ciudad y Fecha de Nacimiento — combinar lugar + fecha de cédula
    var colCedulaLugar = _regCol("cedulaLugar");
    var colCedulaFecha = _regCol("cedulaFecha");
    var ciudadNac = colCedulaLugar >= 0 ? (personaRow[colCedulaLugar] || "").toString() : "";
    var fechaNac = colCedulaFecha >= 0 ? _formatFecha(personaRow[colCedulaFecha]) : "";
    var ciudadYFecha = ciudadNac + (fechaNac ? ", " + fechaNac : "");
    body.replaceText("\\{\\{CIUDAD_NAC\\}\\}", ciudadYFecha);

    // Correo
    var colCorreo = _regCol("correo");
    body.replaceText("\\{\\{CORREO\\}\\}", colCorreo >= 0 ? (personaRow[colCorreo] || "").toString() : "");

    // Teléfono — primero buscar en Vivienda, luego en Registro
    var dbViv = ss.getSheetByName('Vivienda');
    var viviendaRow = null;
    if (dbViv) {
      var dataViv = dbViv.getDataRange().getValues();
      var vivCedulaCol = _findCol(dataViv[0], "Cédula");
      for (var i = 1; i < dataViv.length; i++) {
        if (dataViv[i][vivCedulaCol].toString().trim() == cedula) {
          viviendaRow = dataViv[i];
          break;
        }
      }
    }

    // Teléfono: Vivienda col 8 (telefono) o col 9 (celular), o Registro como fallback
    var telefono = "";
    if (viviendaRow) {
      telefono = (viviendaRow[9] || viviendaRow[8] || "").toString().trim();
    }
    if (!telefono) {
      var colTel = _regCol("tel");
      if (colTel < 0) colTel = _regCol("celular");
      if (colTel >= 0) telefono = (personaRow[colTel] || "").toString().trim();
    }
    body.replaceText("\\{\\{TELEFONO\\}\\}", telefono);

    // Municipio y Departamento — desde Vivienda (ciudad col 6, departamento col 5)
    var municipio = "";
    var departamento = "";
    if (viviendaRow) {
      municipio = (viviendaRow[6] || "").toString();
      departamento = (viviendaRow[5] || "").toString();
    } else {
      // Fallback: buscar en Registro
      var colMunicipio = _regCol("lugarresidencia");
      if (colMunicipio >= 0) municipio = (personaRow[colMunicipio] || "").toString();
    }
    body.replaceText("\\{\\{MUNICIPIO\\}\\}", municipio);
    body.replaceText("\\{\\{DEPARTAMENTO\\}\\}", departamento);

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
    var tablas = body.getTables();
    
    // Buscar la tabla de datos personales por contenido (no por índice)
    var tablaDatos = null;
    for (var t = 0; t < tablas.length; t++) {
      try {
        var cellText = tablas[t].getCell(0, 0).getText();
        if (cellText.indexOf("DATOS PERSONALES") >= 0) {
          tablaDatos = tablas[t];
          break;
        }
      } catch (e) { }
    }
    
    // Marcar nivel de educación con X
    if (tablaDatos) {
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
          tablaDatos.getCell(2, 10).setText("X");
        }
      } catch (e) {
        Logger.log("Error al marcar nivel educación: " + e.toString());
      }
    }

    // =============================================
    // 5. FORMACIÓN ACADÉMICA + INFO LABORAL
    //    ENFOQUE: LLENAR FILAS EXISTENTES + BORRAR SOBRANTES
    // =============================================

    if (tablaDatos) {
      var tablaGrande = tablaDatos;
      
      // Buscar la primera fila de datos (la que tiene los marcadores o la primera fila vacía después de los encabezados)
      var filaInicioData = -1;
      for (var r = 0; r < tablaGrande.getNumRows(); r++) {
        try {
          var cellText = tablaGrande.getCell(r, 0).getText();
          if (cellText.indexOf("{{F_UNIVERSIDAD}}") >= 0 || cellText.indexOf("F_UNIVERSIDAD") >= 0) {
            filaInicioData = r;
            break;
          }
        } catch (e) { }
      }
      if (filaInicioData < 0) filaInicioData = tablaGrande.getNumRows() - 1;

      // Contar cuántas filas de datos hay disponibles (desde filaInicioData hasta el final)
      var filasDisponibles = tablaGrande.getNumRows() - filaInicioData;
      var maxFilas = Math.max(formaciones.length, experiencias.length);

      // Llenar las filas existentes con formación (lado izquierdo)
      for (var f = 0; f < formaciones.length && f < filasDisponibles; f++) {
        _llenarFilaFormacion(tablaGrande, filaInicioData + f, formaciones[f]);
      }

      // Llenar las filas existentes con info laboral (lado derecho)
      for (var e = 0; e < experiencias.length && e < filasDisponibles; e++) {
        _llenarFilaInfoLaboral(tablaGrande, filaInicioData + e, experiencias[e]);
      }

      // Borrar filas sobrantes (de abajo hacia arriba para no desplazar índices)
      var totalFilas = tablaGrande.getNumRows();
      for (var r = totalFilas - 1; r >= filaInicioData + maxFilas; r--) {
        try {
          tablaGrande.removeRow(r);
        } catch (e) {
          Logger.log("Error borrando fila sobrante " + r + ": " + e.toString());
        }
      }

      // Limpiar marcadores restantes en filas que no se llenaron
      for (var r = filaInicioData; r < tablaGrande.getNumRows(); r++) {
        try {
          for (var c = 0; c < tablaGrande.getRow(r).getNumCells(); c++) {
            var txt = tablaGrande.getCell(r, c).getText();
            if (txt.indexOf("{{") >= 0) {
              tablaGrande.getCell(r, c).setText("");
            }
          }
        } catch (e) { }
      }
    }

    // =============================================
    // 6. EXPERIENCIA DETALLADA (Objeto + Funciones)
    //    ENFOQUE: LLENAR FILAS EXISTENTES + BORRAR SOBRANTES
    // =============================================

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
      // 1 FILA por experiencia (todo con saltos de línea dentro de cada celda)
      var filaInicioExp = 2;
      var filasDisponiblesExp = tablaExp.getNumRows() - filaInicioExp;

      // Llenar las experiencias en las filas existentes
      for (var e = 0; e < experiencias.length && e < filasDisponiblesExp; e++) {
        _llenarExpDetalle(tablaExp, filaInicioExp + e, experiencias[e]);
      }

      // Borrar filas sobrantes (de abajo hacia arriba)
      var totalFilasExp = tablaExp.getNumRows();
      for (var r = totalFilasExp - 1; r >= filaInicioExp + experiencias.length; r--) {
        try {
          tablaExp.removeRow(r);
        } catch (e) {
          Logger.log("Error borrando fila exp sobrante " + r + ": " + e.toString());
        }
      }

      // Limpiar marcadores sobrantes
      for (var r = filaInicioExp; r < tablaExp.getNumRows(); r++) {
        try {
          for (var c = 0; c < tablaExp.getRow(r).getNumCells(); c++) {
            var txt = tablaExp.getCell(r, c).getText();
            if (txt.indexOf("{{") >= 0) {
              tablaExp.getCell(r, c).setText("");
            }
          }
        } catch (e) { }
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
  // Manejar números seriales de fecha (ej: 37219 = fecha de Excel/Sheets)
  if (typeof valor === 'number' && valor > 10000 && valor < 100000) {
    var fecha = new Date((valor - 25569) * 86400 * 1000);
    var d = fecha.getDate().toString().padStart(2, '0');
    var m = (fecha.getMonth() + 1).toString().padStart(2, '0');
    var y = fecha.getFullYear();
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
function _llenarExpDetalle(tabla, rowIdx, experiencia) {
  try {
    var row = tabla.getRow(rowIdx);
    
    // Columna 0: Objeto + Empresa — TODO en negrilla
    var cell0 = row.getCell(0);
    cell0.clear();
    var p0 = cell0.appendParagraph("");
    p0.appendText(experiencia.objeto + "\n" + experiencia.empresa).setBold(true);
    
    // Columna 1: Cargo, Dedicación, Funciones con formato mixto
    var cell1 = row.getCell(1);
    cell1.clear();
    
    // CARGO: todo en negrilla
    var pCargo = cell1.appendParagraph("");
    pCargo.appendText("CARGO: " + experiencia.cargo).setBold(true);
    
    // % DEDICACIÓN: todo en negrilla
    var pDedic = cell1.appendParagraph("");
    pDedic.appendText("% DEDICACIÓN: " + _formatDedicacion(experiencia.dedicacion)).setBold(true);
    
    // FUNCIONES: título en negrilla, texto sin negrilla
    var pFunc = cell1.appendParagraph("");
    pFunc.appendText("FUNCIONES: ").setBold(true);
    pFunc.appendText(experiencia.funciones).setBold(false);
    
    // Columna 2: Fecha inicio — sin negrilla
    var cell2 = row.getCell(2);
    cell2.clear();
    cell2.appendParagraph(experiencia.fechaIn);
    
    // Columna 3: Fecha fin — sin negrilla
    var cell3 = row.getCell(3);
    cell3.clear();
    cell3.appendParagraph(experiencia.fechaFin);

  } catch (e) {
    Logger.log("Error llenando exp detalle fila " + rowIdx + ": " + e.toString());
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
