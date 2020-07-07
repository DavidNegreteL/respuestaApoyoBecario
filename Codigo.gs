function enviar(e){
    var range = e.range;
    var check = range.getValue();
    var colum = range.getColumn();
    var row = range.getRow();
    
    if(check == 1 && colum == 182){
      var miSpread = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = miSpread.getActiveSheet();
      
      var cellEmail = sheet.getRange(row,2);
      var valEmail = cellEmail.getValues();
      var email = valEmail[0][0];
      
      var cellNombre = sheet.getRange(row,6);
      var valNombre = cellNombre.getValues();
      var nombre = valNombre[0][0];
      
      var cellApaterno = sheet.getRange(row,7);
      var valApaterno = cellApaterno.getValues();
      var aPaterno = valApaterno[0][0];
      
      var cellAmaterno = sheet.getRange(row,8);
      var valAmaterno = cellAmaterno.getValues();
      var aMaterno = valAmaterno[0][0];
      
      var fullName = `${aPaterno} ${aMaterno} ${nombre}`;
      
      var cellProximoGrado = sheet.getRange(row,20);
      var valProximoGrado = cellProximoGrado.getValues();
      var proximoGrado = valProximoGrado[0][0];
      
      var cellInscripcion = sheet.getRange(row,179);
      var valInscripcion = cellInscripcion.getValues();
      var montoInscripcion = valInscripcion[0][0];
      
      var cellMensualidad = sheet.getRange(row,180);
      var valMensualidad = cellMensualidad.getValues();
      var montoMensualidad = valMensualidad[0][0];
      
      var cellPeriodo = sheet.getRange(row,181);
      var valPeriodo = cellPeriodo.getValues();
      var periodo = valPeriodo[0][0];
      
      var fecha = new Date();
      var fechaParaDoc = parseFecha(fecha);
      
      const PDFFile = crearPDF(email,fullName,proximoGrado,montoInscripcion,montoMensualidad,periodo,fechaParaDoc);
      enviarRespuesta(email,PDFFile,fullName);
      }
      else{
      }
      /*Bloque de prueba para tomar datos de las celdas y visualizar en la spreadsheet
      =IMPORTRANGE("https://docs.google.com/spreadsheets/d/1HEcX7mBc1YxSNitUBeDqO0S1WmOX6bEdtRVjnGl4eDs/edit#gid=362266264","Respuestas de formulario 1!A1:FV100")
      miSpread.getRange('FW20').activate();
      miSpread.getCurrentCell().setValue(valInscripcion);*/
  }
  function crearPDF(email,fullName,proximoGrado,montoInscripcion,montoMensualidad,periodo,fechaParaDoc){
    const docFile = DriveApp.getFileById('1pMauijf1AnMYg_QHLYxsl3eBcS37sumxhvpRw27jfHQ');
    const docFolder = DriveApp.getFolderById('1pT_yoqSdapMSqoyrwk8GKtrU_n5En4NC');
    const pdfFolder = DriveApp.getFolderById('1O8xf1xxPPBgrkEJhtxt1CPvS7taHBt7N');
    const tempFile = docFile.makeCopy(docFolder);
    const tempDocFile = DocumentApp.openById(tempFile.getId());
    const body = tempDocFile.getBody();
    
    body.replaceText("{hoy}",fechaParaDoc);
    body.replaceText("{alumno}",fullName);
    body.replaceText("{grado}",proximoGrado);
    body.replaceText("{inscripcion}", montoInscripcion);
    body.replaceText("{mensualidad}", montoMensualidad);
    body.replaceText("{periodo}", periodo);
    tempDocFile.saveAndClose();
    
    const pdfContentBlob = tempFile.getAs(MimeType.PDF);
    const PDFFile = pdfFolder.createFile(pdfContentBlob).setName(fullName);
    return PDFFile;
  }
  
  function enviarRespuesta(email,PDFFile,fullName){
    var html_confirmacion = HtmlService.createTemplateFromFile('mail_confirmacion');
    var asunto = 'Respuesta a solicitud de Apoyo Becario para ' + fullName;
    html_confirmacion.nombreAlumno = fullName;
    var email_html = html_confirmacion.evaluate().getContent();
        
        MailApp.sendEmail(
          {
            to: email,
            subject: asunto,
            htmlBody: email_html,
            bcc: 'dnegrete@jegv.mx' + ',' + 'administracion@jegv.mx',
            name: 'Informática | JEGV',
            attachments: [PDFFile]
          }
        );
  }
  
  function parseFecha(fechaString){
    const dias = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado']
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
                   'Septiembre','Octubre','Noviembre','Diciembre']
    const fecha = new Date(fechaString)
    const nuevoFormatoFecha = `${dias[fecha.getDay()]} ${fecha.getDate()} de ${meses[fecha.getMonth()]} Ciudad de México`
    return nuevoFormatoFecha
  }
  