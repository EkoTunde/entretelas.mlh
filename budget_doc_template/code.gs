function onOpen(e) {
  DocumentApp.getUi()
    .createMenu("Más")
    .addSubMenu(
      DocumentApp.getUi()
        .createMenu("PDF")
        .addItem("Exportar como PDF", "convertPDF")
        .addItem("Exportar y enviar como PDF", "convertAndSend")
    )
    .addToUi();
}


function convertPDF() {

  var doc = DocumentApp.getActiveDocument();

  var docId = doc.getId();

  var ui = DocumentApp.getUi();
  var result = ui.alert(
    'Exportar',
    '¿Guardar documento como ' + doc.getName() + '.pdf)?',
    ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    var docblob = doc.getAs('application/pdf');
    /* Add the PDF extension */
    docblob.setName(doc.getName() + ".pdf");
    var file = DriveApp.createFile(docblob);
    // ADDED
    var fileId = file.getId();
    moveFileId(fileId);
    // ADDED
    ui.alert('Tu archivo PDF se encuentra disponible en ' + file.getUrl());
    return file.getName();
  } else {
    ui.alert('Se canceló la solicitud.');
    return null;
  }
}

function moveFileId(fileId) {
  var file = DriveApp.getFileById(fileId);
  var source_folder = DriveApp.getFileById(fileId).getParents().next();
  var folder = DriveApp.getFolderById("1DNLYjMLXBDjKkk4YYMVVK164akfgGai0")
  folder.addFile(file);
  source_folder.removeFile(file);
}

function convertAndSend() {

  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s");
  var db_orders = db.getSheetByName("DB Pedidos");
  var db_clients = db.getSheetByName("DB Clientes");

  var orders = db_orders.getDataRange().getValues();
  var URL_COL = 8;
  var budget_id;
  var client_id;
  var budget_expiration;
  for (var i = 0; i < orders.length; i++) {
    var row = orders[i];
    if (row[URL_COL] == DocumentApp.getActiveDocument().getUrl()) {
      budget_id = row[0];
      client_id = row[3];
      budget_expiration = row[9];
      break;
    }
  }

  var clients = db_clients.getDataRange().getValues();
  var client_name;
  var emailAdress;
  for (var i = 0; i < clients.length; i++) {
    var client = clients[i];
    if (client[0] == client_id) {
      if (client[7] === "") {
        alert("El cliente no tiene email.");
        return;
      }
      client_name = client[2];
      emailAdress = client[7];
      break;
    }
  }

  var fileName = convertPDF();
  if (fileName == null) return;

  sendEmail(emailAdress, budget_id, budget_expiration, fileName, client_name);
}

function sendEmail(emailAddress, budget_id, budget_expiration, pdfName, client_name) {
  var file = DriveApp.getFilesByName(pdfName);
  var subject = "Presupuesto Nro. " + budget_id;

  var template = HtmlService.createTemplateFromFile('email_msg');
  template.client_name = client_name;
  template.budget_id = budget_id;
  template.budget_expiration = Utilities.formatDate(budget_expiration, "GMT-3", "dd/MM/yyyy");
  var message = template.evaluate().getContent();

  var entretelas_img = DriveApp.getFilesByName("logo_template_contornos 100x100.png");

  var entretelasLogoBlob;

  if (entretelas_img.hasNext()) {
    entretelasLogoBlob = entretelas_img.next().getAs('image/png').setName("entretelasLogoBlob");
  }

  if (file.hasNext()) {
    GmailApp.sendEmail(emailAddress, subject, "", {
      attachments: [file.next().getAs(MimeType.PDF)],
      name: 'Entretelas MLH',
      htmlBody: message,
      inlineImages: { entretelasLogo: entretelasLogoBlob, }
    });
  }
}