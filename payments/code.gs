function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Más')
    .addItem('Eliminar', 'deletePayment')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Configurar").addItem("Configurar ahora", "installOnSelected"))
    .addToUi();
}

function deletePayment() {
  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  var selectedId = uiForm.getRange("J1").getValue();

  if (selectedId == "") {
    SpreadsheetApp.getUi().alert("Ningún pago seleccionado.");
    return;
  }

  var values = uiForm.getDataRange().getValues();

  var payment = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[0] == selectedId) {
      payment = {
        date: Utilities.formatDate(row[1], "GMT-3", "dd/MM/yyyy"),
        amount: Number(row[2]),
        client_last: row[6],
        client_name: row[7]
      }
      break;
    }
  }

  var payment_abbr = " el pago de $ " + payment.amount + " del día " + payment.date + " perteneciente al presupuesto a nombre de " + payment.client_last + ", " + payment.client_name

  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar" + payment_abbr + "?",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {

    var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s");
    var db_payments = db.getSheetByName("DB Pagos");
    var payments = db_payments.getDataRange().getValues();

    Logger.log("selected => " + selectedId);
    for (var i = 1; i < payments.length; i++) {
      Logger.log("i => " + payments[i][0]);
      if (payments[i][0] == selectedId) {
        db_payments.deleteRow(i + 1);
        SpreadsheetApp.getActiveSpreadsheet().toast("Se eliminó" + payment_abbr);
        break;
      }
    }
  }
  uiForm.getRange("J1").clearContent();
}

function onSelectionChange(e) {
  if (e.range.getSheet().getName() == "UI") {
    var row = e.range.getRow();
    var column = e.range.getColumn();
    if (row > 1 && column > 0 && column < 10) {
      var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
      var selected = uiForm.getRange(row, 1, 1, 1).getValue();
      uiForm.getRange("J1").setValue(selected);
    }
  }
}

function installOnSelected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onSelectionChange')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// function deleteTriggers() {
//   // Deletes all triggers in the current project.
//   var triggers = ScriptApp.getProjectTriggers();
//   for (var i = 0; i < triggers.length; i++) {
//     ScriptApp.deleteTrigger(triggers[i]);
//   }
// }