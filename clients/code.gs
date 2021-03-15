function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Más')
    .addItem('Editar', 'editClient')
    .addItem('Eliminar', 'deleteClient')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Configurar").addItem("Configurar ahora", "installOnSelected"))
    .addToUi();
}

function editClient() {
  var html = HtmlService.createTemplateFromFile("page_edit_client");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1200).setHeight(1000), "Editar cliente");
}

function getClientData() {
  var id = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange('J1').getValue();

  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Clientes");
  var values = db.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[0] == id) {
      return {
        id: row[0],
        last: row[1],
        first: row[2],
        adress: row[3],
        zip: row[4],
        city: row[5],
        state: row[6],
        email: row[7],
        phone: row[8]
      };
    }
  }
  return null;
}

function save(data) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Clientes");
  var db_values = db.getDataRange().getValues();
  for (var i = 0; i < db_values.length; i++) {
    var row = db_values[i];
    if (parseInt(row[0]) == parseInt(data[0])) {
      db.getRange(i+1, 1, 1, 9).setValues([data]);
      break;
    }
  }
}

function deleteClient() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiForm = ss.getSheetByName("UI");
  var ui = SpreadsheetApp.getUi();
  var id = uiForm.getRange("J1").getValue();

  var last;
  var name;

  var ui_values = uiForm.getDataRange().getValues();

  for (var i = 0; i < ui_values.length; i++) {
    var row = ui_values[i];
    if (row[0] == id) {
      last = row[1];
      name = row[2];
      break;
    }
  }

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar a \"" + last + ", " + name + " (" + id + ")\"?",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {

    var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Clientes");
    var db_values = db.getDataRange().getValues();

    for (var i = 1; i < db_values.length; i++) {
      if (db_values[i][0] == id) {
        db.deleteRow(i + 1);
        SpreadsheetApp.getActiveSpreadsheet().toast("Se eliminó a \"" + last + ", " + name + " (" + id + ")\"");
        break;
      }
    }
  }
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