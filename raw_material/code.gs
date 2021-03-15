function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Más')
    .addItem('Agregar', 'addMaterial')
    .addItem('Editar', 'editMaterial')
    .addSeparator()
    .addItem('Eliminar', 'deleteMaterial')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Configurar").addItem("Configurar ahora", "installOnSelected"))
    .addToUi();
}

function addMaterial() {
  var html = HtmlService.createTemplateFromFile("add_material");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(400), "Agregar item");
}

function getNames() {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima").getDataRange().getValues();
  var names = [];
  for (var i = 0; i < db.length; i++) {
    var row = db[i];
    names.push(row[0]);
  }
  return names;
}

function add(data) {
  SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima").appendRow(data);
}

function editMaterial() {
  var selected_id = SpreadsheetApp.getActiveSpreadsheet().getRange("C1").getValue(); 

  if (selected_id == "") {
    SpreadsheetApp.getUi().alert("No se seleccionó ningún material.");
    return;
  }

  var html = HtmlService.createTemplateFromFile("edit_material");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(400), "Editar item");
}

function getSelectedItem() {
  var selectedItemName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange("C1").getValue();
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima");
  var db_values = db.getDataRange().getValues();
  for (var i = 0; i < db_values.length; i++) {
    var row = db_values[i];
    if (row[0] == selectedItemName) {
      Logger.log({
        name: row[0],
        amount: row[1]
      });
      return {
        name: row[0].toString(),
        amount: Number(row[1])
      }
    }
  }
  return null;
}

function save(data) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima");
  var db_values = db.getDataRange().getValues();
  for (var i = 0; i < db_values.length; i++) {
    var row = db_values[i];
    if (row[0] == data[0]) {
      db.getRange(i + 1, 1, 1, 2).setValues([data]);
      break;
    }
  }
}

function deleteMaterial() {
  var selectedItemName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange("C1").getValue();
  if (selectedItemName == "") {
    SpreadsheetApp.getUi().alert("No se seleccionó ningún material.");
    return;
  }
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar a \"" + selectedItemName + "\"?",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {

    var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima");
    var db_values = db.getDataRange().getValues();

    for (var i = 1; i < db_values.length; i++) {
      if (db_values[i][0] == selectedItemName) {
        db.deleteRow(i + 1);
        SpreadsheetApp.getActiveSpreadsheet().toast("Se eliminó \"" + selectedItemName + "\"");
        break;
      }
    }
  }
}

function onSelectionChange(e) {
  if (e.range.getSheet().getName() == "UI") {
    var row = e.range.getRow();
    var column = e.range.getColumn();
    if (row > 1 && column > 0 && column < 3) {
      var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
      var selected = uiForm.getRange(row, 1, 1, 1).getValue();
      uiForm.getRange("C1").setValue(selected);
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