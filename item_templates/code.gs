function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Más")
    .addItem("Agregar", "addItem")
    .addItem("Editar", "editItem")
    .addSeparator()
    .addItem("Eliminar", "deleteSelectedItems")
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Configuración").addItem("Configurar", "installOnSelected"))
    .addToUi();
}

function addItem() {
  var html = HtmlService.createTemplateFromFile("page_add_item");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1200).setHeight(1000), "Agregar plantilla");
}

function editItem() {
  var name = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange('F1').getValue();
  if (name == "") {
    SpreadsheetApp.getUi().alert("No se seleccionó una plantilla");
    return;
  }
  var html = HtmlService.createTemplateFromFile("page_edit_item");
  html.name = name;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1200).setHeight(1000), "Editar plantilla");
}

function getSelectedItemName() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange("F1").getValue();
}

function getTemplateData(name) {
  var db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("db").getDataRange().getValues();
  for (var i = 0; i < db.length; i++) {
    var row = db[i];
    if (row[0] == name) {
      return {
        name: row[0],
        item_1: row[1],
        mult_1: row[2],
        factor_1: row[3],
        tolerance_1: row[4],
        item_2: row[5],
        mult_2: row[6],
        factor_2: row[7],
        tolerance_2: row[8],
        item_3: row[9],
        mult_3: row[10],
        factor_3: row[11],
        tolerance_3: row[12],
        item_manuf: row[13],
        mult_manuf: row[14],
        factor_manuf: row[15],
        tolerance_manuf: row[16],
      };
    }
  }
}

function getMaterials() {
  var db_values = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Materia prima").getDataRange().getValues();
  var materials = {};
  for (var i = 1; i < db_values.length; i++) {
    var row = db_values[i];
    materials[row[0]] = row[1];
  }
  //Logger.log(Object.keys(materials));
  return materials;
}

function getTemplates() {
  var db_values = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados").getDataRange().getValues();
  var templates = [];
  for (var i = 1; i < db_values.length; i++) {
    templates.push(db_values[i][0]);
  }
  return templates;
}

function getManufactureCosts() {
  var db_values = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Costos confección").getDataRange().getValues();
  var costs = {};
  for (var i = 1; i < db_values.length; i++) {
    var row = db_values[i];
    costs[row[0]] = row[1];
  }
  Logger.log(JSON.stringify(costs));
  return costs;
}

function prueba() {
  var template = [];
  add(template);
}

function add(template) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados");
  db.appendRow(template);
}

function save(template) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados");
  var db_values = db.getDataRange().getValues();
  for (var i = 0; i < db_values.length; i++) {
    var row = db_values[i];
    if (row[0] == template[0]) {
      db.getRange(i+1, 1, 1, 17).setValues([template]);
      break;
    }
  }
}

function deleteSelectedItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiForm = ss.getSheetByName("UI");
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados");
  var ui = SpreadsheetApp.getUi();

  if (uiForm.getRange("F1").getValue() == '') {
    ui.alert("Ningún producto seleccionado.")
    return;
  }

  var selectedName = uiForm.getRange("F1").getValue();
  var lastRow = db.getLastRow();
  var db_values = db.getDataRange().getValues();

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar todos los productos con el nombre \"" + selectedName + "\"? Esta operación pueda tardar un poco.",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    for (var i = lastRow - 1; i > 0; i--) {
      var row = db_values[i];
      if (row[0] == selectedName) {
        db.deleteRow(i + 1);
      }
    }
  }
  uiForm.getRange("F1").clearContent();
}

function onSelectionChange(e) {
  if (e.range.getSheet().getName() == "UI") {
    var row = e.range.getRow();
    var column = e.range.getColumn();
    if (row > 1 && column > 0 && column < 6) {
      var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
      var selected = uiForm.getRange(row, 1, 1, 1).getValue();
      uiForm.getRange("F1").setValue(selected);
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