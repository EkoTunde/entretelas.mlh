function onOpen(e) {
  createMenus();
}

function createMenus() {
  SpreadsheetApp.getUi()
    .createMenu("Más")
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Presupuesto")
        .addItem("Nuevo", "addBudget")
        .addItem("Finalizar", "endBudget")
        .addSeparator()
        .addItem("Cancelar", "cancelBudget")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Pagos")
        .addItem("Agregar pago", "addPayment"))
    .addToUi();
}

function addBudget() {
  var html = HtmlService.createTemplateFromFile("add_budget");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(800).setHeight(1000), "Nuevo presupuesto");
}

function getSearchData() {
  var clients_db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Clientes").getDataRange().getValues();

  var clients = {};
  var search_items = {};

  for (var i = 1; i < clients_db.length; i++) {
    var row = clients_db[i];
    var search_str = row[1] + ", " + row[2] + " (" + row[0] + ")";
    clients[search_str] = {
      id: row[0],
      last_name: row[1],
      first_name: row[2],
      adress: row[3],
      zip: row[4],
      city: row[5],
      state: row[6],
      email: row[7],
      phone: row[8],
    };
    search_items[search_str] = null;
  }
  return { clients: clients, search_items: search_items };
}

function getNextBudgetId() {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("Ids");
  var lastBudgetId = db.getRange('B2').getValue();
  return parseInt(lastBudgetId) + 1;
}

function getNextClientId() {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("Ids");
  var lastClientId = db.getRange('B1').getValue();
  return parseInt(lastClientId) + 1;
}

function saveNewOrder(new_order) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s");
  var orders_db = db.getSheetByName("DB Pedidos");
  var clients_db = db.getSheetByName("DB Clientes");
  var ids_db = db.getSheetByName("Ids");

  var client = new_order.client;
  var date = new Date(new_order.timestamp);
  orders_db.appendRow([new_order.budgetId, date, new_order.finalizado, client.id, client.last_name, client.first_name, new_order.order]);
  ids_db.getRange('B2').setValue(new_order.budgetId);

  if (isNewClient(new_order.search_data_clients, client.id)) {
    clients_db.appendRow([client.id, client.last_name, client.first_name, client.adress, client.zip, client.city, client.state, client.email, client.phone]);
    if (client.id == getNextClientId()) {
      ids_db.getRange('B1').setValue(client.id);
    }
  }

  var data = {
    id: new_order.budgetId,
    date: date,
    order: new_order.order,
    client_id: client.id,
    client_full_name: client.last_name.toUpperCase() + "_" + client.first_name.toUpperCase()
  };
  createBudgetSSheetFromTemplate(data);
}

function createBudgetSSheetFromTemplate(data) {
  var budgetSSheetTemplate = DriveApp.getFileById("1xbn3QBiG1EfbYIY3pYBiE2WifXZ_OPN5lIOJFncJrQI");
  var budgetsDir = DriveApp.getFolderById("19kI3ZSihjszkTckEo9U6ZPC5CIhxqv-S");
  var name = data.id + "_" + data.date.getFullYear() + (data.date.getMonth() + 1) + data.date.getDate() + "_" + data.client_full_name;
  var copy = budgetSSheetTemplate.makeCopy(name, budgetsDir);
  var ss = SpreadsheetApp.openById(copy.getId());

  var order_sheet = ss.getSheetByName("Pedido");
  order_sheet.getRange('B3').setValue(data.order);

  var client_sheet = ss.getSheetByName("Datos cliente");
  client_sheet.getRange('B2').setValue(data.client_id);

  var budget_sheet = ss.getSheetByName("Presupuesto");
  budget_sheet.getRange('C2').setValue(data.id);

  var ui_sheet = ss.getSheetByName("UI");
  ui_sheet.getRange("B11").setValue("=IFERROR(QUERY(selected_products!$A$2:$I;\"SELECT A,E,F,G,H,I,B,C,D WHERE A<>'' ORDER BY A\");\"No se agregaron items\")");

  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Pedidos");
  var values = db.getDataRange().getValues();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0] == data.id) {
      db.getRange(i + 1, 8, 1, 1).setValue(copy.getUrl());
      break;
    }
  }
}

function isNewClient(clients, client_id) {
  for (var [key, value] of Object.entries(clients)) {
    if (value.id == client_id) {
      return false;
    }
  }
  return true;
}

function endBudget() {
  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  var budget_db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Pedidos");
  var ui = SpreadsheetApp.getUi();

  if (uiForm.getRange("B1").getValue() == '') {
    ui.alert("Ningún presupuesto seleccionado.")
    return;
  }

  var selectedId = uiForm.getRange("B1").getValue();
  var lastRow = budget_db.getLastRow();
  var selected_values = budget_db.getDataRange().getValues();

  var result = ui.alert(
    'Confirmar',
    "¿Querés finalizar el presupuesto #" + selectedId + "? Esta operación pueda tardar un poco.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result == ui.Button.OK) {
    for (var i = lastRow - 1; i > 0; i--) {
      var row = selected_values[i];
      if (row[0] == selectedId) {
        budget_db.getRange(i + 1, 3, 1, 1).setValue('SI');
        uiForm.getRange("B1").clearContent();
        break;
      }
    }
  }
}

function cancelBudget() {
  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  var budget_db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Pedidos");
  var ui = SpreadsheetApp.getUi();

  if (uiForm.getRange("B1").getValue() == '') {
    ui.alert("Ningún presupuesto seleccionado.")
    return;
  }

  var selectedId = uiForm.getRange("B1").getValue();
  var lastRow = budget_db.getLastRow();
  var selected_values = budget_db.getDataRange().getValues();

  var result = ui.alert(
    'Confirmar',
    "¿Querés cancelar el presupuesto #" + selectedId + "? Esta operación pueda tardar un poco.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result == ui.Button.OK) {
    for (var i = lastRow - 1; i > 0; i--) {
      var row = selected_values[i];
      if (row[0] == selectedId) {
        var url = row[7];
        deleteFileWithUrl(url);
        budget_db.deleteRow(i + 1);
        uiForm.getRange("B1").clearContent();
        break;
      }
    }
  }
}

function deleteFileWithUrl(url) {
  var id = SpreadsheetApp.openByUrl(url).getId();
  DriveApp.getFileById(id).setTrashed(true);
}

function addPayment() {
  var html = HtmlService.createTemplateFromFile("add_payment");

  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  if (uiForm.getRange("B1").getValue() == '') {
    ui.alert("Ningún presupuesto seleccionado.")
    return null;
  }
  var selectedId = uiForm.getRange("B1").getValue();
  var values = uiForm.getDataRange().getValues();
  for (var i = 2; i < values.length; i++) {
    var row = values[i];
    if (row[1] == selectedId) {
      if (row[10] != "" && Number(row[10]) > 0) {
        html.id = row[1];
        html.client = row[5] + " " + row[4] + " (" + row[3] + ")";
        html.date = Utilities.formatDate(uiForm.getRange(i + 1, 3, 1, 1).getValue(), "GMT-3", "dd/MM/yyyy HH:mm");
        html.amount_left = "$ " + row[10].toLocaleString('es-ES', { minimumFractionDigits: 2 });
        break;
      } else {
        SpreadsheetApp.getUi().alert("El presupuesto aun no se confirmó imprimiéndolo/enviándolo.");
        return;
      }
    }
  }
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(500), "Nuevo pago");
}

function getBudgetIdForPayment() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange("B1").getValue();
}

function savePayment(budget_id, amount, date) {
  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s");
  var new_id = Number(db.getSheetByName("Ids").getRange("B3").getValue()) + 1;
  var payments_db = db.getSheetByName("DB Pagos");
  payments_db.appendRow([new_id, budget_id, "$ " + Number(amount).toLocaleString('es-AR', { minimumFractionDigits: 2 }), date]);
  db.getSheetByName("Ids").getRange("B3").setValue(new_id);
}

function installOnSelected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onSelectionChange')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
}

function onSelectionChange(e) {
  if (e.range.getSheet().getName() == "UI") {
    var row = e.range.getRow();
    //var column = e.range.getColumn();
    if (row > 1) { //&& column > 0 && column < 11) {
      var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
      var selected = uiForm.getRange(row, 2, 1, 1).getValue();
      uiForm.getRange("B1").setValue(selected);
    }
  }
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