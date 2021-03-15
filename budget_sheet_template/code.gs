function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Más")
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Items")
        .addSubMenu(
          SpreadsheetApp.getUi()
            .createMenu("Agregar")
            .addItem("Producto", "addNewItem")
            .addItem("Tela", "addNewFabric")
        )
        .addItem("Eliminar", "deleteItem")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Enviar")
        .addItem("Doc", "exportAsDoc")
    )
    .addToUi();
}

function addNewItem() {
  var html = HtmlService.createTemplateFromFile("add_product");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1200).setHeight(1000), "Agregar ítem");
}

function addNewFabric() {
  var html = HtmlService.createTemplateFromFile("add_fabric");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(500), "Agregar tela");
}

function getData() {
  return {
    materials: getMaterials(),
    templates: getTemplates(),
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
  var templates = {};

  for (var i = 1; i < db_values.length; i++) {
    var row = db_values[i];
    templates[row[0]] = {
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

function saveAsTemplate2(template) {
  var db_values = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados");
  db_values.appendRow(template);
}

function addProduct(product) {
  var db_values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("selected_products");
  db_values.appendRow(product);
}

function deleteItem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiForm = ss.getSheetByName("UI");
  var selectedForm = ss.getSheetByName("selected_products");
  var ui = SpreadsheetApp.getUi();

  var selectedName = uiForm.getRange("A1").getValue();

  if (selectedName == '') {
    ui.alert("Ningún producto seleccionado.")
    return;
  }
  
  var lastRow = selectedForm.getLastRow();
  var selected_values = selectedForm.getDataRange().getValues();

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar todos los productos con el nombre \"" + selectedName + "\"? Esta operación pueda tardar un poco.",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    for (var i = lastRow - 1; i > 0; i--) {
      var row = selected_values[i];
      if (row[0] == selectedName) {
        selectedForm.deleteRow(i + 1);
      }
    }
  }
  uiForm.getRange("J21").clearContent();
}

function exportAsDoc() {
  var budgetDocTemplate = DriveApp.getFileById("1HQxEEMbMPhFPrSIlZ9PKgWz5x3guwDKKr2f27HaNxw0");
  var budgetsDir = DriveApp.getFolderById("1MIsw1-hkMmgGA3ireMg0I7qJOht3VjsI");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiForm = ss.getSheetByName("UI");

  var number = ss.getSheetByName("Presupuesto").getRange("C2").getValue();
  const today = new Date();
  var creationDate = getNewCreationDate(today);
  var dateForName = getDateForName(today);
  var expirationDate = getExpirationDate(today);
  var client = getClient();
  var subtotal = "$ " + Number(uiForm.getRange('B8').getValue()).toLocaleString('es-ES', { minimumFractionDigits: 2 });
  var discount = Number(uiForm.getRange('E8').getValue() * 100).toLocaleString('es-ES', { minimumFractionDigits: 2 }) + "%";
  var total = "$ " + Number(uiForm.getRange('I8').getValue()).toLocaleString('es-ES', { minimumFractionDigits: 2 });
  var products = getProducts();

  var budgetDocName = number + "_" + dateForName + "_" + client.last_name + "_" + client.name;

  const copy = budgetDocTemplate.makeCopy(budgetDocName, budgetsDir);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const tables = body.getTables();

  var client_info = client.id + "\n" + client.address + "\n" + client.zip_code + " " + client.city + "\n" + client.state + "\n" + client.phone + "\n" + client.email;

  body.replaceText("{{fecha_creacion}}", creationDate);
  body.replaceText("{{nro}}", number);
  body.replaceText("{{nombre}}", client.full_name);
  body.replaceText("{{client_info}}", client_info);
  body.replaceText("{{subtotal}}", subtotal);
  body.replaceText("{{fecha_vencimiento}}", expirationDate);

  if (discount === '' || discount === 0 || discount === "0" || discount === "0.0" || discount === 0.0) {
    tables[2].removeRow(1);
  } else {
    body.replaceText("{{descuento}}", discount);
  }

  body.replaceText("{{total}}", total);

  var products_table = tables[1];
  var template_row = products_table.getRow(1);

  for (var i = 0; i < products.length; i++) {
    var ROW_INDEX = i + 2;
    products_table.appendTableRow(template_row.copy());
    var current_row = products_table.getRow(ROW_INDEX);
    for (var j = 0; j < 4; j++) {
      current_row.getCell(j).setText(products[i][j]);
    }
  }

  products_table.removeRow(1);

  var budget_sheet = ss.getSheetByName("Presupuesto");
  budget_sheet.getRange('C4').setValue(doc.getUrl());
  budget_sheet.getRange('C6').setValue(expirationDate);

  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Pedidos");
  var db_values = db.getDataRange().getValues();

  for (var i = 0; i < db_values.length; i++) {
    var row = db_values[i];
    if (row[0] == number) {
      db.getRange(i+1,9,1,1).setValue(doc.getUrl());
      db.getRange(i+1,10,1,1).setValue(expirationDate);
      db.getRange(i+1,11,1,1).setValue(total);
      break;
    }
  }

  doc.saveAndClose();
}

function getNewCreationDate(today) {
  return Utilities.formatDate(today, "GMT-3", "dd-MM-yyyy").toString();
}

function getExpirationDate(date) {
  var d = date;
  d.setDate(d.getDate() + 15);
  return Utilities.formatDate(d, "GMT-3", "dd-MM-yyyy").toString();
}

function getDateForName(date) {
  return Utilities.formatDate(date, "GMT-3", "yyyyMMdd").toString();
}

function getClient() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clientForm = ss.getSheetByName("Datos cliente");
  var id = clientForm.getRange('B2').getValue();
  var last_name = clientForm.getRange('B3').getValue();
  var name = clientForm.getRange('B4').getValue();
  var full_name = last_name + ", " + name;
  var adress = clientForm.getRange('B5').getValue();
  var zip_code = clientForm.getRange('B6').getValue();
  var city = clientForm.getRange('B7').getValue();
  var state = clientForm.getRange('B8').getValue();
  var email = clientForm.getRange('B9').getValue();
  var phone = clientForm.getRange('B10').getValue();
  var data = {
    id: id,
    name: name,
    last_name: last_name,
    full_name: full_name,
    address: adress,
    zip_code: zip_code,
    city: city,
    state: state,
    email: email,
    phone: phone,
  }
  return data;
}

function getProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var productsForm = ss.getSheetByName("selected_products");
  var products_values = productsForm.getDataRange().getValues();

  var products = [];

  for (var i = 1; i < products_values.length; i++) {
    var row = products_values[i]
    var name = row[0];
    var value = "$ " + row[1].toLocaleString('es-ES', { minimumFractionDigits: 2 });
    var quantity = row[2].toLocaleString('es-ES', { minimumFractionDigits: 2 });
    var amount = "$ " + row[3].toLocaleString('es-ES', { minimumFractionDigits: 2 });
    var product = [name, quantity, value, amount];
    products.push(product);
  }

  return products;
}

function onEdit(e) {
  var range = e.range.getA1Notation();
  if (e.range.getSheet().getName() == "UI") {
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI").getRange(range).getValue() == "") {
      if (range == "I3") clearVariables("L3", "N3", "P3");
      if (range == "I6") clearVariables("L6", "N6", "P6");
      if (range == "I9") clearVariables("L9", "N9", "P9");
    }
    if (range == "I12") updateManufactureVariables();
  }
}

function dialogSearchTemplateProduct() {

  const rows = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s/edit#gid=0").getSheetByName("DB Predeterminados").getDataRange().getValues();

  let data = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const name = row[0];
    if (name != "") {
      data[name] = null;
    } else break;
  }

  const dialog = HtmlService.createTemplateFromFile("dialog_template_product_search");
  dialog.templates = data;
  SpreadsheetApp.getUi().showSidebar(dialog.evaluate().setTitle("Seleccionar plantilla"));
}

function updateManufactureVariables() {
  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  var currentName = uiForm.getRange("I12").getValue();

  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB Costos confección").getDataRange().getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var name = row[0];
    if (name == currentName) {
      uiForm.getRange("L2").setValue(row[2]);
      uiForm.getRange("N2").setValue(row[3]);
      uiForm.getRange("P2").setValue(row[4]);
      break;
    }
  }
}

function clearVariables(mult, factor, tolerance) {
  var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
  uiForm.getRange(mult).setValue("nada");
  uiForm.getRange(factor).setValue(1);
  uiForm.getRange(tolerance).setValue(0);
}

function setTemplateProduct(templateName) {
  const ext_ssheet = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s");
  const cur_ssheet = SpreadsheetApp.getActiveSpreadsheet();
  const templates = ext_ssheet.getSheetByName("DB Predeterminados");
  const completable = cur_ssheet.getSheetByName("UI");

  const values = templates.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    let row = values[i];
    if (row[0] == templateName) {
      completable.getRange('C3').setValue(row[0]); // name
      completable.getRange('I3').setValue(row[1]); // comp1
      completable.getRange("L3").setValue(row[2]); // mult1
      completable.getRange("N3").setValue(row[3]); // factor1
      completable.getRange("P3").setValue(row[4]); // tolerance1
      completable.getRange("I6").setValue(row[5]); // comp2
      completable.getRange("L6").setValue(row[6]); // mult2
      completable.getRange("N6").setValue(row[7]); // factor2
      completable.getRange("P6").setValue(row[8]); // tolerance2
      completable.getRange("I9").setValue(row[9]); // comp3
      completable.getRange("L9").setValue(row[10]); // mult3
      completable.getRange("N9").setValue(row[11]); // factor3
      completable.getRange("P9").setValue(row[12]); // tolerance3
      completable.getRange("I12").setValue(row[13]); // manuf
      completable.getRange("L12").setValue(row[14]); // manufMult
      completable.getRange("N12").setValue(row[15]); // manufFactor
      completable.getRange("P12").setValue(row[16]); // manufTolerance
      break;
    }
  }
}

function deleteSelectedItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiForm = ss.getSheetByName("UI");
  var selectedForm = ss.getSheetByName("selected_products");
  var ui = SpreadsheetApp.getUi();

  if (uiForm.getRange("I21").getValue() == '') {
    ui.alert("Ningún producto seleccionado.")
    return;
  }

  var selectedName = uiForm.getRange("I21").getValue();
  var lastRow = selectedForm.getLastRow();
  var selected_values = selectedForm.getDataRange().getValues();

  var result = ui.alert(
    'Confirmar',
    "¿Querés eliminar todos los productos con el nombre \"" + selectedName + "\"? Esta operación pueda tardar un poco.",
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    for (var i = lastRow - 1; i > 0; i--) {
      var row = selected_values[i];
      if (row[0] == selectedName) {
        selectedForm.deleteRow(i + 1);
      }
    }
  }
  uiForm.getRange("J21").clearContent();
}

function addItem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = ss.getSheetByName("UI");
  const product = ui.getRange('C3').getValue();
  const height = ui.getRange('C9').getValue();
  const width = ui.getRange('E9').getValue();
  const name = product.concat(" (", height, "x", width, " m)");
  const amount = ui.getRange('E16').getValue();
  const units = ui.getRange('I16').getValue();
  const subtotal = ui.getRange('N16').getValue();
  const selected_sheet = ss.getSheetByName("selected_products");
  selected_sheet.getRange(selected_sheet.getLastRow() + 1, 1, 1, 4).setValues([[name, amount, units, subtotal]]);
}

function saveAsTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uiForm = ss.getSheetByName("UI");
  var name = uiForm.getRange('C3').getValue();

  var db = SpreadsheetApp.openById("1LAz6ueun9TKEPNSOFzs1He_I4sIhsqBOkZydhImQi9s").getSheetByName("DB Predeterminados");
  var db_values = db.getDataRange().getValues();
  for (var i = 0; i < db_values.length; i++) {
    if (db_values[i][0] == name) {
      SpreadsheetApp.getUi().alert("Ya existe un producto con ese nombre.")
      return;
    }
  }

  // comp1
  var comp1 = uiForm.getRange('I3').getValue();
  var mult1 = uiForm.getRange('L3').getValue();
  var factor1 = uiForm.getRange('N3').getValue();
  var tolerance1 = uiForm.getRange('P3').getValue();

  // comp2
  var comp2 = uiForm.getRange('I6').getValue();
  var mult2 = uiForm.getRange('L6').getValue();
  var factor2 = uiForm.getRange('N6').getValue();
  var tolerance2 = uiForm.getRange('P6').getValue();

  // comp3
  var comp3 = uiForm.getRange('I9').getValue();
  var mult3 = uiForm.getRange('L9').getValue();
  var factor3 = uiForm.getRange('N9').getValue();
  var tolerance3 = uiForm.getRange('P9').getValue();

  // manuf
  var manuf = uiForm.getRange('I12').getValue();
  var multManuf = uiForm.getRange('L12').getValue();
  var factorManuf = uiForm.getRange('N12').getValue();
  var toleranceManuf = uiForm.getRange('P12').getValue();

  var template = [name, comp1, mult1, factor1, tolerance1, comp2, mult2, factor2, tolerance2, comp3, mult3, factor3, tolerance3, manuf, multManuf, factorManuf, toleranceManuf];

  db.appendRow(template);
}

function onSelectionChange(e) {
  if (e.range.getSheet().getName() == "UI") {
    var row = e.range.getRow();
    var column = e.range.getColumn();
    if (row > 10 && column > 1 && column < 11) {
      var uiForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UI");
      var selected = uiForm.getRange(row, 2, 1, 1).getValue();
      uiForm.getRange("A1").setValue(selected);
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