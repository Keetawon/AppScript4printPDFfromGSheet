function generatePDFsFromSheet(sheetName = "Data") {
  // Spreadsheet and sheet setup
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  // Google Doc template ID
  const docTemplateId = "1WTUI8wqhJ9g0t59sYtJV98yHY-ctHda5LxrdV2x17wk";

  // Create or find the base folder named "Generate ใบรับสินค้า"
  //const rootFolderName = "Generate ใบรับสินค้า";
  //const rootFolder = DriveApp.getFoldersByName(rootFolderName);
  const rootFolder = DriveApp.getFolderById("1HT3VaVUeKI-SfzIDSo4DTx71wPfYgA3I");
  
  // Create or find the main folder "ใบรับสินค้าที่สร้าง_pdf" inside "Generate ใบรับสินค้า"
  const mainFolderName = "ใบรับสินค้าที่สร้าง_pdf";
  const mainFolder = getOrCreateFolder(rootFolder, mainFolderName);

  // Initialize a map to group items by Unit Number
  const groupMap = new Map();

  // Start at i = 1 to skip the header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Data from the current row
    const customerName = row[13];   // Column N (13) - Name
    const telNo = row[14];
    const so_No = row[0];           // Column A (0) - SO_No
    const po_No = row[36];
    const project = row[8];
    const floor = row[56];
    const unitType = row[18];
    const unitNo = row[9];
    const projectnunitNo = row[7];
    const building = row[55];
    const houseNo = row[10];
    const dateObj = new Date(row[41]);
    
    // Format the date to dd/mm/yyyy
    const deliveryDate = formatDateToDDMMYYYY(dateObj);
    const deliveryTimeRange = row[42];
    const installedProduct = row[23];
    const category = row[22];
    const brand = row[25];
    const productHi1 = row[26];
    const productColor = row[24];
    const installedRoom = row[27];
    const installedPoint = row[28];
    const quantity = row[34];
    const note = row[45];

    // Create a unique key based on unit number details
    const unitKey = `${so_No}-${customerName}`

    // If the unitKey doesn't exist in the map, initialize an object
    if (!groupMap.has(unitKey)) {
      groupMap.set(unitKey, {
        customerName: customerName,
        telNo: telNo,
        so_No: so_No,
        po_No: po_No,
        project: project,
        floor: floor,
        unitType: unitType,
        unitNo: unitNo,
        projectnunitNo: projectnunitNo,
        building: building,
        houseNo: houseNo,
        items: []
      });
    }

    // Add the item to the unit's list of items
    groupMap.get(unitKey).items.push({
      installedProduct: installedProduct,
      category: category,
      brand: brand,
      productHi1: productHi1,
      productColor: productColor,
      installedRoom: installedRoom,
      installedPoint: installedPoint,
      quantity: quantity,
      deliveryDate: deliveryDate,
      deliveryTimeRange: deliveryTimeRange,
      note: note
    });
  }

  // Iterate through the map and generate PDFs for each unit number
  groupMap.forEach((unitData, unitKey) => {
    // Create a Project Folder if it doesn't exist
    const projectFolder = getOrCreateFolder(mainFolder, unitData.project);
    const projectnUnitNoFolder = getOrCreateFolder(projectFolder, unitData.projectnunitNo);

    // Create a copy of the template
    const newDocId = DriveApp.getFileById(docTemplateId).makeCopy().getId();
    const doc = DocumentApp.openById(newDocId);
    const body = doc.getBody();

    // Replace placeholders with Unit details
    body.replaceText('{{CustomerName}}', unitData.customerName);
    body.replaceText('{{telNo}}', unitData.telNo);
    body.replaceText('{{SO_No}}', unitData.so_No);
    body.replaceText('{{PO_No}}', unitData.po_No);
    body.replaceText('{{Project}}', unitData.project);
    body.replaceText('{{floor}}', unitData.floor);   // Add the item name to the document
    body.replaceText('{{UnitType}}', unitData.unitType); // Add the index (1, 2, etc.) to represent the PDF count
    body.replaceText('{{UnitNo}}', unitData.unitNo);
    body.replaceText('{{Building}}', unitData.building);
    body.replaceText('{{UnitType}}', unitData.unitType);
    body.replaceText('{{HouseNo}}', unitData.houseNo);

    // Populate the placeholder table with items
    populateItemTable(unitData.items, doc);

    // Save and close the document
    doc.saveAndClose();

    // Create a PDF file name based on unit no
    const pdfFileName = `ใบรับสินค้า_${unitData.so_No}.pdf`;

    // Check if a file with the same name already exists and delete it if found
    const existingFiles = projectnUnitNoFolder.getFilesByName(pdfFileName);
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setTrashed(true);
    }

    // Convert the document to PDF
    const pdfBlob = convertDocToPDF(newDocId, pdfFileName);

    // Save the PDF to the designated folder
    projectnUnitNoFolder.createFile(pdfBlob);

    // Clean up by deleting the temporary Google Doc
    DriveApp.getFileById(newDocId).setTrashed(true);
  });
}

// Function to find and populate the item table in the document
function populateItemTable(items, doc) {
  const tables = doc.getBody().getTables();

  // Ensure there are at least two tables (one for unit details and one for items)
  if (tables.length > 1) {
    const table = tables[1]; // The second table is for items

    // Remove all rows except the header row (assuming the first row is the header)
    while (table.getNumRows() > 1) {
      table.removeRow(1);
    }

  // Loop through the items data and add each as a row
  items.forEach(item => {
    // Ensure values are valid strings or use a default placeholder if empty
    const installedProduct = item.installedProduct || "-";
    const category = item.category || "-";
    const brand = item.brand || "-";
    const productHi1 = item.productHi1 || "-";
    const productColor = item.productColor || "-";
    const installedRoom = item.installedRoom || "-";
    const installedPoint = item.installedPoint || "-";
    const note = item.note || "-";
    const deliveryDate = item.deliveryDate || "-";
    const deliveryTimeRange = item.deliveryTimeRange || "-";
    const quantity = item.quantity.toString() || "-";

    const row = table.appendTableRow();
    row.appendTableCell(installedProduct);
    row.appendTableCell(category);
    row.appendTableCell(brand);
    row.appendTableCell(productHi1);
    row.appendTableCell(productColor);
    row.appendTableCell(installedRoom);
    row.appendTableCell(installedPoint);
    row.appendTableCell(note);
    row.appendTableCell(deliveryDate);
    row.appendTableCell(deliveryTimeRange);
    row.appendTableCell(quantity);
    row.appendTableCell(" ");
    row.appendTableCell(" ");
    });
  } else {
    Logger.log("The document doesn't contain the required second table for items.");
  }
}

// Funcion to convert Google Doc to PDF
function convertDocToPDF(docId, pdfFileName) {
  const docFile = DriveApp.getFileById(docId);
  const pdfBlob = docFile.getAs(MimeType.PDF);
  pdfBlob.setName(pdfFileName);
  return pdfBlob;
}

function formatDateToDDMMYYYY(dateObj) {
  const day = ("0" + dateObj.getDate()).slice(-2); // Extract day and pad with 0 if needed
  const month = ("0" + (dateObj.getMonth() + 1)).slice(-2); // Extract month and pad with 0 if needed
  const year = dateObj.getFullYear(); // Extract year
  return `${day}/${month}/${year}`;
}

// Function to get or create a folder by name
function getOrCreateFolder(parentFolder, folderName) {
  if (!folderName) {
    throw new Error("Folder name cannot be null or empty.");
  }

  // Get folders by name within the parent folder
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }

  // If not found, create a new folder with the specified name
  return parentFolder.createFolder(folderName);
}
