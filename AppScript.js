function generatePDFsFromSheet(sheetName = "Data") {
    // Spreadsheet and sheet setup
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
  
    // Create or find the base folder named "Generate ใบรับสินค้า"
    const rootFolderName = "Generate ใบรับสินค้า";
    const rootFolder = createFolderInDrive(rootFolderName);
    
    // Create or find the main folder "ใบรับสินค้าที่สร้าง_pdf" inside "Generate ใบรับสินค้า"
    const mainFolderName = "ใบรับสินค้าที่สร้าง_pdf";
    const mainFolder = createFolderInDrive(mainFolderName, rootFolder);
  
    // Skipping the first row (header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
  
      // Data from the current row
      const customerName = row[13];   // Column N (13) - Name
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
  
      // Select Google Doc template based on category
      const docTemplateId = getTemplateIdForCategory(category);
  
      // If no matching template found, skip this row
      if (!docTemplateId) {
        Logger.log(`No template found for category: ${category}`);
        continue;
      }
  
      // Create or find the parent folder (named after Name) inside the mainFolder folder
      const parentFolderName = project;           // Parent folder: John_Doe
      const parentFolder = createFolderInDrive(parentFolderName, mainFolder);
  
      // Create or find the subfolder (named after Date) inside the parent folder
      const subFolderName = projectnunitNo;              // Subfolder: 2024-09-17
      const subFolder = createFolderInDrive(subFolderName, parentFolder);
  
      // Generate multiple PDFs based on the quantity
      for (let j = 1; j <= quantity; j++) {
        let pdfFileName;
  
        // Check if item is blank, and set the file name accordingly
        if (!installedRoom || installedRoom.trim() === "") {
          pdfFileName = `${projectnunitNo}_${category}-${installedProduct}_${j}.pdf`;
        } else {
          if (!installedPoint || installedPoint.trim() === "") {
            pdfFileName = `${projectnunitNo}_${category}-${installedProduct} (${installedRoom})_${j}.pdf`
          } else {
            pdfFileName = `${projectnunitNo}_${category}-${installedProduct} (${installedRoom} ${installedPoint})_${j}.pdf`
          }
        }
  
        pdfFileName = pdfFileName.trim()
  
        // Check if a file with the same name already exists in the folder
        const existingFile = findFileInFolder(pdfFileName, subFolder);
        if (existingFile) {
          // If the file exists, delete it
          existingFile.setTrashed(true);
          Logger.log(`Existing file ${pdfFileName} deleted.`)
        }
  
        // Create a copy of the template, fill it with data
        const newDocId = fillTemplateWithData(
          docTemplateId,
          {
            customerName, so_No, po_No,
            project, floor, unitType, unitNo, projectnunitNo, building, houseNo,
            deliveryDate, deliveryTimeRange,
            installedProduct, category, brand, productHi1, productColor, installedRoom, installedPoint,
            quantity, index: j
          }
        );
  
        // Convert the new Google Doc to PDF
        const pdfFile = convertDocToPDF(newDocId, pdfFileName);
  
        // Save the PDF in the subfolder (John_Doe/2024-09-17)
        savePDFToFolder(pdfFile, subFolder);
        
        // Optionally delete the intermediate Google Doc
        DriveApp.getFileById(newDocId).setTrashed(true);
      }
    }
  }
  
  // Map categories to Google Dco template IDs
  function getTemplateIdForCategory(category) {
    const templateMap = {
      'เครื่องทำน้ำอุ่น': '14E9ngQE6GDgOiT49nt_Gw13m66NvHc-6mFSKhUJriLE',
      'เครื่องทำน้ำอุ่นพร้อมติดตั้ง': '1xBZxMIFuMSjNrCi8KUd2aihg94ifASQwc58Uq6m-syE',
      'บริการติดตั้ง : เครื่องทำน้ำอุ่น' : '1DKhbBTynhrt6cjJWSp25VqrEmK0eYH1A-xonypRY0m0',
      'เครื่องปรับอากาศ' : '14E9ngQE6GDgOiT49nt_Gw13m66NvHc-6mFSKhUJriLE',
      'บริการติดตั้ง : เครื่องปรับอากาศ' : '1DKhbBTynhrt6cjJWSp25VqrEmK0eYH1A-xonypRY0m0',
      'โครงหลังคา-หน้าบ้าน' : '1xBZxMIFuMSjNrCi8KUd2aihg94ifASQwc58Uq6m-syE',
      'โครงหลังคา-หลังบ้าน' : '1xBZxMIFuMSjNrCi8KUd2aihg94ifASQwc58Uq6m-syE',
      'กระจกกั้นห้องอาบน้ำพร้อมติดตั้ง' : '1xBZxMIFuMSjNrCi8KUd2aihg94ifASQwc58Uq6m-syE',
      'ม่าน' : '1xBZxMIFuMSjNrCi8KUd2aihg94ifASQwc58Uq6m-syE'
      // Add more categories and template IDs as needed
    };
    return templateMap[category] || null;
  }
  
  // Fill the template with data and return the new document's ID
  function fillTemplateWithData(templateId, data) {
    const template = DriveApp.getFileById(templateId);
    const newDoc = template.makeCopy(`Document for ${data.name}`).getId();
    const doc = DocumentApp.openById(newDoc);
    const body = doc.getBody();
  
    // Replace placeholders with data
    body.replaceText('{{CustomerName}}', data.customerName);
    body.replaceText('{{SO_No}}', data.so_No);
    body.replaceText('{{PO_No}}', data.po_No);
    body.replaceText('{{Project}}', data.project);
    body.replaceText('{{floor}}', data.floor);   // Add the item name to the document
    body.replaceText('{{UnitType}}', data.unitType); // Add the index (1, 2, etc.) to represent the PDF count
    body.replaceText('{{UnitNo}}', data.unitNo);
    body.replaceText('{{Building}}', data.building);
    body.replaceText('{{HouseNo}}', data.unitType);
    body.replaceText('{{DeliveryDate}}', data.deliveryDate);
    body.replaceText('{{DeliveryTimeRange}}', data.deliveryTimeRange);
    body.replaceText('{{InstalledProduct}}', data.installedProduct);
    body.replaceText('{{ProductCate}}', data.category);
    body.replaceText('{{Brand}}', data.brand);
    body.replaceText('{{ProductHi1}}', data.productHi1);
    body.replaceText('{{ProductColor}}', data.productColor);
    body.replaceText('{{InstalledRoom}}', data.installedRoom);
    body.replaceText('{{InstalledPoint}}', data.installedPoint);
  
    doc.saveAndClose();
    return newDoc;
  }
  
  function formatDateToDDMMYYYY(dateObj) {
    const day = ("0" + dateObj.getDate()).slice(-2); // Extract day and pad with 0 if needed
    const month = ("0" + (dateObj.getMonth() + 1)).slice(-2); // Extract month and pad with 0 if needed
    const year = dateObj.getFullYear(); // Extract year
    return `${day}/${month}/${year}`;
  }
  
  // Convert the Google Doc to a PDF and return the PDF file
  function convertDocToPDF(docId, fileName) {
    const doc = DriveApp.getFileById(docId);
    const pdfBlob = doc.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(fileName);
    return pdfFile;
  }
  
  // Helper function to find a file by name in a specific folder
  function findFileInFolder(fileName, folder) {
    const files = folder.getFilesByName(fileName);
    return files.hasNext() ? files.next() : null;
  }
  
  // Create folders in Google Drive (in a parent folder if specified)
  function createFolderInDrive(folderName, parentFolder = null) {
    let folder;
    
    if (parentFolder) {
      // If a parent folder is specified, look inside the parent folder
      const subFolders = parentFolder.getFoldersByName(folderName);
      if (subFolders.hasNext()) {
        folder = subFolders.next();
      } else {
        folder = parentFolder.createFolder(folderName);
      }
    } else {
      // Otherwise, look at the root level
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }
    }
    return folder;
  }
  
  // Save the PDF file to the specific folder
  function savePDFToFolder(pdfFile, folder) {
    folder.addFile(pdfFile);
    Logger.log(`PDF saved to folder: ${folder.getName()} with file name: ${pdfFile.getName()}`);
  }
  