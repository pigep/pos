/**
* @OnlyCurrentDoc
*/

// --- KONFIGURÁCIA ---
const CONFIG = {
  FOLDER_ID: "1K3pRi0pCysPWgxbD2iIByYYw5LaC-smm", // ID priečinka na Google Drive pre nahrávané súbory
  HEADER_ROW: 16, // Číslo riadku, kde začínajú hlavičky tabuľky
  LOGIN_SHEET: 'Set up login' // Názov hárku s prihlasovacími údajmi
};

/**
* Hlavná funkcia, ktorá sa spustí pri načítaní webovej aplikácie.
* @param {object} e - Objekt udalosti.
* @returns {HtmlOutput} - HTML výstup pre zobrazenie v prehliadači.
*/
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setTitle("Servis Web App")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
* Vloží obsah iného súboru (napr. CSS alebo JavaScript) do HTML šablóny.
* @param {string} filename - Názov súboru na vloženie.
* @returns {string} - Obsah súboru.
*/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
* Získa zoznam relevantných hárkov (ktorých názov začína na "data_").
* @returns {Array<object>} Pole objektov s pôvodným a čistým názvom hárku.
*/
function getSheetNames() {
  try {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    return sheets
      .map(sheet => sheet.getName())
      .filter(name => name.toLowerCase().startsWith('data_'))
      .map(name => ({
        originalName: name,
        cleanName: name.substring(5) // Odstráni "data_" prefix
      }));
  } catch (e) {
    Logger.log(`Chyba pri získavaní názvov hárkov: ${e.message}`);
    return [];
  }
}

/**
* Získa dáta a metainformácie z konkrétneho hárku.
* @param {Sheet} sheet - Objekt hárku Apps Script.
* @returns {object|null} Objekt s hlavičkami, dátami, zoznamom ID a možnosťami pre dropdown menu, alebo null pri chybe.
*/
function getSheetData(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < CONFIG.HEADER_ROW) {
      return { headers: [], data: [], idList: [], dropdownOptions: {} };
    }

    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getValues()[0];
    let data = [];
    let fullDataRange;

    if (lastRow > CONFIG.HEADER_ROW) {
      fullDataRange = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastColumn);
      data = fullDataRange.getDisplayValues();
    }

    const idRange = (lastRow > CONFIG.HEADER_ROW) ?
      sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, 1).getValues() : [];

    const dropdownOptions = {};
    headers.forEach((header, index) => {
      if (header.trim().endsWith('-')) {
        if (fullDataRange) {
          const columnValues = fullDataRange.offset(0, index, fullDataRange.getNumRows(), 1).getValues();
          const uniqueValues = [...new Set(columnValues.flat())].filter(String).sort();
          dropdownOptions[index] = uniqueValues;
        } else {
          dropdownOptions[index] = [];
        }
      }
    });

    return {
      spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      headers: headers,
      data: data,
      idList: idRange.map(row => String(row[0])),
      dropdownOptions: dropdownOptions,
    };
  } catch (e) {
    Logger.log(`Chyba v getSheetData pre hárok '${sheet.getName()}': ${e.message}`);
    return null;
  }
}

/**
* Načíta dáta zo VŠETKÝCH relevantných hárkov a filtruje ich podľa používateľa.
* @param {string} userKey - Kľúč prihláseného používateľa.
* @returns {object} Objekt obsahujúci dáta zo všetkých hárkov.
*/
function getAllSheetsData(userKey) {
  const allSheetsData = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().toLowerCase().startsWith('data_'));
  const loginCredentials = getLoginData();
  const adminUser = loginCredentials.length > 0 ? loginCredentials[0][0] : null;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const sheetInfo = getSheetData(sheet);

    if (sheetInfo) {
      let filteredData = sheetInfo.data;

      if (userKey !== adminUser) {
        const userColumnIndex = sheetInfo.headers.indexOf('Užívateľ');
        if (userColumnIndex !== -1) {
          filteredData = sheetInfo.data.filter(row => row[userColumnIndex] === userKey);
        }
      }

      allSheetsData[sheetName] = {
        headers: sheetInfo.headers,
        tableData: filteredData,
        dropdownOptions: sheetInfo.dropdownOptions
      };
    }
  });

  return allSheetsData;
}

/**
* Overí prihlasovacie údaje a vráti počiatočné dáta pre aplikáciu.
* @param {object} objInput - Objekt s prihlasovacími údajmi (input1, input2).
* @returns {object|null} Objekt s dátami alebo null pri neúspešnom prihlásení.
*/
function getInitialData(objInput) {
  try {
    const loginData = getLoginData();
    const userIndex = loginData.findIndex(cred =>
      String(cred[0]).trim() === String(objInput.input1).trim() &&
      String(cred[1]).trim() === String(objInput.input2).trim()
    );

    if (userIndex === -1) {
      return null;
    }

    const userKey = loginData[userIndex][0];
    const allSheetsData = getAllSheetsData(userKey);

    return {
      allSheetsData: allSheetsData,
      userKey: userKey,
      sheetNames: getSheetNames()
    };
  } catch (e) {
    Logger.log(`Kritická chyba v getInitialData: ${e.message}`);
    return null;
  }
}

/**
* Spracuje dáta z formulára - pridá, upraví alebo duplikuje záznam.
* @param {object} formObjects - Objekt s dátami z formulára.
* @returns {object} Aktualizované dáta pre všetkých hárky.
*/
function processForm(formObjects) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(formObjects.sheetName);
  const sheetData = getSheetData(sheet);
  let values = formObjects.values;

  // Spracovanie nahraných súborov
  if (formObjects.files) {
    Object.keys(formObjects.files).forEach(header => {
      const fileInfo = formObjects.files[header];
      const colIndex = sheetData.headers.indexOf(header);
      
      if (colIndex !== -1 && fileInfo.fileName) {
        const data = Utilities.base64Decode(fileInfo.fileData);
        const blob = Utilities.newBlob(data, fileInfo.mimeType, fileInfo.fileName);
        
        // Create structured folder path
        const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
        const sheetFolderName = `att_${formObjects.sheetName.replace('data_', '')}`;
        let sheetFolder;
        
        try {
          sheetFolder = mainFolder.getFoldersByName(sheetFolderName).next();
        } catch (e) {
          sheetFolder = mainFolder.createFolder(sheetFolderName);
        }
        
        let recordFolder;
        const recordId = formObjects.RecId || values[0];
        
        try {
          recordFolder = sheetFolder.getFoldersByName(recordId).next();
        } catch (e) {
          recordFolder = sheetFolder.createFolder(recordId);
        }
        
        const newFile = recordFolder.createFile(blob);
        values[colIndex] = newFile.getUrl();
      }
    });
  }

  // Rozhodne, či ide o pridanie/duplikáciu nového riadku alebo úpravu existujúceho
  if (formObjects.action === 'add' || formObjects.action === 'duplicate' || !formObjects.RecId || !sheetData.idList.includes(formObjects.RecId.toString())) {
    // Generate 10-digit timestamp ID
    values[0] = Math.floor(new Date().getTime() / 1000).toString();
    sheet.appendRow(values);
  } else {
    values[0] = formObjects.RecId;
    const rowIndex = sheetData.idList.indexOf(formObjects.RecId.toString()) + CONFIG.HEADER_ROW + 1;
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  }

  return getAllSheetsData(formObjects.key);
}

/**
* Vymaže záznam z hárku.
* @param {object} objRecord - Objekt s informáciami o zázname na vymazanie.
* @returns {object} Aktualizované dáta pre všetkých hárky.
*/
function deleteData(objRecord) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(objRecord.sheetName);
  const sheetData = getSheetData(sheet);

  if (!sheetData) return getAllSheetsData(objRecord.key);

  const position = sheetData.idList.indexOf(objRecord.recordId.toString());
  if (position > -1) {
    sheet.deleteRow(position + CONFIG.HEADER_ROW + 1);
  }

  return getAllSheetsData(objRecord.key);
}

/**
* Načíta prihlasovacie údaje z hárku 'Set up login'.
* @returns {Array<Array<string>>} Pole s prihlasovacími menami a heslami.
*/
function getLoginData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(CONFIG.LOGIN_SHEET);
    if (!ws || ws.getLastRow() < 2) return [];
    return ws.getRange(2, 1, ws.getLastRow() - 1, 2).getValues();
  } catch (e) {
    Logger.log(`Chyba pri načítaní prihlasovacích údajov: ${e.toString()}`);
    return [];
  }
}

/**
* Exportuje dáta do formátu CSV
* @param {string} sheetName - Názov hárku
* @param {Array} headers - Hlavičky tabuľky
* @param {Array} data - Dáta tabuľky
* @param {Array} selectedRows - Indexy vybraných riadkov
* @param {Array} visibleColumns - Indexy viditeľných stĺpcov
* @returns {string} CSV obsah
*/
function exportToCSV(sheetName, headers, data, selectedRows, visibleColumns) {
  let csvContent = "";
  
  // Filter and map headers
  const filteredHeaders = headers.filter((_, index) => visibleColumns.includes(index + 3)); // +3 for checkbox, actions, duplicate columns
  csvContent += filteredHeaders.join(",") + "\r\n";
  
  // Filter data based on selection and visibility
  const filteredData = selectedRows.length > 0 ? 
    selectedRows.map(rowIndex => data[rowIndex]) : 
    data;
  
  filteredData.forEach(row => {
    const filteredRow = visibleColumns.map(colIndex => {
      const value = row[colIndex - 3]; // Adjust for the extra columns
      return `"${value ? value.toString().replace(/"/g, '""') : ''}"`;
    });
    csvContent += filteredRow.join(",") + "\r\n";
  });
  
  return csvContent;
}

/**
* Exportuje dáta do HTML
* @param {string} sheetName - Názov hárku
* @param {Array} headers - Hlavičky tabuľky
* @param {Array} data - Dáta tabuľky
* @param {Array} selectedRows - Indexy vybraných riadkov
* @param {Array} visibleColumns - Indexy viditeľných stĺpcov
* @param {boolean} includeFooter - Či zahrnúť pätičku
* @returns {string} HTML obsah
*/
function exportToHTML(sheetName, headers, data, selectedRows, visibleColumns, includeFooter) {
  let htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>HTML Export ${sheetName}</title>
      <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .sum-row { font-weight: bold; background-color: #e6e6e6; }
      </style>
    </head>
    <body>
      <h2>${sheetName.replace('data_', '')} - Export</h2>
      <table>
        <thead>
          <tr>
  `;
  
  // Add headers
  visibleColumns.forEach(colIndex => {
    const header = headers[colIndex - 3]; // Adjust for extra columns
    htmlContent += `<th>${header}</th>`;
  });
  htmlContent += `</tr></thead><tbody>`;
  
  // Add data rows
  const filteredData = selectedRows.length > 0 ? 
    selectedRows.map(rowIndex => data[rowIndex]) : 
    data;
  
  filteredData.forEach(row => {
    htmlContent += '<tr>';
    visibleColumns.forEach(colIndex => {
      const value = row[colIndex - 3] || '';
      htmlContent += `<td>${value}</td>`;
    });
    htmlContent += '</tr>';
  });
  
  // Add footer if needed
  if (includeFooter) {
    htmlContent += '<tr class="sum-row"><td colspan="' + (visibleColumns.length - 1) + '">Suma</td>';
    
    // Calculate sums for columns ending with *
    headers.forEach((header, index) => {
      if (header.trim().endsWith('*') && visibleColumns.includes(index + 3)) {
        const sum = filteredData.reduce((acc, row) => {
          const num = parseFloat(row[index]) || 0;
          return acc + num;
        }, 0);
        htmlContent += `<td>${sum.toFixed(2)}</td>`;
      }
    });
    
    htmlContent += '</tr>';
  }
  
  htmlContent += `</tbody></table></body></html>`;
  return htmlContent;
}
