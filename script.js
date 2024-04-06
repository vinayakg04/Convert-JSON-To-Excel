const fs = require('fs');
const XLSX = require('xlsx');

function traverseJson(jsonContent, sheets, prefix = '') {
  if (!jsonContent || typeof jsonContent !== 'object') {
    return;
  }
  if (jsonContent === null) {
    return;
  }
  const currentSheet = {};

  Object.keys(jsonContent).forEach((key) => {
    const sheetContent = jsonContent[key];
    const sheetName = prefix + key;

    if (Array.isArray(sheetContent) && sheetContent.length > 0 && typeof sheetContent[0] === 'object') {
      // Treat arrays of objects as separate sheets with hyperlink
      currentSheet[key] = `=HYPERLINK("#${sheetName}!A1", "SHEET::${sheetName}")`; // Hyperlink to the sheet
      const newSheetName = sheetName; // Keep the current sheet name for arrays
      sheets[newSheetName] = []; // Initialize sheet in sheets object
      sheetContent.forEach((item, index) => {
        const newSheetNameIndex = `${newSheetName}_${index + 1}`;
        traverseJson(item, sheets, newSheetNameIndex + '_');
        sheets[newSheetName].push(...sheets[newSheetNameIndex]);
        delete sheets[newSheetNameIndex];
      });
    } else if (typeof sheetContent === 'object') {
      // Treat nested objects as separate sheets with hyperlink
      currentSheet[key] = `=HYPERLINK("#${sheetName}!A1", "SHEET::${sheetName}")`; // Hyperlink to the sheet
      traverseJson(sheetContent, sheets, sheetName + '_');
    } else {
      // Keep other variables in the same sheet
      currentSheet[key] = sheetContent;
    }
  });

  if (Object.keys(currentSheet).length > 0) {
    // Add current sheet to sheets object
    sheets[prefix.slice(0, -1)] = [currentSheet];
  }
}


function insertInExcel(sheets, outputfile) {
  const wb = XLSX.utils.book_new();

  Object.keys(sheets).forEach((sheetName) => {
    const ws = XLSX.utils.json_to_sheet(sheets[sheetName]);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, outputfile);
}


function readAndtraverseInputFile(filePath, outputfile) {
  try {
    const jsonContent = JSON.parse(fs.readFileSync(filePath, 'utf8'));

    if (!jsonContent) {
      throw new Error('JSON content is Undefined OR Null');
    }

    const sheets = {};
    traverseJson(jsonContent, sheets);
    insertInExcel(sheets, outputfile);
    console.log(`Succesfully written Data to ${outputfile}`);
  } catch (error) {
    console.error('Error:', error.message);
  }
}


readAndtraverseInputFile('input.json', 'assignment_output.xlsx');
