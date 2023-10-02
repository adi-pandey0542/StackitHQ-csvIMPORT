function onInstall(e){
  onOpen(e);
}

// add menu to menubar
function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('NowStackIt')
    .addItem('Import CSV', 'importCSVModal')
    .addToUi();
}

// create modal
function importCSVModal(){
  var html = HtmlService.createHtmlOutputFromFile('csvImport')
    .setTitle('Import CSV to sheet')
    .setWidth(500)
    .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Import CSV');
}

// decode base64 data to process further
function decodeCSVData(base64CsvData, selectedColumns, options){
  try{
    const decodedData = Utilities.base64Decode(base64CsvData);
    const csvData = Utilities.newBlob(decodedData).getDataAsString();
    

    const processedData = processCSVData(csvData, selectedColumns, options);

    if(processedData){  
      SpreadsheetApp.getUi().alert  ('CSV data imported successfully.');
      applyFilters(options);       
    }else{
      SpreadsheetApp.getUi().alert('Cannot import csv data at the moment.');      
    }

    
  }catch(err){
    return 'Error: ' + err.toString();
  }
}

// process decoded data and add data to sheet
function processCSVData(csvData, selectedColumns, options){   
  const parsedData = Utilities.parseCsv(csvData);
  const headers = parsedData[0];

  const selectedColumnIndices = selectedColumns.map(col => headers.indexOf(col));

  const finalData = parsedData.map(row => selectedColumnIndices.map(index => row[index]));
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  var targetSheet;  

  if(options.action === "create"){    
    targetSheet = options.newSheet;
    ss.insertSheet(targetSheet);
  }else{
    targetSheet = ss.getActiveSheet().getName();
  } 
  // SpreadsheetApp.getUi().alert(targetSheet);
  const sheet = ss.getSheetByName(targetSheet);  

  if(sheet){    

    if(options.action === "replace"){
      sheet.clear();
    }

    // if(options.action === 'append'){      
    //   // TODO: append without header
    //   // sheet.getRange(sheet.getLastRow() + 1, 1, finalData.length, selectedColumns.length).setValues(finalData.slice(1));      
    //   sheet.getRange(sheet.getLastRow() + 1, 1, finalData.length, selectedColumns.length).setValues(finalData); 
    //   return true;
    // }

    sheet.getRange(sheet.getLastRow() + 1, 1, finalData.length , selectedColumns.length).setValues(finalData);

    return true

  }else{
    return false
  }
}

// function applyFilters(options){  
//     if(options.removeDupes){
//      sheet.getDataRange().removeDuplicates();
//     }
// }


