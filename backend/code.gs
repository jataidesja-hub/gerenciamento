/**
 * Sistema de Gerenciamento - Google Apps Script Backend
 * Handles CRUD operations for sales management.
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // Replace with your ID if necessary, or use active spreadsheet
const SHEET_NAME = 'Vendas';

/**
 * Main GET handler for the API.
 */
function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : null;
    
    if (action === 'getSales') {
      return jsonResponse(getSales());
    } else if (action === 'setup') {
      return jsonResponse(setupSheet());
    }
    return jsonResponse({ error: 'Invalid or missing action', parameter: e.parameter }, 400);
  } catch (error) {
    return jsonResponse({ error: error.toString(), stack: error.stack }, 500);
  }
}

/**
 * Main POST handler for the API.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ error: 'No POST data received' }, 400);
    }
    
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === 'saveSale') {
      return jsonResponse(saveSale(data.sale));
    }
    return jsonResponse({ error: 'Invalid action', action: action }, 400);
  } catch (error) {
    return jsonResponse({ error: error.toString(), stack: error.stack }, 500);
  }
}

/**
 * Helper to return JSON responses with CORS headers.
 */
function jsonResponse(data, status = 200) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Setup the spreadsheet structure.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  const headers = [
    'ID da Venda', 'Status do Pagamento', 'Nome do Cliente', 'Cidade/UF', 'Telefone / WhatsApp',
    'Data da Compra', 'Valor Total (R$)', 'Forma de Pagamento', 'Parcelas', 'Valor da Parcela (R$)',
    'Ninhada', 'Sexo', 'Cor', 'Data de Entrega', 'Responsável'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  
  return { success: true, message: 'Sheet configured successfully' };
}

/**
 * Get all sales from the sheet.
 */
function getSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map((row, index) => {
    const sale = { rowIndex: index + 2 };
    headers.forEach((header, i) => {
      sale[header] = row[i];
    });
    return sale;
  });
}

/**
 * Save or update a sale.
 */
function saveSale(sale) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(header => {
    let value = sale[header] || '';
    // Basic numeric conversion for specific fields
    if (header.includes('(R$)') || header === 'Parcelas') {
      const num = parseFloat(value.toString().replace(',', '.'));
      return isNaN(num) ? value : num;
    }
    return value;
  });
  
  if (sale.rowIndex) {
    sheet.getRange(parseInt(sale.rowIndex), 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Generate simple ID if not provided
    if (!rowData[0]) {
      rowData[0] = 'SALE-' + Math.random().toString(36).substr(2, 9).toUpperCase();
    }
    sheet.appendRow(rowData);
  }
  
  return { success: true };
}
