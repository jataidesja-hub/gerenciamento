/**
 * Sistema de Gerenciamento - Google Apps Script Backend
 * Handles CRUD operations for sales and installment management.
 */

const SHEET_NAME = 'Vendas';
const INSTALLMENTS_SHEET = 'Parcelas';

/**
 * Main GET handler for the API.
 */
function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : null;
    
    if (action === 'getSales') {
      return jsonResponse(getSales());
    } else if (action === 'getInstallments') {
      return jsonResponse(getInstallments());
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
    } else if (action === 'payInstallment') {
      return jsonResponse(payInstallment(data.saleId, data.installmentNumber));
    }
    return jsonResponse({ error: 'Invalid action', action: action }, 400);
  } catch (error) {
    return jsonResponse({ error: error.toString(), stack: error.stack }, 500);
  }
}

/**
 * Helper to return JSON responses.
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
  
  // Setup Vendas sheet
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
  
  // Setup Parcelas sheet
  let installSheet = ss.getSheetByName(INSTALLMENTS_SHEET);
  if (!installSheet) {
    installSheet = ss.insertSheet(INSTALLMENTS_SHEET);
  }
  
  const installHeaders = [
    'ID Venda', 'Nº Parcela', 'Valor (R$)', 'Vencimento', 'Status', 'Data Pagamento'
  ];
  
  installSheet.getRange(1, 1, 1, installHeaders.length).setValues([installHeaders]);
  installSheet.getRange(1, 1, 1, installHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
  installSheet.setFrozenRows(1);
  
  return { success: true, message: 'Sheets configured successfully' };
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
 * Get all installments from the Parcelas sheet.
 */
function getInstallments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INSTALLMENTS_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map((row, index) => {
    const inst = { rowIndex: index + 2 };
    headers.forEach((header, i) => {
      inst[header] = row[i];
    });
    return inst;
  });
}

/**
 * Save or update a sale. Auto-generates installments if applicable.
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
    if (header.includes('(R$)') || header === 'Parcelas') {
      const num = parseFloat(value.toString().replace(',', '.'));
      return isNaN(num) ? value : num;
    }
    return value;
  });
  
  let saleId;
  
  if (sale.rowIndex) {
    // Update existing sale
    sheet.getRange(parseInt(sale.rowIndex), 1, 1, rowData.length).setValues([rowData]);
    saleId = rowData[0];
  } else {
    // New sale — generate ID
    saleId = 'SALE-' + Math.random().toString(36).substr(2, 9).toUpperCase();
    rowData[0] = saleId;
    sheet.appendRow(rowData);
    
    // Auto-generate installments for new sales
    const numInstallments = parseInt(sale['Parcelas']) || 1;
    const installmentValue = parseFloat((sale['Valor da Parcela (R$)'] || '0').toString().replace(',', '.')) || 0;
    const totalValue = parseFloat((sale['Valor Total (R$)'] || '0').toString().replace(',', '.')) || 0;
    const calcValue = installmentValue > 0 ? installmentValue : (numInstallments > 0 ? totalValue / numInstallments : totalValue);
    
    if (numInstallments > 1) {
      generateInstallments(saleId, numInstallments, calcValue, sale['Data da Compra']);
    } else if (numInstallments === 1) {
      // Single payment — auto-set based on status
      const status = sale['Status do Pagamento'] === 'Pago' ? 'Pago' : 'Pendente';
      generateInstallments(saleId, 1, totalValue, sale['Data da Compra'], status);
    }
  }
  
  return { success: true, saleId: saleId };
}

/**
 * Generate installment rows in the Parcelas sheet.
 */
function generateInstallments(saleId, count, value, startDate, forceStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INSTALLMENTS_SHEET);
  
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(INSTALLMENTS_SHEET);
  }
  
  const baseDate = startDate ? new Date(startDate) : new Date();
  const rows = [];
  
  for (let i = 1; i <= count; i++) {
    const dueDate = new Date(baseDate);
    dueDate.setMonth(dueDate.getMonth() + (i - 1));
    
    rows.push([
      saleId,
      i,
      value,
      dueDate,
      forceStatus || 'Pendente',
      forceStatus === 'Pago' ? new Date() : ''
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}

/**
 * Mark a specific installment as paid and update the parent sale status.
 */
function payInstallment(saleId, installmentNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INSTALLMENTS_SHEET);
  
  if (!sheet) return { error: 'Sheet Parcelas not found' };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('ID Venda');
  const numCol = headers.indexOf('Nº Parcela');
  const statusCol = headers.indexOf('Status');
  const dateCol = headers.indexOf('Data Pagamento');
  
  // Find and update the specific installment
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === saleId && data[i][numCol] == installmentNumber) {
      sheet.getRange(i + 1, statusCol + 1).setValue('Pago');
      sheet.getRange(i + 1, dateCol + 1).setValue(new Date());
      break;
    }
  }
  
  // Now check all installments for this sale to update the main sale status
  const allInstallments = data.filter((row, idx) => idx > 0 && row[idCol] === saleId);
  const updatedData = sheet.getDataRange().getValues();
  const currentInstallments = updatedData.filter((row, idx) => idx > 0 && row[idCol] === saleId);
  
  const totalCount = currentInstallments.length;
  const paidCount = currentInstallments.filter(r => r[statusCol] === 'Pago').length;
  
  // Update the sale status
  let newStatus;
  if (paidCount === 0) {
    newStatus = 'Em aberto';
  } else if (paidCount < totalCount) {
    newStatus = 'Parcial';
  } else {
    newStatus = 'Pago';
  }
  
  updateSaleStatus(saleId, newStatus);
  
  return { success: true, paidCount: paidCount, totalCount: totalCount, newStatus: newStatus };
}

/**
 * Update the status of a sale in the Vendas sheet.
 */
function updateSaleStatus(saleId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('ID da Venda');
  const statusCol = headers.indexOf('Status do Pagamento');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === saleId) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      break;
    }
  }
}
