const XLSX = require('xlsx');
const { google } = require('googleapis');
const path = require('path');

const CONFIG = {
  SHEET_ID: process.env.GOOGLE_SHEET_ID,
  SHEET_NAME: 'DataUtama',
  MAX_ROWS: 30000,           
  MIN_ROWS: 10,
  BATCH_SIZE: 100,           
  FORMULA_BATCH_SIZE: 50,    
  REQUEST_DELAY: 100         
};

function getGoogleSheetsClient() {
  // Parsing credentials dari Environment Variable Vercel
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: credentials.client_email,
      private_key: credentials.private_key.replace(/\\n/g, "\n"), // Memastikan format enter pada key terbaca
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

function parseExcelFile(fileBuffer) {
  try {
    console.log('📊 Parsing Excel file...');
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1,
      defval: '',
      raw: false
    });
    
    const filtered = data.filter(row => 
      row.some(cell => cell !== null && cell !== undefined && cell !== '')
    );
    
    console.log(`✓ Parsed ${filtered.length - 1} data rows`);
    return filtered;
  } catch (error) {
    throw new Error(`Parse error: ${error.message}`);
  }
}

async function getExistingIdProyek(sheets, idProyekColumnIndex) {
  try {
    console.log('🔍 Fetching existing Id Proyek (optimized)...');
    const startTime = Date.now();
    
    const colLetter = getColumnLetter(idProyekColumnIndex);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEET_ID,
      range: `${CONFIG.SHEET_NAME}!${colLetter}2:${colLetter}`,
      valueRenderOption: 'UNFORMATTED_VALUE',
      majorDimension: 'COLUMNS'
    });
    
    const column = response.data.values ? response.data.values[0] : [];
    
    const existingIds = new Set(
      column
        .filter(id => id !== null && id !== undefined && id !== '')
        .map(id => String(id).trim().toLowerCase())
    );
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`✓ Loaded ${existingIds.size} existing Id Proyek in ${elapsed}s`);
    
    return existingIds;
    
  } catch (error) {
    console.error('Error getting existing Id Proyek:', error);
    return new Set();
  }
}

async function getLastNumber(sheets) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEET_ID,
      range: `${CONFIG.SHEET_NAME}!A:A`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const rows = response.data.values || [];
    
    console.log(`Total rows in sheet: ${rows.length}`);
    
    if (rows.length <= 1) {
      console.log('No data rows, starting from 0');
      return 0;
    }
    
    let lastNumber = 0;
    
    // Scan from bottom to top (faster for large datasets)
    for (let i = rows.length - 1; i >= 1; i--) {
      const value = rows[i][0];
      
      if (value === null || value === undefined || value === '') continue;
      
      let num;
      if (typeof value === 'number') {
        num = value;
      } else {
        const cleaned = String(value).replace(/[,.\s]/g, '');
        num = parseInt(cleaned);
      }
      
      if (!isNaN(num) && num > 0 && num < 100000000) {
        lastNumber = num;
        console.log(`✓ Last number at row ${i + 1}: ${lastNumber}`);
        break;
      }
    }
    
    return lastNumber;
    
  } catch (error) {
    console.error('Error getting last number:', error);
    return 0;
  }
}

async function getSheetHeaders(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SHEET_ID,
    range: `${CONFIG.SHEET_NAME}!1:1`,
  });
  
  const headers = response.data.values ? response.data.values[0] : [];
  if (headers.length === 0) throw new Error('Sheet tidak memiliki header');
  
  return headers.map(h => String(h).trim());
}

async function getSheetId(sheets) {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId: CONFIG.SHEET_ID,
    });
    
    const sheet = response.data.sheets.find(
      s => s.properties.title === CONFIG.SHEET_NAME
    );
    
    return sheet ? sheet.properties.sheetId : 0;
  } catch (error) {
    console.error('Error getting sheet ID:', error);
    return 0;
  }
}

function getColumnLetter(index) {
  let letter = '';
  let num = index;
  
  while (num >= 0) {
    letter = String.fromCharCode((num % 26) + 65) + letter;
    num = Math.floor(num / 26) - 1;
  }
  
  return letter;
}

async function autoFillFormulas(sheets, sourceRow, startRow, endRow) {
  try {
    console.log(`📐 Copying formulas from row ${sourceRow} to rows ${startRow}-${endRow}...`);
    const startTime = Date.now();
    
    const sourceResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEET_ID,
      range: `${CONFIG.SHEET_NAME}!${sourceRow}:${sourceRow}`,
      valueRenderOption: 'FORMULA'
    });
    
    if (!sourceResponse.data.values || sourceResponse.data.values.length === 0) {
      console.log('No formulas in source row');
      return;
    }
    
    const sourceFormulas = sourceResponse.data.values[0];
    const formulaColumns = [];
    
    sourceFormulas.forEach((cell, index) => {
      if (typeof cell === 'string' && cell.startsWith('=')) {
        formulaColumns.push({
          index: index,
          formula: cell,
          column: getColumnLetter(index)
        });
      }
    });
    
    if (formulaColumns.length === 0) {
      console.log('No formula columns detected');
      return;
    }
    
    console.log(`Found ${formulaColumns.length} formula columns:`, 
      formulaColumns.map(c => c.column).join(', '));
    
    const totalRows = endRow - startRow + 1;
    const batchSize = CONFIG.FORMULA_BATCH_SIZE;

    for (const col of formulaColumns) {
      const colLetter = col.column;
      const baseFormula = col.formula;
     
      const sourceRowMatch = baseFormula.match(/\d+/);
      const sourceRowNum = sourceRowMatch ? parseInt(sourceRowMatch[0]) : sourceRow;
      
      console.log(`  Processing column ${colLetter}...`);
     
      for (let batchStart = startRow; batchStart <= endRow; batchStart += batchSize) {
        const batchEnd = Math.min(batchStart + batchSize - 1, endRow);
        const batchRows = [];
    
        for (let targetRow = batchStart; targetRow <= batchEnd; targetRow++) {
          const rowRegex = new RegExp(`\\b${sourceRowNum}\\b`, 'g');
          const adjustedFormula = baseFormula.replace(rowRegex, targetRow);
          batchRows.push([adjustedFormula]);
        }

        const batchRange = `${CONFIG.SHEET_NAME}!${colLetter}${batchStart}:${colLetter}${batchEnd}`;
        
        await sheets.spreadsheets.values.update({
          spreadsheetId: CONFIG.SHEET_ID,
          range: batchRange,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: batchRows
          }
        });
        
        if (batchEnd < endRow) {
          console.log(`    ✓ Applied ${batchEnd - batchStart + 1} formulas (${batchEnd - startRow + 1}/${totalRows})`);
        }
      }
      
      console.log(`  ✓ Column ${colLetter} complete (${totalRows} cells)`);
    }
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`✓ All formulas applied in ${elapsed}s`);
    
  } catch (error) {
    console.error('Error auto-filling formulas:', error);
  }
}

function mapHeaders(fileHeaders, sheetHeaders) {
  return fileHeaders.map(fileHeader => {
    const fh = String(fileHeader).trim().toLowerCase();
    return sheetHeaders.findIndex(sh => 
      sh.toLowerCase() === fh
    );
  });
}

function processValue(value, columnName) {
  if (!value || value === '') return '';
  
  const col = String(columnName).toLowerCase();
  
  if (col.includes('jumlah investasi')) return parseNumber(value);
  if (col.includes('latitude') || col.includes('longitude')) return parseNumber(value);
  if (col.includes('tanggal') || col.includes('day of')) return formatDate(value);
  if (col.includes('tki') || col.includes('luas')) {
    const num = parseNumber(value);
    return num !== 0 ? num : value;
  }
  
  return String(value);
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const str = String(value).replace(/[^\d.-]/g, '');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function formatDate(value) {
  try {
    let date;
    if (value instanceof Date) {
      date = value;
    } else if (typeof value === 'number') {
      const d = XLSX.SSF.parse_date_code(value);
      date = new Date(d.y, d.m - 1, d.d);
    } else {
      date = new Date(value);
    }
    
    if (!isNaN(date.getTime())) {
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const year = date.getFullYear();
      return `${day}/${month}/${year}`;
    }
  } catch (e) {}
  
  return value;
}

async function formatColumns(sheets, updatedRange, headers) {
  try {
    const match = updatedRange.match(/!A(\d+):.*(\d+)/);
    if (!match) return;
    
    const startRow = parseInt(match[1]);
    const endRow = parseInt(match[2]);
    
    const sheetId = await getSheetId(sheets);
    if (!sheetId) return;
    
    const jumlahCol = headers.findIndex(h => 
      h.toLowerCase().includes('jumlah investasi')
    );
    
    if (jumlahCol === -1) return;
    
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.SHEET_ID,
      requestBody: {
        requests: [{
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: startRow - 1,
              endRowIndex: endRow,
              startColumnIndex: jumlahCol,
              endColumnIndex: jumlahCol + 1,
            },
            cell: {
              userEnteredFormat: {
                numberFormat: {
                  type: 'CURRENCY',
                  pattern: '"Rp "#,##0',
                },
              },
            },
            fields: 'userEnteredFormat.numberFormat',
          },
        }],
      },
    });
    
    console.log('✓ Currency format applied');
  } catch (error) {
    console.error('Format error:', error);
  }
}

async function batchAppendData(sheets, data, sheetHeaders, onProgress) {
  const totalRows = data.length;
  const batchSize = CONFIG.BATCH_SIZE;
  let processedRows = 0;
  let startRow = 0;
  let endRow = 0;
  
  console.log(`📤 Uploading ${totalRows.toLocaleString()} rows in batches of ${batchSize}...`);
  
  for (let i = 0; i < totalRows; i += batchSize) {
    const batch = data.slice(i, Math.min(i + batchSize, totalRows));
  
    if (i > 0 && CONFIG.REQUEST_DELAY) {
      await new Promise(resolve => setTimeout(resolve, CONFIG.REQUEST_DELAY));
    }
    
    const response = await sheets.spreadsheets.values.append({
      spreadsheetId: CONFIG.SHEET_ID,
      range: `${CONFIG.SHEET_NAME}!A:A`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: batch },
    });
    
    processedRows += batch.length;
    
    if (i === 0) {
      const match = response.data.updates.updatedRange.match(/!A(\d+)/);
      startRow = match ? parseInt(match[1]) : 2;
    }
    
    if (i + batchSize >= totalRows) {
      const match = response.data.updates.updatedRange.match(/!A(\d+):.*(\d+)/);
      endRow = match ? parseInt(match[2]) : startRow + totalRows - 1;
    }
  
    if (onProgress) {
      const progress = Math.round((processedRows / totalRows) * 100);
      onProgress(processedRows, totalRows, progress);
    }
    
    console.log(`  ✓ Uploaded ${processedRows.toLocaleString()}/${totalRows.toLocaleString()} rows (${Math.round(processedRows/totalRows*100)}%)`);
  }
  
  return { startRow, endRow };
}

async function importToGoogleSheet(file, onProgress = null) {
  const overallStart = Date.now();
  
  try {
    console.log('\n========================================');
    console.log('🚀 STARTING OPTIMIZED IMPORT');
    console.log('========================================\n');
    
    console.log('1️⃣ Parsing Excel...');
    const rawData = parseExcelFile(file.buffer);
    const dataRowCount = rawData.length - 1;
    
    if (dataRowCount < CONFIG.MIN_ROWS) {
      throw new Error(`Min ${CONFIG.MIN_ROWS} rows required`);
    }
    if (dataRowCount > CONFIG.MAX_ROWS) {
      throw new Error(`Max ${CONFIG.MAX_ROWS} rows allowed. Please split the file.`);
    }
    
    console.log('\n2️⃣ Connecting to Google Sheets...');
    const sheets = getGoogleSheetsClient();

    console.log('\n3️⃣ Reading sheet structure...');
    const sheetHeaders = await getSheetHeaders(sheets);
    console.log(`✓ Found ${sheetHeaders.length} columns`);
    
    const fileHeaders = rawData[0].map(h => String(h).trim());
    const dataRows = rawData.slice(1);

    console.log('\n4️⃣ Mapping columns...');
    const columnMapping = mapHeaders(fileHeaders, sheetHeaders);
    const mappedCount = columnMapping.filter(i => i !== -1).length;
    console.log(`✓ Mapped ${mappedCount}/${fileHeaders.length} columns`);
  
    console.log('\n5️⃣ Getting last number...');
    let lastNumber = await getLastNumber(sheets);
    
    console.log('\n6️⃣ Setting up duplicate detection...');
    const idProyekColIndex = sheetHeaders.findIndex(h => {
      const normalized = h.toLowerCase().trim().replace(/[_\s-]/g, '');
      return normalized === 'idproyek';
    });
    
    const fileIdProyekColIndex = fileHeaders.findIndex(h => {
      const normalized = h.toLowerCase().trim().replace(/[_\s-]/g, '');
      return normalized === 'idproyek';
    });
    
    if (idProyekColIndex === -1 || fileIdProyekColIndex === -1) {
      console.log('⚠️ WARNING: Id Proyek column not found - duplicate detection disabled');
    } else {
      console.log(`✓ Id Proyek: Sheet column ${idProyekColIndex + 1}, File column ${fileIdProyekColIndex + 1}`);
    }

    const existingIdProyek = (idProyekColIndex !== -1) 
      ? await getExistingIdProyek(sheets, idProyekColIndex)
      : new Set();

    console.log('\n7️⃣ Processing data and filtering duplicates...');
    const processStart = Date.now();
    
    const processedData = [];
    const duplicateDetails = [];
    let duplicatesSkipped = 0;
    let currentNumber = lastNumber;
    
    dataRows.forEach((row, idx) => {

      let currentIdProyek = null;
      
      if (fileIdProyekColIndex !== -1 && row[fileIdProyekColIndex]) {
        currentIdProyek = String(row[fileIdProyekColIndex]).trim().toLowerCase();
      }
      
      if (currentIdProyek && existingIdProyek.has(currentIdProyek)) {
        duplicatesSkipped++;
        duplicateDetails.push({
          rowNumber: idx + 2,
          idProyek: row[fileIdProyekColIndex]
        });
        return; 
      }
      
      const newRow = new Array(sheetHeaders.length).fill('');
      newRow[0] = ++currentNumber;
      
      row.forEach((cell, i) => {
        const sheetCol = columnMapping[i];
        if (sheetCol === -1 || sheetCol === 0) return;
        newRow[sheetCol] = processValue(cell, sheetHeaders[sheetCol]);
      });
      
      processedData.push(newRow);
      
      if (processedData.length % 1000 === 0) {
        console.log(`  ✓ Processed: ${processedData.length.toLocaleString()} rows...`);
      }
    });
    
    lastNumber = currentNumber;
    
    const processElapsed = ((Date.now() - processStart) / 1000).toFixed(2);
    console.log(`✓ Processed ${processedData.length.toLocaleString()} rows in ${processElapsed}s`);
    
    if (duplicatesSkipped > 0) {
      console.log(`⚠️ Skipped ${duplicatesSkipped.toLocaleString()} duplicates`);
    }
    
    if (processedData.length === 0) {
      throw new Error(`No new data to import. All ${duplicatesSkipped} rows are duplicates.`);
    }
    
    console.log('\n8️⃣ Uploading to Google Sheets...');
    const { startRow, endRow } = await batchAppendData(
      sheets, 
      processedData, 
      sheetHeaders,
      onProgress
    );
    
    console.log(`✓ Written to rows ${startRow.toLocaleString()}-${endRow.toLocaleString()}`);
    
    console.log('\n9️⃣ Applying formulas...');
    if (startRow > 2) {
      await autoFillFormulas(sheets, startRow - 1, startRow, endRow);
    }
    
    console.log('\n🔟 Formatting columns...');
    const updatedRange = `${CONFIG.SHEET_NAME}!A${startRow}:${getColumnLetter(sheetHeaders.length - 1)}${endRow}`;
    await formatColumns(sheets, updatedRange, sheetHeaders);
    
    const overallElapsed = ((Date.now() - overallStart) / 1000).toFixed(2);
    
    console.log('\n========================================');
    console.log('✅ IMPORT COMPLETED SUCCESSFULLY!');
    console.log('========================================');
    console.log(`📊 Total Time: ${overallElapsed}s`);
    console.log(`✅ New Rows: ${processedData.length.toLocaleString()}`);
    console.log(`⚠️ Duplicates Skipped: ${duplicatesSkipped.toLocaleString()}`);
    console.log(`🔢 Number Range: ${(lastNumber - processedData.length + 1).toLocaleString()} - ${lastNumber.toLocaleString()}`);
    console.log('========================================\n');
    
    return {
      rowsImported: processedData.length,
      duplicatesSkipped: duplicatesSkipped,
      duplicateDetails: duplicateDetails.slice(0, 100),
      updatedRange: updatedRange,
      startRow: startRow,
      endRow: endRow,
      lastNumber: lastNumber,
      totalTime: overallElapsed
    };
    
  } catch (error) {
    const elapsed = ((Date.now() - overallStart) / 1000).toFixed(2);
    console.error(`\n❌ IMPORT FAILED after ${elapsed}s:`, error.message);
    throw error;
  }
}

module.exports = { 
  importToGoogleSheet,
  parseExcelFile
};