'use strict';

const ExcelJS = require('exceljs');

class ExcelMultiSheetEditor {
  constructor() {
    this.description = {
      displayName: 'Excel Multi-Sheet Editor',
      name: 'excelMultiSheetEditor',
      group: ['transform'],
      version: 2,
      description: 'Edit multi-sheet Excel files - read, write, list sheets with full style preservation',
      defaults: {
        name: 'Excel Multi-Sheet Editor',
      },
      inputs: ['main'],
      outputs: ['main'],
      properties: [
        {
          displayName: 'Operation',
          name: 'operation',
          type: 'options',
          options: [
            { name: 'Add Rows to End', value: 'addRows', description: 'Add rows to the end of a sheet' },
            { name: 'Add Range at Position', value: 'addRange', description: 'Add data in a specific range/cell position' },
            { name: 'Read Sheet', value: 'readSheet', description: 'Read all data from a sheet' },
            { name: 'List Sheets', value: 'listSheets', description: 'List all sheet names in the workbook' },
          ],
          default: 'addRows',
        },
        {
          displayName: 'Input Binary Field',
          name: 'binaryPropertyName',
          type: 'string',
          default: 'data',
          description: 'Name of the binary property containing the Excel file',
        },
        {
          displayName: 'Sheet Name',
          name: 'sheetName',
          type: 'options',
          typeOptions: {
            loadOptionsMethod: 'getSheetNames',
          },
          default: '',
          displayOptions: {
            hide: {
              operation: ['listSheets'],
            },
          },
          description: 'Choose an existing sheet or type a new name to create one',
        },
        {
          displayName: 'Output File Name',
          name: 'outputFileName',
          type: 'string',
          default: '',
          placeholder: 'Leave empty to keep original name',
          displayOptions: {
            hide: {
              operation: ['readSheet', 'listSheets'],
            },
          },
          description: 'Custom file name for the output Excel file',
        },
        {
          displayName: 'Start Cell',
          name: 'startCell',
          type: 'string',
          default: 'A1',
          displayOptions: { show: { operation: ['addRange'] } },
        },
        {
          displayName: 'Data Format',
          name: 'dataFormat',
          type: 'options',
          options: [
            { name: 'Define Below', value: 'defineBelow' },
            { name: 'From JSON Property', value: 'fromJson' },
          ],
          default: 'defineBelow',
          displayOptions: { show: { operation: ['addRows', 'addRange'] } },
        },
        {
          displayName: 'JSON Data Property',
          name: 'jsonDataProperty',
          type: 'string',
          default: 'excelData',
          displayOptions: { show: { dataFormat: ['fromJson'], operation: ['addRows', 'addRange'] } },
        },
        {
          displayName: 'Data to Add',
          name: 'dataToAdd',
          type: 'json',
          default: '[\n  ["Name", "Age", "City"],\n  ["John", 30, "New York"],\n  ["Jane", 25, "London"]\n]',
          displayOptions: { show: { dataFormat: ['defineBelow'], operation: ['addRows', 'addRange'] } },
        },
        {
          displayName: 'Target Columns',
          name: 'targetColumns',
          type: 'string',
          default: '',
          placeholder: 'e.g. A,B,C or 1,2,3',
          displayOptions: { show: { operation: ['addRows', 'addRange'] } },
        },
        {
          displayName: 'Read Options',
          name: 'readOptions',
          type: 'collection',
          placeholder: 'Add Option',
          default: {},
          displayOptions: { show: { operation: ['readSheet'] } },
          options: [
            {
              displayName: 'Include Empty Cells',
              name: 'includeEmpty',
              type: 'boolean',
              default: false,
            },
            {
              displayName: 'First Row as Headers',
              name: 'firstRowHeaders',
              type: 'boolean',
              default: false,
            },
            {
              displayName: 'Max Rows',
              name: 'maxRows',
              type: 'number',
              default: 0,
            },
          ],
        },
        {
          displayName: 'Options',
          name: 'options',
          type: 'collection',
          placeholder: 'Add Option',
          default: {},
          displayOptions: { show: { operation: ['addRows', 'addRange'] } },
          options: [
            {
              displayName: 'Overwrite Existing Cells',
              name: 'overwrite',
              type: 'boolean',
              default: false,
              displayOptions: { show: { '/operation': ['addRange'] } },
            },
            {
              displayName: 'Add Empty Row Before',
              name: 'appendEmptyRow',
              type: 'boolean',
              default: false,
              displayOptions: { show: { '/operation': ['addRows'] } },
            },
            {
              displayName: 'Output Binary Field',
              name: 'outputBinaryField',
              type: 'string',
              default: 'data',
            },
          ],
        },
      ],
    };
  }

  // 🔧 Dynamic sheet name loader
  methods = {
    loadOptions: {
      async getSheetNames() {
        const items = this.getInputData();
        
        if (!items || items.length === 0) {
          return [{ name: 'Sheet1', value: 'Sheet1' }];
        }

        try {
          const binaryPropertyName = this.getNodeParameter('binaryPropertyName', 0);
          
          if (items[0].binary && items[0].binary[binaryPropertyName]) {
            const workbook = new ExcelJS.Workbook();
            const binaryData = await this.helpers.getBinaryDataBuffer(0, binaryPropertyName);
            await workbook.xlsx.load(binaryData);
            
            const sheets = workbook.worksheets.map(ws => ({
              name: ws.name,
              value: ws.name,
            }));
            
            // Add option to create new sheet
            sheets.push({
              name: '➕ Create new sheet...',
              value: '__new__',
            });
            
            return sheets;
          }
        } catch (error) {
          // Fall through to defaults
        }

        return [
          { name: 'Sheet1', value: 'Sheet1' },
          { name: '➕ Create new sheet...', value: '__new__' },
        ];
      },
    },
  };

  async execute() {
    const items = this.getInputData();
    const returnData = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const operation = this.getNodeParameter('operation', i);
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        // Load workbook
        const workbook = new ExcelJS.Workbook();
        let originalFileName = 'workbook.xlsx';

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          await workbook.xlsx.load(binaryData);
          
          if (items[i].binary[binaryPropertyName].fileName) {
            originalFileName = items[i].binary[binaryPropertyName].fileName;
          }
        } else {
          workbook.addWorksheet('Sheet1');
        }

        // ==========================================
        // OPERATION: LIST SHEETS
        // ==========================================
        if (operation === 'listSheets') {
          const sheetNames = workbook.worksheets.map(ws => ws.name);
          
          returnData.push({
            json: {
              ...items[i].json,
              operation: 'listSheets',
              sheetCount: sheetNames.length,
              sheetNames: sheetNames,
              sheets: sheetNames.map((name, index) => ({
                index: index + 1,
                name: name,
              })),
            },
          });
          continue;
        }

        // 🔧 Get sheet name - handle "__new__" for creating new sheets
        let sheetName = this.getNodeParameter('sheetName', i) || 'Sheet1';
        if (sheetName === '__new__') {
          sheetName = `Sheet${workbook.worksheets.length + 1}`;
        }

        // ==========================================
        // OPERATION: READ SHEET
        // ==========================================
        if (operation === 'readSheet') {
          const worksheet = workbook.getWorksheet(sheetName);
          const readOptions = this.getNodeParameter('readOptions', i, {});
          
          if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
          }

          const data = [];
          const rowCount = worksheet.rowCount;
          const maxRows = readOptions.maxRows > 0 ? Math.min(readOptions.maxRows, rowCount) : rowCount;

          if (readOptions.firstRowHeaders && rowCount > 0) {
            const headerRow = worksheet.getRow(1);
            const headers = [];
            
            headerRow.eachCell({ includeEmpty: readOptions.includeEmpty }, (cell, colNumber) => {
              headers[colNumber] = cell.value !== null && cell.value !== undefined ? String(cell.value) : `Column${colNumber}`;
            });

            for (let rowIdx = 2; rowIdx <= maxRows; rowIdx++) {
              const row = worksheet.getRow(rowIdx);
              const rowData = {};
              
              row.eachCell({ includeEmpty: readOptions.includeEmpty }, (cell, colNumber) => {
                if (headers[colNumber]) {
                  rowData[headers[colNumber]] = cell.value;
                }
              });
              
              if (Object.keys(rowData).length > 0 || readOptions.includeEmpty) {
                data.push(rowData);
              }
            }
          } else {
            for (let rowIdx = 1; rowIdx <= maxRows; rowIdx++) {
              const row = worksheet.getRow(rowIdx);
              const rowData = [];
              
              row.eachCell({ includeEmpty: readOptions.includeEmpty }, (cell, colNumber) => {
                rowData[colNumber - 1] = cell.value;
              });
              
              if (rowData.length > 0 || readOptions.includeEmpty) {
                data.push(rowData);
              }
            }
          }

          returnData.push({
            json: {
              ...items[i].json,
              operation: 'readSheet',
              sheetName: sheetName,
              rowCount: worksheet.rowCount,
              columnCount: worksheet.columnCount,
              data: data,
            },
          });
          continue;
        }

        // ==========================================
        // OPERATION: ADD ROWS / ADD RANGE
        // ==========================================
        const outputFileName = this.getNodeParameter('outputFileName', i);
        const dataFormat = this.getNodeParameter('dataFormat', i);
        const targetColumnsStr = this.getNodeParameter('targetColumns', i);
        const options = this.getNodeParameter('options', i, {});

        // Parse target columns
        let targetColumns = [];
        if (targetColumnsStr && targetColumnsStr.trim()) {
          targetColumns = targetColumnsStr.split(',').map(col => col.trim().toUpperCase());
        }

        // Get data
        let dataToInsert;
        if (dataFormat === 'defineBelow') {
          dataToInsert = this.getNodeParameter('dataToAdd', i);
        } else {
          const jsonProperty = this.getNodeParameter('jsonDataProperty', i);
          dataToInsert = items[i].json[jsonProperty];
        }

        if (typeof dataToInsert === 'string') {
          try { dataToInsert = JSON.parse(dataToInsert); }
          catch (e) { throw new Error(`Invalid JSON: ${e.message}`); }
        }

        if (!Array.isArray(dataToInsert)) {
          throw new Error(`Data must be array, got ${typeof dataToInsert}`);
        }

        dataToInsert = dataToInsert.map(row => {
          if (Array.isArray(row)) return row;
          if (typeof row === 'object' && row !== null) return Object.values(row);
          return [row];
        });

        // Get or create worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          worksheet = workbook.addWorksheet(sheetName);
        }

        // Determine start position
        let startRow;
        let startCol = 1;

        if (operation === 'addRows') {
          startRow = worksheet.rowCount + 1;
          
          if (options.appendEmptyRow && worksheet.rowCount > 0) {
            const lastRow = worksheet.getRow(worksheet.rowCount);
            const hasContent = Array.isArray(lastRow.values) && 
              lastRow.values.some(v => v !== null && v !== undefined && v !== '');
            if (hasContent) startRow++;
          }
        } else {
          const startCell = this.getNodeParameter('startCell', i);
          const parsed = parseCellRef(startCell);
          startRow = parsed.row;
          startCol = parsed.col;
        }

        // Add data
        for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
          const row = dataToInsert[rowIdx];
          const targetRow = worksheet.getRow(startRow + rowIdx);
          
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const targetCol = targetColumns.length > 0 && colIdx < targetColumns.length
              ? parseColumnIndex(targetColumns[colIdx]) + 1
              : startCol + colIdx;

            if (operation === 'addRange' && !options.overwrite) {
              const existingCell = targetRow.getCell(targetCol);
              if (existingCell.value !== null && existingCell.value !== undefined && existingCell.value !== '') {
                continue;
              }
            }

            targetRow.getCell(targetCol).value = row[colIdx];
          }

          targetRow.commit();
        }

        // Write to buffer
        const buffer = await workbook.xlsx.writeBuffer();

        // Determine output filename
        const finalFileName = outputFileName 
          ? (outputFileName.endsWith('.xlsx') ? outputFileName : `${outputFileName}.xlsx`)
          : originalFileName;

        // Prepare output
        const outputBinaryField = options.outputBinaryField || binaryPropertyName || 'data';
        
        returnData.push({
          json: {
            ...items[i].json,
            excelOperation: operation,
            sheetModified: sheetName,
            rowsAffected: dataToInsert.length,
            sheetNames: workbook.worksheets.map(ws => ws.name),
            outputFileName: finalFileName,
          },
          binary: {
            [outputBinaryField]: await this.helpers.prepareBinaryData(
              Buffer.from(buffer),
              finalFileName,
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ),
          },
        });

      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        if (this.continueOnFail()) {
          returnData.push({
            json: { error: errorMessage, ...items[i].json },
          });
          continue;
        }
        throw error;
      }
    }

    return [returnData];
  }
}

function parseCellRef(cellRef) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (match) {
    return {
      col: parseColumnIndex(match[1]) + 1,
      row: parseInt(match[2]),
    };
  }
  return { col: 1, row: 1 };
}

function parseColumnIndex(col) {
  if (/^[A-Z]+$/.test(col)) {
    let idx = 0;
    for (let i = 0; i < col.length; i++) {
      idx = idx * 26 + (col.charCodeAt(i) - 64);
    }
    return idx - 1;
  }
  return parseInt(col) - 1;
}

module.exports = { ExcelMultiSheetEditor };