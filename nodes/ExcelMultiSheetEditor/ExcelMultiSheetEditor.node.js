'use strict';

const XLSX = require('xlsx');

class ExcelMultiSheetEditor {
  constructor() {
    this.description = {
      displayName: 'Excel Multi-Sheet Editor',
      name: 'excelMultiSheetEditor',
      group: ['transform'],
      version: 1,
      description: 'Edit multi-sheet Excel files - add rows individually and by ranges',
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
            { name: 'Add Rows to End', value: 'addRows' },
            { name: 'Add Range at Position', value: 'addRange' },
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
          type: 'string',
          default: 'Sheet1',
          description: 'Sheet name to edit. Will be created if it does not exist.',
        },
        {
          displayName: 'Output File Name',
          name: 'outputFileName',
          type: 'string',
          default: '',
          placeholder: 'Leave empty to keep original name',
          description: 'Custom file name for the output Excel file. Leave empty to keep original name.',
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
        },
        {
          displayName: 'JSON Data Property',
          name: 'jsonDataProperty',
          type: 'string',
          default: 'excelData',
          displayOptions: { show: { dataFormat: ['fromJson'] } },
        },
        {
          displayName: 'Data to Add',
          name: 'dataToAdd',
          type: 'json',
          default: '[\n  ["Name", "Age", "City"],\n  ["John", 30, "New York"],\n  ["Jane", 25, "London"]\n]',
          displayOptions: { show: { dataFormat: ['defineBelow'] } },
        },
        {
          displayName: 'Target Columns',
          name: 'targetColumns',
          type: 'string',
          default: '',
          placeholder: 'e.g. A,B,C or 1,2,3',
          description: 'Optional: columns to write to. Empty = auto.',
        },
        {
          displayName: 'Options',
          name: 'options',
          type: 'collection',
          placeholder: 'Add Option',
          default: {},
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

  async execute() {
    const items = this.getInputData();
    const returnData = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const operation = this.getNodeParameter('operation', i);
        const sheetName = this.getNodeParameter('sheetName', i) || 'Sheet1';
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

        // Load workbook
        let workbook;
        let originalFileName = 'workbook.xlsx';
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          workbook = XLSX.read(binaryData, { type: 'buffer' });
          // 🔧 Get original filename
          if (items[i].binary[binaryPropertyName].fileName) {
            originalFileName = items[i].binary[binaryPropertyName].fileName;
          }
        } else {
          workbook = XLSX.utils.book_new();
        }

        // Get or create worksheet
        let worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          worksheet = XLSX.utils.aoa_to_sheet([[]]);
          XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        // 🔧 Get current data range
        const currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        // Determine start position
        let startRow = currentData.length;
        let startCol = 0;

        if (options.appendEmptyRow && startRow > 0) {
          const lastRow = currentData[startRow - 1];
          if (lastRow.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
            startRow++;
          }
        }

        if (operation === 'addRange') {
          const startCell = this.getNodeParameter('startCell', i);
          const startRef = XLSX.utils.decode_cell(startCell);
          startRow = startRef.r;
          startCol = startRef.c;
        }

        // 🔧 WRITE DIRECTLY TO CELLS - preserves existing cells
        let maxColUsed = startCol;
        let maxRowUsed = startRow;

        for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
          const row = dataToInsert[rowIdx];
          
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const targetCol = targetColumns.length > 0 && colIdx < targetColumns.length
              ? parseColumnIndex(targetColumns[colIdx])
              : startCol + colIdx;

            const cellRef = XLSX.utils.encode_cell({ r: startRow + rowIdx, c: targetCol });
            
            // Check overwrite
            if (operation === 'addRange' && !options.overwrite) {
              const existingCell = worksheet[cellRef];
              if (existingCell && existingCell.v !== undefined && existingCell.v !== '') continue;
            }

            // Write new value - 🔧 keep existing formatting if cell exists
            const value = row[colIdx];
            const existingCell = worksheet[cellRef];
            
            worksheet[cellRef] = {
              v: value,
              t: typeof value === 'number' ? 'n' : 's',
            };

            // 🔧 Preserve existing cell style
            if (existingCell && existingCell.s) {
              worksheet[cellRef].s = existingCell.s;
            }
            if (existingCell && existingCell.z) {
              worksheet[cellRef].z = existingCell.z;
            }

            if (targetCol > maxColUsed) maxColUsed = targetCol;
            if (startRow + rowIdx > maxRowUsed) maxRowUsed = startRow + rowIdx;
          }
        }

        // Update range
        const oldRange = worksheet['!ref'] 
          ? XLSX.utils.decode_range(worksheet['!ref']) 
          : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
        
        worksheet['!ref'] = XLSX.utils.encode_range({
          s: { r: Math.min(oldRange.s.r, startRow), c: Math.min(oldRange.s.c, startCol) },
          e: { r: Math.max(oldRange.e.r, maxRowUsed), c: Math.max(oldRange.e.c, maxColUsed) },
        });

        workbook.Sheets[sheetName] = worksheet;

        // Write to buffer
        const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        // 🔧 Determine output filename
        const finalFileName = outputFileName 
          ? (outputFileName.endsWith('.xlsx') ? outputFileName : `${outputFileName}.xlsx`)
          : originalFileName;

        // Prepare output
        const outputBinaryField = options.outputBinaryField || binaryPropertyName || 'data';
        const newItem = {
          json: {
            ...items[i].json,
            excelOperation: operation,
            sheetModified: sheetName,
            rowsAffected: dataToInsert.length,
            sheetNames: workbook.SheetNames,
            outputFileName: finalFileName,
          },
          binary: {},
        };

        newItem.binary[outputBinaryField] = await this.helpers.prepareBinaryData(
          wbout,
          finalFileName,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );

        returnData.push(newItem);
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