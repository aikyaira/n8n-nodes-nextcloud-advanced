'use strict';

const ExcelJS = require('exceljs');

class ExcelMultiSheetEditor {
  constructor() {
    this.description = {
      displayName: 'Excel Multi-Sheet Editor',
      name: 'excelMultiSheetEditor',
      group: ['transform'],
      version: 2,
      description: 'Edit multi-sheet Excel files with full style preservation',
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

        // Load workbook with ExcelJS (preserves ALL styles)
        const workbook = new ExcelJS.Workbook();
        let originalFileName = 'workbook.xlsx';
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          await workbook.xlsx.load(binaryData);
          
          if (items[i].binary[binaryPropertyName].fileName) {
            originalFileName = items[i].binary[binaryPropertyName].fileName;
          }
        } else {
          workbook.addWorksheet('Sheet1');
        }

        // Get or create worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          worksheet = workbook.addWorksheet(sheetName);
        }

        // Determine start position
        let startRow;
        let startCol = 1; // ExcelJS uses 1-based indexing

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

        // 🔧 ADD DATA - ExcelJS preserves all existing cell styles
        for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
          const row = dataToInsert[rowIdx];
          const targetRow = worksheet.getRow(startRow + rowIdx);
          
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const targetCol = targetColumns.length > 0 && colIdx < targetColumns.length
              ? parseColumnIndex(targetColumns[colIdx]) + 1 // ExcelJS 1-based
              : startCol + colIdx;

            // Check overwrite for range
            if (operation === 'addRange' && !options.overwrite) {
              const existingCell = targetRow.getCell(targetCol);
              if (existingCell.value !== null && existingCell.value !== undefined && existingCell.value !== '') {
                continue;
              }
            }

            // 🔧 Set value - ExcelJS keeps existing cell style automatically
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
        const newItem = {
          json: {
            ...items[i].json,
            excelOperation: operation,
            sheetModified: sheetName,
            rowsAffected: dataToInsert.length,
            sheetNames: workbook.worksheets.map(ws => ws.name),
            outputFileName: finalFileName,
          },
          binary: {},
        };

        newItem.binary[outputBinaryField] = await this.helpers.prepareBinaryData(
          Buffer.from(buffer),
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

// Parse cell reference like "A1" or "B5"
function parseCellRef(cellRef) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (match) {
    return {
      col: parseColumnIndex(match[1]) + 1, // ExcelJS 1-based
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