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
            {
              name: 'Add Rows to End',
              value: 'addRows',
              description: 'Add rows to the end of a sheet',
            },
            {
              name: 'Add Range at Position',
              value: 'addRange',
              description: 'Add data in a specific range/cell position',
            },
          ],
          default: 'addRows',
          description: 'The operation to perform',
        },
        {
          displayName: 'Input Binary Field',
          name: 'binaryPropertyName',
          type: 'string',
          default: 'data',
          description: 'Name of the binary property containing the Excel file. Leave empty to create new file.',
        },
        {
          displayName: 'Sheet Name',
          name: 'sheetName',
          type: 'string',
          default: 'Sheet1',
          description: 'Name of the sheet to edit. Will be created if it does not exist.',
        },
        {
          displayName: 'Start Cell',
          name: 'startCell',
          type: 'string',
          default: 'A1',
          displayOptions: {
            show: {
              operation: ['addRange'],
            },
          },
          description: 'Starting cell for range insertion (e.g., A1, B5, C10)',
        },
        {
          displayName: 'Data Format',
          name: 'dataFormat',
          type: 'options',
          options: [
            {
              name: 'Define Below',
              value: 'defineBelow',
              description: 'Define data in the node parameters',
            },
            {
              name: 'From JSON Property',
              value: 'fromJson',
              description: 'Take data from an incoming JSON property',
            },
          ],
          default: 'defineBelow',
          description: 'Where the data to insert comes from',
        },
        {
          displayName: 'JSON Data Property',
          name: 'jsonDataProperty',
          type: 'string',
          default: 'excelData',
          displayOptions: {
            show: {
              dataFormat: ['fromJson'],
            },
          },
          description: 'Name of the property which contains the data to insert',
        },
        {
          displayName: 'Data to Add',
          name: 'dataToAdd',
          type: 'json',
          default: '[\n  ["Name", "Age", "City"],\n  ["John", 30, "New York"],\n  ["Jane", 25, "London"]\n]',
          displayOptions: {
            show: {
              dataFormat: ['defineBelow'],
            },
          },
          description: 'Array of arrays representing the data to add. Each inner array is a row.',
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
              displayOptions: {
                show: {
                  '/operation': ['addRange'],
                },
              },
              description: 'Whether to overwrite existing data in the range',
            },
            {
              displayName: 'Add Empty Row Before',
              name: 'appendEmptyRow',
              type: 'boolean',
              default: false,
              displayOptions: {
                show: {
                  '/operation': ['addRows'],
                },
              },
              description: 'Add an empty row before new data',
            },
            {
              displayName: 'Output Binary Field',
              name: 'outputBinaryField',
              type: 'string',
              default: 'data',
              description: 'Name of the binary field to output',
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
        const sheetName = this.getNodeParameter('sheetName', i);
        const dataFormat = this.getNodeParameter('dataFormat', i);
        const options = this.getNodeParameter('options', i, {});

        // Get data to insert
        let dataToInsert;
        if (dataFormat === 'defineBelow') {
          dataToInsert = this.getNodeParameter('dataToAdd', i);
        } else {
          const jsonProperty = this.getNodeParameter('jsonDataProperty', i);
          dataToInsert = this.getNodeParameter(jsonProperty, i);
        }

        // 🔧 FIX: Parse JSON string if needed (n8n passes JSON as string)
        if (typeof dataToInsert === 'string') {
          try {
            dataToInsert = JSON.parse(dataToInsert);
          } catch (parseError) {
            throw new Error(`Failed to parse data as JSON. Make sure it's valid JSON array. Error: ${parseError.message}`);
          }
        }

        // 🔧 FIX: Better validation with detailed error
        if (!Array.isArray(dataToInsert)) {
          throw new Error(`Data must be an array of arrays. Received type: ${typeof dataToInsert}. Value: ${JSON.stringify(dataToInsert).substring(0, 100)}`);
        }

        // Normalize data - convert objects to arrays if needed
        dataToInsert = dataToInsert.map((row, index) => {
          if (Array.isArray(row)) return row;
          if (typeof row === 'object' && row !== null) return Object.values(row);
          throw new Error(`Row ${index} must be an array or object. Received: ${typeof row}`);
        });

        // Load or create workbook
        let workbook;
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          workbook = XLSX.read(binaryData, { type: 'buffer' });
        } else {
          workbook = XLSX.utils.book_new();
        }

        // Get or create worksheet
        let worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          worksheet = XLSX.utils.aoa_to_sheet([[]]);
          XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        // Perform operation
        if (operation === 'addRows') {
          addRowsToEnd(worksheet, dataToInsert, options.appendEmptyRow);
        } else if (operation === 'addRange') {
          const startCell = this.getNodeParameter('startCell', i);
          addRange(worksheet, startCell, dataToInsert, options.overwrite);
        }

        // Update workbook
        workbook.Sheets[sheetName] = worksheet;

        // Write to buffer
        const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        // Prepare output
        const outputBinaryField = options.outputBinaryField || binaryPropertyName || 'data';
        const newItem = {
          json: {
            ...items[i].json,
            excelOperation: operation,
            sheetModified: sheetName,
            rowsAffected: dataToInsert.length,
            totalSheets: workbook.SheetNames.length,
            sheetNames: workbook.SheetNames,
          },
          binary: {},
        };

        newItem.binary[outputBinaryField] = await this.helpers.prepareBinaryData(
          wbout,
          `edited_${sheetName}_${Date.now()}.xlsx`,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );

        returnData.push(newItem);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        if (this.continueOnFail()) {
          returnData.push({
            json: {
              error: errorMessage,
              ...items[i].json,
            },
          });
          continue;
        }
        throw error;
      }
    }

    return [returnData];
  }
}

// Helper functions
function addRowsToEnd(worksheet, dataToInsert, appendEmptyRow) {
  const currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  if (appendEmptyRow && currentData.length > 0) {
    const lastRow = currentData[currentData.length - 1];
    const hasContent = lastRow.some(
      (cell) => cell !== '' && cell !== null && cell !== undefined,
    );
    if (hasContent) {
      currentData.push([]);
    }
  }

  const newData = [...currentData, ...dataToInsert];
  const newWorksheet = XLSX.utils.aoa_to_sheet(newData);

  // Clear old worksheet and copy new data
  Object.keys(worksheet).forEach(key => delete worksheet[key]);
  Object.assign(worksheet, newWorksheet);
}

function addRange(worksheet, startCell, dataToInsert, overwrite) {
  const startCellRef = XLSX.utils.decode_cell(startCell);

  const currentRange = worksheet['!ref']
    ? XLSX.utils.decode_range(worksheet['!ref'])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

  let maxCol = 0;
  for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
    const row = dataToInsert[rowIdx];
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const cellRef = XLSX.utils.encode_cell({
        r: startCellRef.r + rowIdx,
        c: startCellRef.c + colIdx,
      });

      const existingCell = worksheet[cellRef];

      if (!overwrite && existingCell && existingCell.v !== undefined && existingCell.v !== '') {
        continue;
      }

      worksheet[cellRef] = {
        v: row[colIdx],
        t: typeof row[colIdx] === 'number' ? 'n' : 's',
      };

      maxCol = Math.max(maxCol, startCellRef.c + colIdx);
    }
  }

  const newRange = {
    s: {
      r: Math.min(currentRange.s.r, startCellRef.r),
      c: Math.min(currentRange.s.c, startCellRef.c),
    },
    e: {
      r: Math.max(currentRange.e.r, startCellRef.r + dataToInsert.length - 1),
      c: Math.max(currentRange.e.c, maxCol),
    },
  };

  worksheet['!ref'] = XLSX.utils.encode_range(newRange);
}

module.exports = { ExcelMultiSheetEditor };