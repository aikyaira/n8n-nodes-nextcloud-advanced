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
          type: 'options',
          typeOptions: {
            loadOptionsMethod: 'getSheetNames',
          },
          default: '',
          description: 'Choose an existing sheet or type a new name to create one',
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
          displayName: 'Target Columns',
          name: 'targetColumns',
          type: 'string',
          default: '',
          placeholder: 'e.g. A,B,C or 1,2,3',
          description: 'Optional: Specify columns to write to (e.g., "A,B,C" or "1,2,3"). Leave empty to write starting from the first available column.',
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

  methods = {
    loadOptions: {
      // Get sheet names from input file
      async getSheetNames() {
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', 0);
        const inputData = this.getInputData();

        if (inputData.length === 0 || !inputData[0].binary || !inputData[0].binary[binaryPropertyName]) {
          return [
            { name: 'Sheet1', value: 'Sheet1' },
          ];
        }

        try {
          const binaryData = await this.helpers.getBinaryDataBuffer(0, binaryPropertyName);
          const workbook = XLSX.read(binaryData, { type: 'buffer' });
          
          return workbook.SheetNames.map(name => ({
            name: name,
            value: name,
          }));
        } catch (error) {
          return [
            { name: 'Sheet1', value: 'Sheet1' },
          ];
        }
      },
    },
  };

  async execute() {
    const items = this.getInputData();
    const returnData = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const operation = this.getNodeParameter('operation', i);
        const sheetName = this.getNodeParameter('sheetName', i) || 'Sheet1';
        const dataFormat = this.getNodeParameter('dataFormat', i);
        const targetColumnsStr = this.getNodeParameter('targetColumns', i);
        const options = this.getNodeParameter('options', i, {});

        // Parse target columns
        let targetColumns = [];
        if (targetColumnsStr && targetColumnsStr.trim()) {
          targetColumns = targetColumnsStr.split(',').map(col => col.trim());
        }

        // Get data to insert
        let dataToInsert;
        if (dataFormat === 'defineBelow') {
          dataToInsert = this.getNodeParameter('dataToAdd', i);
        } else {
          const jsonProperty = this.getNodeParameter('jsonDataProperty', i);
          dataToInsert = this.getNodeParameter(jsonProperty, i);
        }

        // Parse JSON string if needed
        if (typeof dataToInsert === 'string') {
          try {
            dataToInsert = JSON.parse(dataToInsert);
          } catch (parseError) {
            throw new Error(`Failed to parse data as JSON. Make sure it's valid JSON array. Error: ${parseError.message}`);
          }
        }

        // Validate data
        if (!Array.isArray(dataToInsert)) {
          throw new Error(`Data must be an array of arrays. Received type: ${typeof dataToInsert}`);
        }

        // Normalize data
        dataToInsert = dataToInsert.map((row, index) => {
          if (Array.isArray(row)) return row;
          if (typeof row === 'object' && row !== null) return Object.values(row);
          throw new Error(`Row ${index} must be an array or object. Received: ${typeof row}`);
        });

        // Load or create workbook with cellStyles enabled
        let workbook;
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          // 🔧 FIX: Read with cellStyles to preserve formatting
          workbook = XLSX.read(binaryData, { 
            type: 'buffer',
            cellStyles: true,
            cellFormula: true,
            cellDates: true,
          });
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
          addRowsToEnd(worksheet, dataToInsert, options.appendEmptyRow, targetColumns);
        } else if (operation === 'addRange') {
          const startCell = this.getNodeParameter('startCell', i);
          addRange(worksheet, startCell, dataToInsert, options.overwrite, targetColumns);
        }

        // Update workbook
        workbook.Sheets[sheetName] = worksheet;

        // 🔧 FIX: Write with cellStyles to preserve formatting
        const wbout = XLSX.write(workbook, { 
          type: 'buffer', 
          bookType: 'xlsx',
          cellStyles: true,
        });

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
            targetColumns: targetColumns.length > 0 ? targetColumns : 'auto',
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

function parseColumnIndex(col) {
  // If it's a letter (A, B, C, ...), convert to 0-based index
  if (typeof col === 'string' && /^[A-Za-z]+$/.test(col)) {
    col = col.toUpperCase();
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index - 1; // 0-based
  }
  // If it's a number (1, 2, 3, ...), convert to 0-based
  return parseInt(col) - 1;
}

function addRowsToEnd(worksheet, dataToInsert, appendEmptyRow, targetColumns) {
  // Convert to array while preserving existing styles
  const currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  
  // Store original cell styles
  const originalStyles = {};
  Object.keys(worksheet).forEach(cellRef => {
    if (cellRef.startsWith('!')) return;
    if (worksheet[cellRef].s) {
      originalStyles[cellRef] = worksheet[cellRef].s;
    }
  });

  if (appendEmptyRow && currentData.length > 0) {
    const lastRow = currentData[currentData.length - 1];
    const hasContent = lastRow.some(
      (cell) => cell !== '' && cell !== null && cell !== undefined,
    );
    if (hasContent) {
      currentData.push([]);
    }
  }

  // If target columns specified, map data to those columns
  if (targetColumns && targetColumns.length > 0) {
    const startRow = currentData.length;
    targetColumns.forEach((col, colIndex) => {
      const colNum = parseColumnIndex(col);
      dataToInsert.forEach((row, rowIndex) => {
        if (colIndex < row.length) {
          currentData[startRow + rowIndex] = currentData[startRow + rowIndex] || [];
          currentData[startRow + rowIndex][colNum] = row[colIndex];
        }
      });
    });
  } else {
    // Just append all data
    const newData = [...currentData, ...dataToInsert];
    Object.keys(worksheet).forEach(key => delete worksheet[key]);
    const newWorksheet = XLSX.utils.aoa_to_sheet(newData);
    Object.assign(worksheet, newWorksheet);
    
    // Restore original styles
    Object.keys(originalStyles).forEach(cellRef => {
      if (worksheet[cellRef] && !worksheet[cellRef].s) {
        worksheet[cellRef].s = originalStyles[cellRef];
      }
    });
    return;
  }

  // Rebuild worksheet with updated data
  const newWorksheet = XLSX.utils.aoa_to_sheet(currentData);
  Object.keys(worksheet).forEach(key => delete worksheet[key]);
  Object.assign(worksheet, newWorksheet);

  // Restore original styles
  Object.keys(originalStyles).forEach(cellRef => {
    if (worksheet[cellRef] && !worksheet[cellRef].s) {
      worksheet[cellRef].s = originalStyles[cellRef];
    }
  });
}

function addRange(worksheet, startCell, dataToInsert, overwrite, targetColumns) {
  const startCellRef = XLSX.utils.decode_cell(startCell);

  const currentRange = worksheet['!ref']
    ? XLSX.utils.decode_range(worksheet['!ref'])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

  let maxCol = 0;
  let maxRow = 0;

  for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
    const row = dataToInsert[rowIdx];
    const colsToUse = targetColumns && targetColumns.length > 0 ? targetColumns : row.map((_, idx) => idx);
    
    for (let colIdx = 0; colIdx < (targetColumns ? targetColumns.length : row.length); colIdx++) {
      let targetCol;
      if (targetColumns && targetColumns.length > 0) {
        targetCol = parseColumnIndex(colsToUse[colIdx]);
      } else {
        targetCol = startCellRef.c + colIdx;
      }

      const cellRef = XLSX.utils.encode_cell({
        r: startCellRef.r + rowIdx,
        c: targetCol,
      });

      const existingCell = worksheet[cellRef];

      // Skip if not overwriting and cell has content
      if (!overwrite && existingCell && existingCell.v !== undefined && existingCell.v !== '') {
        continue;
      }

      // Preserve existing style if any
      const existingStyle = existingCell && existingCell.s ? existingCell.s : {};

      worksheet[cellRef] = {
        v: row[colIdx],
        t: typeof row[colIdx] === 'number' ? 'n' : 's',
        s: existingStyle, // 🔧 Preserve existing cell style
      };

      maxCol = Math.max(maxCol, targetCol);
      maxRow = Math.max(maxRow, startCellRef.r + rowIdx);
    }
  }

  const newRange = {
    s: {
      r: Math.min(currentRange.s.r, startCellRef.r),
      c: Math.min(currentRange.s.c, startCellRef.c),
    },
    e: {
      r: Math.max(currentRange.e.r, maxRow),
      c: Math.max(currentRange.e.c, maxCol),
    },
  };

  worksheet['!ref'] = XLSX.utils.encode_range(newRange);
}

module.exports = { ExcelMultiSheetEditor };