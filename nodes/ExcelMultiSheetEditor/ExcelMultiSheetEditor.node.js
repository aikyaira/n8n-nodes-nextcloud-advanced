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
          description: 'Optional: Specify columns to write to. Leave empty to start from column A.',
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
      async getSheetNames() {
        const items = this.getInputData();
        
        if (!items || items.length === 0) {
          return [{ name: 'Sheet1', value: 'Sheet1' }];
        }

        try {
          const binaryPropertyName = this.getNodeParameter('binaryPropertyName', 0);
          
          if (items[0].binary && items[0].binary[binaryPropertyName]) {
            const binaryData = await this.helpers.getBinaryDataBuffer(0, binaryPropertyName);
            const workbook = XLSX.read(binaryData, { type: 'buffer' });
            
            if (workbook.SheetNames && workbook.SheetNames.length > 0) {
              return workbook.SheetNames.map(name => ({
                name: name,
                value: name,
              }));
            }
          }
        } catch (error) {
          // Silently fall back to defaults
        }

        return [
          { name: 'Sheet1', value: 'Sheet1' },
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
        const sheetName = this.getNodeParameter('sheetName', i) || 'Sheet1';
        const dataFormat = this.getNodeParameter('dataFormat', i);
        const targetColumnsStr = this.getNodeParameter('targetColumns', i);
        const options = this.getNodeParameter('options', i, {});

        // Parse target columns
        let targetColumns = [];
        if (targetColumnsStr && targetColumnsStr.trim()) {
          targetColumns = targetColumnsStr.split(',').map(col => col.trim().toUpperCase());
        }

        // Get data to insert
        let dataToInsert;
        if (dataFormat === 'defineBelow') {
          dataToInsert = this.getNodeParameter('dataToAdd', i);
        } else {
          const jsonProperty = this.getNodeParameter('jsonDataProperty', i);
          dataToInsert = items[i].json[jsonProperty];
        }

        // Parse JSON string if needed
        if (typeof dataToInsert === 'string') {
          try {
            dataToInsert = JSON.parse(dataToInsert);
          } catch (parseError) {
            throw new Error(`Failed to parse data as JSON. Error: ${parseError.message}`);
          }
        }

        // Validate
        if (!Array.isArray(dataToInsert)) {
          throw new Error(`Data must be an array. Received: ${typeof dataToInsert}`);
        }

        // Normalize
        dataToInsert = dataToInsert.map((row, index) => {
          if (Array.isArray(row)) return row;
          if (typeof row === 'object' && row !== null) return Object.values(row);
          return [row];
        });

        // Load or create workbook - 🔧 WITHOUT cellStyles
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

        // 🔧 Save original cell properties (styles, formatting)
        const originalCells = {};
        Object.keys(worksheet).forEach(cellRef => {
          if (cellRef.startsWith('!')) return;
          originalCells[cellRef] = { ...worksheet[cellRef] };
        });

        // Perform operation
        if (operation === 'addRows') {
          addRowsToEnd(worksheet, dataToInsert, options.appendEmptyRow, targetColumns);
        } else if (operation === 'addRange') {
          const startCell = this.getNodeParameter('startCell', i);
          addRange(worksheet, startCell, dataToInsert, options.overwrite, targetColumns);
        }

        // 🔧 Restore original cells that weren't modified
        Object.keys(originalCells).forEach(cellRef => {
          if (!worksheet[cellRef]) {
            worksheet[cellRef] = originalCells[cellRef];
          } else if (worksheet[cellRef].v === undefined || worksheet[cellRef].v === '') {
            worksheet[cellRef] = originalCells[cellRef];
          } else {
            // Keep new value but restore original formatting if any
            if (originalCells[cellRef].s && !worksheet[cellRef].s) {
              worksheet[cellRef].s = originalCells[cellRef].s;
            }
            if (originalCells[cellRef].z && !worksheet[cellRef].z) {
              worksheet[cellRef].z = originalCells[cellRef].z;
            }
          }
        });

        workbook.Sheets[sheetName] = worksheet;

        // Write WITHOUT cellStyles
        const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

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

// Helper functions

function parseColumnIndex(col) {
  if (typeof col === 'string' && /^[A-Z]+$/.test(col)) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index - 1;
  }
  return parseInt(col) - 1;
}

function addRowsToEnd(worksheet, dataToInsert, appendEmptyRow, targetColumns) {
  const currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  // Add empty row if needed
  if (appendEmptyRow && currentData.length > 0) {
    const lastRow = currentData[currentData.length - 1];
    const hasContent = lastRow.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (hasContent) currentData.push([]);
  }

  const startRow = currentData.length;

  // Add new rows
  dataToInsert.forEach((row, rowIndex) => {
    const targetRow = startRow + rowIndex;
    if (!currentData[targetRow]) currentData[targetRow] = [];

    if (targetColumns && targetColumns.length > 0) {
      targetColumns.forEach((col, colIndex) => {
        const colNum = parseColumnIndex(col);
        if (colIndex < row.length) {
          currentData[targetRow][colNum] = row[colIndex];
        }
      });
    } else {
      currentData[targetRow] = [...currentData[targetRow], ...row];
    }
  });

  // Rebuild sheet
  const newWorksheet = XLSX.utils.aoa_to_sheet(currentData);
  const newRange = newWorksheet['!ref'];
  
  Object.keys(worksheet).forEach(key => delete worksheet[key]);
  Object.assign(worksheet, newWorksheet);
  worksheet['!ref'] = newRange;
}

function addRange(worksheet, startCell, dataToInsert, overwrite, targetColumns) {
  const startRef = XLSX.utils.decode_cell(startCell);
  let maxCol = startRef.c;
  let maxRow = startRef.r;

  for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
    const row = dataToInsert[rowIdx];
    
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const targetCol = targetColumns && colIdx < targetColumns.length
        ? parseColumnIndex(targetColumns[colIdx])
        : startRef.c + colIdx;

      const cellRef = XLSX.utils.encode_cell({ r: startRef.r + rowIdx, c: targetCol });
      const existing = worksheet[cellRef];

      if (!overwrite && existing && existing.v !== undefined && existing.v !== '') continue;

      worksheet[cellRef] = {
        v: row[colIdx],
        t: typeof row[colIdx] === 'number' ? 'n' : 's',
        ...(existing && existing.s && { s: existing.s }),
      };

      if (targetCol > maxCol) maxCol = targetCol;
      if (startRef.r + rowIdx > maxRow) maxRow = startRef.r + rowIdx;
    }
  }

  const currentRef = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  
  worksheet['!ref'] = XLSX.utils.encode_range({
    s: { r: Math.min(currentRef.s.r, startRef.r), c: Math.min(currentRef.s.c, startRef.c) },
    e: { r: Math.max(currentRef.e.r, maxRow), c: Math.max(currentRef.e.c, maxCol) },
  });
}

module.exports = { ExcelMultiSheetEditor };