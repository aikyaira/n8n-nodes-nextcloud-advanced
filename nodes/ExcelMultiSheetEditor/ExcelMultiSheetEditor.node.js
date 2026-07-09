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
        },
        {
          displayName: 'Sheet Name',
          name: 'sheetName',
          type: 'options',
          typeOptions: {
            loadOptionsMethod: 'getSheetNames',
          },
          default: '',
          description: 'Choose sheet or leave empty for first sheet. Type custom name to create new.',
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

  methods = {
    loadOptions: {
      async getSheetNames() {
        return [
          { name: '[First sheet / Create new]', value: '' },
          { name: 'Sheet1', value: 'Sheet1' },
          { name: 'Sheet2', value: 'Sheet2' },
          { name: 'Sheet3', value: 'Sheet3' },
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
        const dataFormat = this.getNodeParameter('dataFormat', i);
        const targetColumnsStr = this.getNodeParameter('targetColumns', i);
        const options = this.getNodeParameter('options', i, {});

        // Parse target columns
        let targetColumns = [];
        if (targetColumnsStr && targetColumnsStr.trim()) {
          targetColumns = targetColumnsStr.split(',').map(c => c.trim().toUpperCase());
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
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);

        if (items[i].binary && binaryPropertyName && items[i].binary[binaryPropertyName]) {
          const binaryData = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
          workbook = XLSX.read(binaryData, { type: 'buffer' });
        } else {
          workbook = XLSX.utils.book_new();
        }

        // Determine sheet name
        let sheetName = this.getNodeParameter('sheetName', i) || '';
        if (!sheetName && workbook.SheetNames.length > 0) {
          sheetName = workbook.SheetNames[0];
        } else if (!sheetName) {
          sheetName = 'Sheet1';
        }

        // Get or create worksheet
        let worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          worksheet = XLSX.utils.aoa_to_sheet([[]]);
          XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        // Get current data range
        const currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        // 🔧 PRESERVE STYLES: Save original cell objects
        const originalCells = {};
        Object.keys(worksheet).forEach(ref => {
          if (!ref.startsWith('!')) {
            originalCells[ref] = JSON.parse(JSON.stringify(worksheet[ref]));
          }
        });

        // Determine start row
        let startRow = currentData.length;
        let startCol = 0;

        if (options.appendEmptyRow && startRow > 0) {
          const lastRow = currentData[startRow - 1];
          if (lastRow.some(c => c !== '' && c !== null && c !== undefined)) {
            startRow++;
          }
        }

        if (operation === 'addRange') {
          const startCell = this.getNodeParameter('startCell', i);
          const startRef = XLSX.utils.decode_cell(startCell);
          startRow = startRef.r;
          startCol = startRef.c;
        }

        // 🔧 ADD DATA DIRECTLY to cells (preserves other cells' styles)
        for (let rowIdx = 0; rowIdx < dataToInsert.length; rowIdx++) {
          const row = dataToInsert[rowIdx];
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const col = targetColumns.length > 0 && colIdx < targetColumns.length
              ? parseColumnIndex(targetColumns[colIdx])
              : startCol + colIdx;

            const cellRef = XLSX.utils.encode_cell({ r: startRow + rowIdx, c: col });
            
            // Check overwrite
            if (operation === 'addRange' && !options.overwrite) {
              const existing = worksheet[cellRef];
              if (existing && existing.v !== undefined && existing.v !== '') continue;
            }

            // 🔧 Create new cell value
            const value = row[colIdx];
            worksheet[cellRef] = {
              v: value,
              t: typeof value === 'number' ? 'n' : 's',
            };
          }
        }

        // 🔧 RESTORE ORIGINAL CELLS that we didn't modify (preserves their styles)
        Object.keys(originalCells).forEach(ref => {
          if (!worksheet[ref]) {
            worksheet[ref] = originalCells[ref];
          }
        });

        // Update range
        const maxRow = startRow + dataToInsert.length - 1;
        const maxCol = targetColumns.length > 0
          ? Math.max(...targetColumns.map(c => parseColumnIndex(c)))
          : startCol + Math.max(...dataToInsert.map(r => r.length)) - 1;

        const oldRange = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
        
        worksheet['!ref'] = XLSX.utils.encode_range({
          s: { r: Math.min(oldRange.s.r, startRow), c: Math.min(oldRange.s.c, startCol) },
          e: { r: Math.max(oldRange.e.r, maxRow), c: Math.max(oldRange.e.c, maxCol) },
        });

        workbook.Sheets[sheetName] = worksheet;

        // Write
        const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        const outputBinaryField = options.outputBinaryField || binaryPropertyName || 'data';
        const newItem = {
          json: {
            ...items[i].json,
            excelOperation: operation,
            sheetModified: sheetName,
            rowsAffected: dataToInsert.length,
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
        const msg = error instanceof Error ? error.message : 'Unknown error';
        if (this.continueOnFail()) {
          returnData.push({ json: { error: msg, ...items[i].json } });
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
    for (let i = 0; i < col.length; i++) idx = idx * 26 + (col.charCodeAt(i) - 64);
    return idx - 1;
  }
  return parseInt(col) - 1;
}

module.exports = { ExcelMultiSheetEditor };