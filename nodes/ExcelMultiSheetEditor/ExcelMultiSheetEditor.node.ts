import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
} from 'n8n-workflow';
import * as XLSX from 'xlsx';

export class ExcelMultiSheetEditor implements INodeType {
  description: INodeTypeDescription = {
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

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const returnData: INodeExecutionData[] = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const operation = this.getNodeParameter('operation', i) as string;
        const sheetName = this.getNodeParameter('sheetName', i) as string;
        const dataFormat = this.getNodeParameter('dataFormat', i) as string;
        const options = this.getNodeParameter('options', i, {}) as {
          overwrite?: boolean;
          appendEmptyRow?: boolean;
          outputBinaryField?: string;
        };

        // Get data to insert
        let dataToInsert: any[][];
        if (dataFormat === 'defineBelow') {
          dataToInsert = this.getNodeParameter('dataToAdd', i) as any[][];
        } else {
          const jsonProperty = this.getNodeParameter('jsonDataProperty', i) as string;
          dataToInsert = this.getNodeParameter(jsonProperty, i) as any[][];
        }

        // Validate data
        if (!Array.isArray(dataToInsert)) {
          throw new Error('Data must be an array of arrays');
        }

        // Normalize data
        dataToInsert = dataToInsert.map(row => {
          if (Array.isArray(row)) return row;
          if (typeof row === 'object' && row !== null) return Object.values(row);
          return [row];
        });

        // Load or create workbook
        let workbook: XLSX.WorkBook;
        const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;

        if (items[i].binary && binaryPropertyName && items[i].binary![binaryPropertyName]) {
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
          const startCell = this.getNodeParameter('startCell', i) as string;
          addRange(worksheet, startCell, dataToInsert, options.overwrite);
        }

        // Update workbook
        workbook.Sheets[sheetName] = worksheet;

        // Write to buffer
        const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        // Prepare output
        const outputBinaryField = options.outputBinaryField || binaryPropertyName || 'data';
        const newItem: INodeExecutionData = {
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

        newItem.binary![outputBinaryField] = await this.helpers.prepareBinaryData(
          wbout,
          `edited_${sheetName}_${Date.now()}.xlsx`,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );

        returnData.push(newItem);
      } catch (error: unknown) {
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

// Helper functions outside the class
function addRowsToEnd(
  worksheet: XLSX.WorkSheet,
  dataToInsert: any[][],
  appendEmptyRow?: boolean,
): void {
  const currentData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  if (appendEmptyRow && currentData.length > 0) {
    const lastRow = currentData[currentData.length - 1];
    const hasContent = lastRow.some(
      (cell: any) => cell !== '' && cell !== null && cell !== undefined,
    );
    if (hasContent) {
      currentData.push([]);
    }
  }

  const newData = [...currentData, ...dataToInsert];
  const newWorksheet = XLSX.utils.aoa_to_sheet(newData);

  Object.keys(worksheet).forEach(key => delete worksheet[key]);
  Object.assign(worksheet, newWorksheet);
}

function addRange(
  worksheet: XLSX.WorkSheet,
  startCell: string,
  dataToInsert: any[][],
  overwrite?: boolean,
): void {
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