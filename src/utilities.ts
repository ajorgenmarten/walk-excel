import ExcelJs from 'exceljs';
import * as Xls from 'xlsx';

export class ExcelModule {
  async walkRows(path: string, rowHandler: RowHandler) {
    const extension = path.split('.').pop()?.toLowerCase();
    if (extension == 'xlsx') return this.workWithXlsx(path, rowHandler);
    if (extension == 'xls') return this.workWithXls(path, rowHandler);
    throw new Error(
      `No se puede manejar archivos con la extension ${extension}`
    );
  }
  private getHeaders(sheet: ExcelJs.Worksheet) {
    if (sheet.rowCount < 1) return [];
    const headerRow = sheet.getRow(1);
    const headers: HeaderCol[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.value != null && cell.value != undefined) {
        headers.push({
          header: cell.value.toString(),
          columnNumber: cell.col.toString(),
          columnLetter: cell.address.replace(/[0-9]/g, ''),
        });
      }
    });
    return headers;
  }

  private async workWithXlsx(path: string, rowHandler: RowHandler) {
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile(path);
    for (const sheet of workbook.worksheets) {
      const headers = this.getHeaders(sheet);
      for (let index = 2; index < sheet.rowCount; index++) {
        const row = sheet.getRow(index);
        const record: Record<string, any> = {};
        if (Array.isArray(row.values)) {
          row.values.forEach((value: any, index: number) => {
            const key = headers.find(
              (value) => +value.columnNumber == index
            ) as HeaderCol;
            record[key.header] = value;
          });
        }
        await rowHandler(record, index - 2);
      }
    }
  }

  private async workWithXls(path: string, rowHandler: RowHandler) {
    const workbook = Xls.readFile(path, { dense: true });

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const jsonData = Xls.utils.sheet_to_json(sheet);
      for (let index = 0; index < jsonData.length; index++)
        await rowHandler(jsonData[index] as Record<string, any>, index);
    }
  }
}

export type RowHandler = (
  rowModel: Record<string, any>,
  rowIndex?: number
) => void | Promise<void>;
export type HeaderCol = {
  header: string;
  columnNumber: string;
  columnLetter: string;
};
