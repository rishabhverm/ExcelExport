import * as XLSX from 'xlsx-js-style'; // Use xlsx-js-style for styling
import { Injectable } from '@angular/core';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  private readonly EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  downloadExcel(employees: any[] = []) {
    const headers = ['First Name', 'Last Name', 'City', 'Date'];

    // Define styles
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: 'D3D3D3' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: { bottom: { style: 'thin', color: { rgb: '000000' } } },
    };

    const centerAlignStyle = {
      alignment: { horizontal: 'center', vertical: 'center' },
    };

    // Format data with correct alignment
    const data = employees.length
      ? employees.map((emp) => ({
          'First Name': { v: emp.firstName || '', s: centerAlignStyle },
          'Last Name': { v: emp.lastName || '' },
          City: { v: emp.city || '', s: centerAlignStyle },
          Date: { v: emp.date || '', s: centerAlignStyle },
        }))
      : [headers.reduce((acc, header) => ({ ...acc, [header]: { v: '' } }), {})];

    // Create worksheet
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    // Add headers with styles
    XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });
    headers.forEach((_, colIndex) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      if (!worksheet[cellRef]) worksheet[cellRef] = { v: headers[colIndex] };
      worksheet[cellRef].s = headerStyle;
    });

    // Auto-format cells with borders for better readability
    const totalRows = employees.length || 1;
    for (let rowIndex = 1; rowIndex <= totalRows; rowIndex++) {
      headers.forEach((_, colIndex) => {
        const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        if (!worksheet[cellRef]) worksheet[cellRef] = { v: '' };
        worksheet[cellRef].s = {
          alignment: { horizontal: colIndex !== 1 ? 'center' : 'left' },
          border: {
            bottom: rowIndex === totalRows ? { style: 'thin', color: { rgb: '000000' } } : undefined,
            right: colIndex === headers.length - 1 ? { style: 'thin', color: { rgb: '000000' } } : undefined,
          },
        };
      });
    }

    // Set row height & column widths
    worksheet['!rows'] = [{ hpx: 35 }];
    worksheet['!cols'] = [
      { wch: 20 },
      { wch: 57 },
      { wch: 20 },
      { wch: 20 },
    ];

    // Apply auto filter
    if (worksheet['!ref']) {
      worksheet['!autofilter'] = { ref: worksheet['!ref'] };
    }

    // Create workbook and save file
    const workbook: XLSX.WorkBook = { Sheets: { Employees: worksheet }, SheetNames: ['Employees'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    this.saveAsExcelFile(excelBuffer, this.getDynamicFileName());
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: this.EXCEL_TYPE });
    saveAs(data, fileName);
  }

  private getDynamicFileName(): string {
    const now = new Date();
    return `EmployeeData_${now.getMonth() + 1}${now.getDate()}${now.getFullYear()}.xlsx`;
  }
}
