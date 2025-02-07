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

    // Define styles for the header row
    const headerStyle = {
      font: { bold: true }, // Bold text
      fill: { fgColor: { rgb: 'D3D3D3' } }, // Background color: #D3D3D3 (Light Gray)
      alignment: { horizontal: 'center', vertical: 'center' }, // Center text
    };

    // Define data for the Excel file
    const data = employees.length
      ? employees.map((emp) => ({
          'First Name': emp.firstName || '',
          'Last Name': emp.lastName || '',
          City: emp.city || '',
          Date: emp.date || '',
        }))
      : [headers.reduce((acc, header) => ({ ...acc, [header]: '' }), {})];

    // Create a worksheet
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    // Ensure headers are always present
    XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });

    // Apply styling to header row
    headers.forEach((_, colIndex) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      if (!worksheet[cellRef]) worksheet[cellRef] = { v: headers[colIndex] }; // Ensure cell exists
      worksheet[cellRef].s = headerStyle; // Apply header styles
    });

    // Set row height for header (35px)
    worksheet['!rows'] = [{ hpx: 35 }];

    // Set column widths
    worksheet['!cols'] = [
      { wch: 20 }, // First Name - Auto width
      { wch: 57 }, // Last Name - Fixed width (57)
      { wch: 20 }, // City - Auto width
      { wch: 20 }, // Date - Auto width
    ];

    // Apply auto filter
    if (worksheet['!ref']) {
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
    }

    // Create a workbook and append the worksheet
    const workbook: XLSX.WorkBook = {
      Sheets: { Employees: worksheet },
      SheetNames: ['Employees'],
    };

    // Convert workbook to a binary Excel file
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    // Generate and download the Excel file
    const fileName = this.getDynamicFileName();
    this.saveAsExcelFile(excelBuffer, fileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: this.EXCEL_TYPE });
    saveAs(data, fileName);
  }

  private getDynamicFileName(): string {
    const now = new Date();
    const mmddyyyy = `${(now.getMonth() + 1).toString().padStart(2, '0')}${now
      .getDate()
      .toString()
      .padStart(2, '0')}${now.getFullYear()}`;
    return `EmployeeData_${mmddyyyy}.xlsx`;
  }
}