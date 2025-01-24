import * as XLSX from 'xlsx';
import { Injectable } from '@angular/core';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  downloadExcel(employees: any[] = []) {
    const headers = ['First Name', 'Last Name', 'City', 'Date'];

    // Prepare data for the Excel file
    const data = employees.length
      ? employees.map((emp) => ({
          'First Name': emp.firstName || '',
          'Last Name': emp.lastName || '',
          City: emp.city || '',
          Date: emp.date || '',
        }))
      : [headers.reduce((acc, header) => ({ ...acc, [header]: '' }), {})];

    // Create a worksheet with auto filter and column fitting
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    // Add headers manually if no data
    if (!employees.length) {
      XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });
    }

    // Apply auto filter
    const range = XLSX.utils.decode_range(worksheet['!ref']!);
    worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };

    // Auto fit columns
    worksheet['!cols'] = headers.map(() => ({ wpx: 100 }));

    // Create workbook
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data'],
    };

    // Write workbook to file
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    // Save file with dynamic filename
    const fileName = this.getDynamicFileName();
    this.saveAsExcelFile(excelBuffer, fileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: EXCEL_TYPE });
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

const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';