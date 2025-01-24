import { Component, Inject } from '@angular/core';
import { ExcelService } from './excel-export-service.service';

@Component({
  selector: 'app-excelfile',
  standalone: true,
  imports: [],
  templateUrl: './excelfile.component.html',
  styleUrl: './excelfile.component.scss'
})
export class ExcelfileComponent {
  constructor(@Inject(ExcelService) private excelExportService: ExcelService) {}

   downloadExcel() {
    // Example data for testing
    const employees = [
      { firstName: 'John', lastName: 'Doe', city: 'New York', date: '2025-01-24' },
      { firstName: 'Jane', lastName: 'Smith', city: 'Los Angeles', date: '2025-01-23' },
    ];

    // Check if there is data
    if (employees.length > 0) {
      this.excelExportService.downloadExcel(employees);
    } else {
      // Pass an empty array to generate only headers
      this.excelExportService.downloadExcel([]);
    }
  }
}