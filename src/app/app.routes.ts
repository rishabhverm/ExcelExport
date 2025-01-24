import { Routes } from '@angular/router';

export const routes: Routes = [
    {
        path: 'excel/file',
        data:{title:'Excel'},
        loadComponent: () => import('./excelfile/excelfile.component').then(m => m.ExcelfileComponent),
        
    },

];
