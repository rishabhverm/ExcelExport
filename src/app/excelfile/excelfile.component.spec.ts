import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelfileComponent } from './excelfile.component';

describe('ExcelfileComponent', () => {
  let component: ExcelfileComponent;
  let fixture: ComponentFixture<ExcelfileComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExcelfileComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ExcelfileComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
