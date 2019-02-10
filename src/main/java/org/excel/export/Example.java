package org.excel.export;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Example {
    public void test() {
    Workbook workbook = new XSSFWorkbook();
    
    Sheet sheet = workbook.createSheet("Persons");
    sheet.setColumnWidth(0, 6000); 
    sheet.setColumnWidth(1, 4000);
     
    Row header = sheet.createRow(0);
     
    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
     
    XSSFFont font = ((XSSFWorkbook) workbook).createFont();
    font.setFontName("Arial");
    font.setFontHeightInPoints((short) 16);
    font.setBold(true);
    headerStyle.setFont(font);
     
    Cell headerCell = header.createCell(0);
    headerCell.setCellValue("Name");
    headerCell.setCellStyle(headerStyle);
     
    headerCell = header.createCell(1);
    headerCell.setCellValue("Age");
    headerCell.setCellStyle(headerStyle);
    }
}
