package org.excel.export;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookBuilder {
    Workbook workbook = new XSSFWorkbook();
    public Sheet createSheet(String name) {
        return workbook.createSheet(name);
    }
    public void save(String filename) throws IOException {
        OutputStream outputStream=new FileOutputStream(filename);
        workbook.write(outputStream);
    }
}
