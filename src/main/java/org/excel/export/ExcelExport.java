package org.excel.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.export.model.ExportData;

public class ExcelExport {
    Workbook workbook = new XSSFWorkbook();

    public void createReport(String filename, ExportData exportData) throws FileNotFoundException, IOException {
        ElementSheet elementSheet = new ElementSheet(workbook, "Elements");
        elementSheet.create(exportData.getElements());
        DocumentSheet documentSheet = new DocumentSheet(workbook, "Documents");
        documentSheet.create(exportData.getElements(), exportData.getDocuments());
        workbook.write(new FileOutputStream(filename));
        workbook.close();
    }
}
