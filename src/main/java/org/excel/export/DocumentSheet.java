package org.excel.export;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.export.model.Document;
import org.excel.export.model.Element;

public class DocumentSheet extends ElementSheet {

    public DocumentSheet(XSSFWorkbook workbook, String name) {
        super(workbook, name);
    }

    public void create(List<Element> elements, List<Document> documents) {
        addDocuments(documents);
        addElements(elements, documents);
    }

    private void addDocuments(List<Document> documents) {
        Row row = sheet.createRow(rowIndex++);
        row.setRowStyle(headerStyle);
        int cellIndex = 0;
        setHeaderCell(row, cellIndex++, "Code", 2000);
        setHeaderCell(row, cellIndex++, "Name", 8000);
        for (Document document : documents) {
            setHeaderCell(row, cellIndex++, document.getCode(), 1500);
        }
    }

    private void addElements(List<Element> elements, List<Document> documents) {
        for (Element element : elements) {
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            setTextCell(row, cellIndex++, element.getCodeName());
            setTextCell(row, cellIndex++, element.getName().getFi());
            for (Document document : documents) {
                Element foundElement = findElement(document, element.getCode());
                if (foundElement == null) {
                    setTextCell(row, cellIndex++, null);
                } else {
                    if (foundElement.isVisible()) {
                        setTextCell(row, cellIndex++, "[+]");
                    } else {
                        setTextCell(row, cellIndex++, "[*]");
                    }
                }
            }
        }
    }

    private Element findElement(Document document, String code) {
        for (Element element : document.getElements()) {
            if (element.getCode().equals(code)) {
                return element;
            }
        }
        return null;
    }

}
