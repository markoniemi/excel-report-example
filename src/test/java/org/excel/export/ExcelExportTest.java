package org.excel.export;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.excel.export.model.Document;
import org.excel.export.model.Element;
import org.excel.export.model.ExportData;
import org.excel.export.model.Field;
import org.excel.export.model.LocalizedText;
import org.junit.Assert;
import org.junit.Test;

public class ExcelExportTest {
    String[] codes = { "CGU", "DPO", "EIR", "IPO", "OPO", "CW1", "CW2", "TST", "EIA", "CWT" };

    @Test
    public void createReport() throws FileNotFoundException, IOException, InvalidFormatException {
        ExportData exportData = createExportData();
        ExcelExport excelExport = new ExcelExport();
        String filename = "target/export.xlsx";
        excelExport.createReport(filename, exportData);
        assertFile(filename);
    }

    private void assertFile(String filename) throws IOException, InvalidFormatException, FileNotFoundException {
        Workbook workbook = WorkbookFactory.create(new FileInputStream(filename));
        Sheet elementSheet = workbook.getSheet("Elements");
        Assert.assertEquals(41, elementSheet.getPhysicalNumberOfRows());
        Row row = elementSheet.getRow(0);
        Assert.assertEquals(20, row.getPhysicalNumberOfCells());
        workbook.close();
    }

    private ExportData createExportData() {
        ExportData exportData = new ExportData();
        List<Element> elements = new ArrayList<>();
        List<Document> documents = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            elements.add(createElement("code", i));
            documents.add(createDocument(codes[i], i, elements));
        }
        exportData.setElements(elements);
        exportData.setDocuments(documents);
        return exportData;
    }

    private Document createDocument(String code, int index, List<Element> elements) {
        Document document = new Document();
        document.setCode(code);
        document.setName(new LocalizedText("name_fi", "name_sv", "name_en"));
        document.setElements(elements);
        return document;
    }

    private Element createElement(String code, int index) {
        Element element = new Element();
        element.setCode("ElementCode" + index);
        element.setCodeName((index + 1) + "/10");
        element.setGroup("group" + index);
        element.setName(new LocalizedText("Element" + index, "name_sv", "name_en"));
        element.setVisible(index % 2 == 1);
        element.setRepeat(1);
        element.setRepeatAsGroup(true);
        element.setGroupPrintId("groupPrintId" + index);
        element.setPrintId("printId" + index);
        element.setXpath("xpath" + index);
        element.setStartDate(new Date());
        element.setFields(createFields());
        return element;
    }

    private List<Field> createFields() {
        List<Field> fields = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            Field field = new Field();
            field.setCode("field" + i);
            field.setName(new LocalizedText("field" + i, "name_sv", "name_en"));
            field.setType("type");
            field.setCodeset("codeset");
            field.setSubCodeset("subCodeset");
            field.setCodesetExtension("extension");
            field.setVisible(true);
            field.setMaxLength(i);
            field.setPrintId("printId" + i);
            field.setXpath("xpath" + i);
            fields.add(field);
        }
        return fields;
    }
}
