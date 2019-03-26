package org.excel.export;

import java.awt.Color;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.export.model.Element;
import org.excel.export.model.Field;

public class ElementSheet {
    protected static final String NAME = "Elements";
    private static final Color GREY = new java.awt.Color(192,192,192);
    protected Sheet sheet;
    protected int rowIndex;
    protected final CellStyle headerStyle;
    protected final CellStyle elementStyle;
    protected final CellStyle fieldStyle;
    protected CellStyle dateStyle;
    protected XSSFWorkbook workbook;

    public ElementSheet(XSSFWorkbook workbook, String name) {
        this.workbook = workbook;
        sheet = workbook.createSheet(name);
        sheet.createFreezePane(0, 1);
        headerStyle = createStyle(GREY);
        elementStyle = createStyle(IndexedColors.LIGHT_GREEN);
        fieldStyle = createStyle(IndexedColors.WHITE);
        dateStyle = createStyle(IndexedColors.WHITE);
        dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd.MM.yyyy"));
    }

    private CellStyle createStyle(IndexedColors color) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(color.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    private CellStyle createStyle(Color color) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(new XSSFColor(color));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    public void create(List<Element> elements) {
        createHeader();
        for (Element element : elements) {
            addElement(element);
            for (Field field : element.getFields()) {
                addField(field);
            }
        }
    }

    protected void createHeader() {
        Row row = sheet.createRow(rowIndex++);
        row.setRowStyle(headerStyle);
        int cellIndex = 0;
        setHeaderCell(row, cellIndex++, "Code", 1000);
        setHeaderCell(row, cellIndex++, "Name", 8000);
        setHeaderCell(row, cellIndex++, "Group", 4000);
        setHeaderCell(row, cellIndex++, "Mandatory", 1000);
        setHeaderCell(row, cellIndex++, "Type", 2000);
        setHeaderCell(row, cellIndex++, "Codeset", 2000);
        setHeaderCell(row, cellIndex++, "Subcodeset", 2000);
        setHeaderCell(row, cellIndex++, "Extension", 2000);
        setHeaderCell(row, cellIndex++, "Repeat", 2000);
        setHeaderCell(row, cellIndex++, "Repeat as group", 1000);
        setHeaderCell(row, cellIndex++, "Minimum", 1500);
        setHeaderCell(row, cellIndex++, "Maximum", 1500);
        setHeaderCell(row, cellIndex++, "Decimals", 1500);
        setHeaderCell(row, cellIndex++, "Code", 8000);
        setHeaderCell(row, cellIndex++, "PrintGroupId", 1500);
        setHeaderCell(row, cellIndex++, "PrintId", 1500);
        setHeaderCell(row, cellIndex++, "xpath", 1500);
        setHeaderCell(row, cellIndex++, "Name_en", 8000);
        setHeaderCell(row, cellIndex++, "Name_sv", 8000);
        setHeaderCell(row, cellIndex++, "Metafield", 2000);
    }

    protected void addElement(Element element) {
        Row row = sheet.createRow(rowIndex++);
        row.setRowStyle(elementStyle);
        int cellIndex = 0;
        setTextCell(row, cellIndex++, element.getCodeName());
        setTextCell(row, cellIndex++, element.getName().getFi());
        setTextCell(row, cellIndex++, element.getGroup());
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setNumberCell(row, cellIndex++, element.getRepeat());
        setBooleanCell(row, cellIndex++, element.isRepeatAsGroup());
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, element.getCode());
        setTextCell(row, cellIndex++, element.getGroupPrintId());
        setTextCell(row, cellIndex++, element.getPrintId());
        setTextCell(row, cellIndex++, element.getXpath());
        setTextCell(row, cellIndex++, element.getName().getEn());
        setTextCell(row, cellIndex++, element.getName().getSv());
        setTextCell(row, cellIndex++, null);
    }

    protected void addField(Field field) {
        Row row = sheet.createRow(rowIndex++);
        row.setRowStyle(fieldStyle);
        row.setHeight((short) (row.getHeight() * 3));
        int cellIndex = 0;
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, field.getName().getFi());
        setTextCell(row, cellIndex++, null);
        setBooleanCell(row, cellIndex++, field.isMandatory());
        setTextCell(row, cellIndex++, field.getType());
        setTextCell(row, cellIndex++, field.getCodeset());
        setTextCell(row, cellIndex++, field.getSubCodeset());
        setTextCell(row, cellIndex++, field.getCodesetExtension());
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, null);
        setNumberCell(row, cellIndex++, field.getMinLength());
        setNumberCell(row, cellIndex++, field.getMaxLength());
        setNumberCell(row, cellIndex++, field.getDecimals());
        setTextCell(row, cellIndex++, field.getCode());
        setTextCell(row, cellIndex++, null);
        setTextCell(row, cellIndex++, field.getPrintId());
        setTextCell(row, cellIndex++, field.getXpath());
        setTextCell(row, cellIndex++, field.getName().getEn());
        setTextCell(row, cellIndex++, field.getName().getSv());
        setTextCell(row, cellIndex++, field.getMetaFieldIdentifier());
    }

    protected void setHeaderCell(Row row, int cellIndex, String value, int columnWidth) {
        sheet.setColumnWidth(cellIndex, columnWidth);
        setTextCell(row, cellIndex, value);
    }

    protected void setTextCell(Row row, int cellIndex, String value) {
        setCellValue(row, cellIndex, value);
    }

    protected void setBooleanCell(Row row, int cellIndex, boolean value) {
        setCellValue(row, cellIndex, value ? "k" : "e");
    }

    protected void setNumberCell(Row row, int cellIndex, int value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(row.getRowStyle());
        cell.setCellValue(value);
    }

    protected void setDateCell(Row row, int cellIndex, Date value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(dateStyle);
        cell.setCellValue(value);
    }

    protected void setCellValue(Row row, int cellIndex, String value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(row.getRowStyle());
        cell.setCellValue(value);
    }
}
