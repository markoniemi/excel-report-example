package org.excel.export;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.excel.export.model.Field;
import org.excel.export.model.Element;

public class ElementSheet {
    protected static final String NAME = "Elements";
    protected Sheet sheet;
    protected int rowIndex;
    protected final CellStyle headerStyle;
    protected final CellStyle elementStyle;
    protected final CellStyle fieldStyle;
    protected CellStyle dateStyle;

    public ElementSheet(Workbook workbook, String name) {
        sheet = workbook.createSheet(name);
        sheet.createFreezePane(0, 1);
        headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        elementStyle = workbook.createCellStyle();
        elementStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
        elementStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        fieldStyle = workbook.createCellStyle();
        fieldStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        fieldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        dateStyle = workbook.createCellStyle();
        dateStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
        dateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd.MM.yyyy"));
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
        int cellIndex = 0;
        setTextCell(row, cellIndex++, null, fieldStyle);
        setTextCell(row, cellIndex++, field.getName().getFi(), fieldStyle);
        setTextCell(row, cellIndex++, null, fieldStyle);
        setBooleanCell(row, cellIndex++, field.isMandatory(), fieldStyle);
        setTextCell(row, cellIndex++, field.getType(), fieldStyle);
        setTextCell(row, cellIndex++, field.getCodeset(), fieldStyle);
        setTextCell(row, cellIndex++, field.getSubCodeset(), fieldStyle);
        setTextCell(row, cellIndex++, field.getCodesetExtension(), fieldStyle);
        setTextCell(row, cellIndex++, null, fieldStyle);
        setTextCell(row, cellIndex++, null, fieldStyle);
        setNumberCell(row, cellIndex++, field.getMinLength(), fieldStyle);
        setNumberCell(row, cellIndex++, field.getMaxLength(), fieldStyle);
        setNumberCell(row, cellIndex++, field.getDecimals(), fieldStyle);
        setTextCell(row, cellIndex++, field.getCode(), fieldStyle);
        setTextCell(row, cellIndex++, null, fieldStyle);
        setTextCell(row, cellIndex++, field.getPrintId(), fieldStyle);
        setTextCell(row, cellIndex++, field.getXpath(), fieldStyle);
        setTextCell(row, cellIndex++, field.getName().getEn(), fieldStyle);
        setTextCell(row, cellIndex++, field.getName().getSv(), fieldStyle);
        setTextCell(row, cellIndex++, field.getMetaFieldIdentifier(), fieldStyle);
    }

    protected void setHeaderCell(Row row, int cellIndex, String value, int columnWidth) {
        sheet.setColumnWidth(cellIndex, columnWidth);
        setTextCell(row, cellIndex, value, headerStyle);
    }

    protected void setTextCell(Row row, int cellIndex, String value) {
        setTextCell(row, cellIndex, value, elementStyle);
    }

    protected void setTextCell(Row row, int cellIndex, String value, CellStyle style) {
        setCellValue(row, cellIndex, value, style);
    }

    protected void setBooleanCell(Row row, int cellIndex, boolean value) {
        setBooleanCell(row, cellIndex, value, elementStyle);
    }

    protected void setBooleanCell(Row row, int cellIndex, boolean value, CellStyle style) {
        setCellValue(row, cellIndex, value ? "k" : "e", style);
    }

    protected void setNumberCell(Row row, int cellIndex, int value) {
        setNumberCell(row, cellIndex, value, elementStyle);
    }

    protected void setNumberCell(Row row, int cellIndex, int value, CellStyle style) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(style);
        cell.setCellValue(value);
    }

    protected void setDateCell(Row row, int cellIndex, Date value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(dateStyle);
        cell.setCellValue(value);
    }

    protected void setCellValue(Row row, int cellIndex, String value, CellStyle style) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(style);
        cell.setCellValue(value);
    }
}
