package org.excel.export.model;

import java.util.Date;
import java.util.List;

import lombok.Data;

@Data
public class Element {
    private String codeName;
    private LocalizedText name;
    private String group;
    private List<Field> fields;
    private boolean visible;
    private int repeat;
    private boolean repeatAsGroup;
    private String code;
    private String groupPrintId;
    private String printId;
    private String xpath;
    private Date startDate;
}
