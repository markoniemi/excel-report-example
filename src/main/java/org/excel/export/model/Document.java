package org.excel.export.model;

import java.util.List;

import lombok.Data;

@Data
public class Document {
    private String code;
    private LocalizedText name;
    private List<Element> elements;
}
