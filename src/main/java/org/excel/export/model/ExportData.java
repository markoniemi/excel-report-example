package org.excel.export.model;

import java.util.List;

import lombok.Data;
@Data
public class ExportData {
    private List<Group> groups;
    private List<Element> elements;
    private List<Document> documents;
}
