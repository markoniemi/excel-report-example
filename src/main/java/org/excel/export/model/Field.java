package org.excel.export.model;

import lombok.Data;

@Data
public class Field {
    private LocalizedText name;
    private boolean visible;
    private boolean mandatory;
    private String type;
    private String codeset;
    private String subCodeset;
    private String codesetExtension;
    private int minLength;
    private int maxLength;
    private int decimals;
    private String code;
    private String printId;
    private String xpath;
    private String metaFieldIdentifier;
}
