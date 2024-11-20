package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class ColumnInfo {
    private String name;
    private String type;
    private boolean isDerived;
    private Object formula;  // Can be a string or a complex formula object
    private DataValidationRules dataValidation;
    private CellFormatting formatting;
}

