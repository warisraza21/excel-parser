package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class PivotTableSource {
    private String sourceType;  // E.g., Table, Range
    private String tableName;
    private String tableSheetName;
    private String tableRange;
}
