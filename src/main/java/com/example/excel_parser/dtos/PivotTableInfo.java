package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class PivotTableInfo {
    private String id;
    private String name;
    private CellRangeInfo boundaries;
    private PivotTableSource source;
    private PivotTableFields fields;
    private PivotTableOptions options;
}

