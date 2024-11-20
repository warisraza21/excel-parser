package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class NonTableInfo {
    private String id;
    private int columnCount;
    private int rowCount;
    private CellRangeInfo boundaries;
    private List<ColumnInfo> columns;
    private List<List<String>> sampleRowData;
}
