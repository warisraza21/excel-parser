package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class ChartSourceData {
    private String tableName;
    private CellRangeInfo boundaries;
    private List<ChartSourceColumn> columns;
}

@Data
class ChartSourceColumn {
    private String name;
    private String dataType;  // String, Numeric, etc.
    private String role;      // XAxis, YAxis
}
