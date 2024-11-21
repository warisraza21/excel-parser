package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class TableInfo {
    private String id;
    private String name;
    private int columnCount;
    private int rowCount;
    private CellRangeInfo boundaries;
    private List<ColumnInfo> columns;
    private List<MergedCellInfo> mergedCells;
    private TableHeaders headers;
    private List<List<?>> sampleRowData;
    private TableSummaryRow summaryRow;
}

