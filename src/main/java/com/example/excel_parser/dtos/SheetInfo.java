package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class SheetInfo {
    private String id;
    private String name;
    private int sheetIndex;
    private int tableCounts;
    private int chartCounts;
    private int pivotTableCount;
    private int nonTableCount;
    private Boolean visibility;
    private List<TableInfo> tables;
    private List<NonTableInfo> nonTables;
    private List<PivotTableInfo> pivotTables;
    private List<NamedRangeInfo> namedRanges;
    private List<ChartInfo> charts;
}

