package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class PivotTableOptions {
    private PivotTableGrandTotals grandTotals;
}

@Data
class PivotTableGrandTotals {
    private boolean rows;
    private boolean columns;
}
