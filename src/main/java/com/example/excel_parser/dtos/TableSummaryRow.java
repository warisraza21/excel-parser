package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class TableSummaryRow {
    private boolean totalRowPresent;
    private int totalRowIndex;
    private List<ColumnAggregation> columnAggregations;
}

@Data
class ColumnAggregation {
    private String columnName;
    private String formula;
    private String value;
}
