package com.example.excel_parser.model;

import lombok.Data;

@Data
public class CellCoordinate {
    private final int rowIndex;
    private final int columnIndex;

    public CellCoordinate(int rowIndex, int columnIndex) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
    }
}
