package com.example.excel_parser.model;

import lombok.Data;

import java.io.Serializable;

@Data
public class CellData implements Serializable {
    private final int rowIndex;
    private final int columnIndex;
    private String value;
    private String formula;
    private String dataType;

    public CellData(int rowIndex, int columnIndex) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
    }

    @Override
    public String toString() {
        return "CellData{" +
                "rowIndex=" + rowIndex +
                ", columnIndex=" + columnIndex +
                ", value='" + value + '\'' +
                ", formula='" + formula + '\'' +
                ", dataType='" + dataType + '\''+
                '}';
    }
}

