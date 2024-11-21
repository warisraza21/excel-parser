package com.example.excel_parser.model;

import com.example.excel_parser.dtos.CellRangeInfo;
import com.example.excel_parser.dtos.ColumnInfo;
import lombok.Data;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

@Data
public class TableData implements Serializable {

    private int rowCount;
    private int columnCount;
    private CellRangeInfo boundaries;
    private List<ColumnInfo> columns = new ArrayList<>();
    private  final List<CellData> cells = new ArrayList<>();

    public void addCell(CellData data){
        this.cells.add(data);
    }

    @Override
    public String toString() {
        return "TableData{" +
                "cells=" + cells +
                ", boundaries=" + boundaries +
                ", rowCount=" + rowCount +
                ", columnCount=" + columnCount +
                ", columns=" + columns +
                '}';
    }
}