package com.example.excel_parser.model;

import lombok.Data;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

@Data
public class TableData implements Serializable {
    private final List<CellData> cells = new ArrayList<>();

    public void addCell(CellData cell) {
        cells.add(cell);
    }

    @Override
    public String toString() {
        return "TableData{" +
                "cells=" + cells +
                '}';
    }
}