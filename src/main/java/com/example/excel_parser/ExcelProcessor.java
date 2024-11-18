package com.example.excel_parser;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.ml.clustering.Cluster;
import org.apache.commons.math3.ml.clustering.DBSCANClusterer;
import org.apache.commons.math3.ml.clustering.DoublePoint;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import java.util.*;

@Slf4j
public class ExcelProcessor {

    public static void test(String filePath) throws IOException {
        // Load the Excel file
        FileInputStream file = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(file);
       for (Sheet sheet : workbook) {

           // Step 1: Extract all cell data
           List<CellData> cells = extractCellData(sheet);

           // Step 2: Identify structured tables and unstructured data
           ProcessedSheet processedSheet = processSheet(cells);

           log.info("Processing sheet {}", sheet.getSheetName());
           // Step 3: Print results
           log.info("Detected Structured Tables:");
           for (Table table : processedSheet.tables()) {
               log.info("Table : {}", table);
           }

           log.info("\nDetected Unstructured Data:");
           for (UnstructuredData data : processedSheet.unstructuredData()) {
               log.info("Unstructured table : {}", data);
           }
       }

        workbook.close();
    }

    // Step 1: Extract all cells
    private static List<CellData> extractCellData(Sheet sheet) {
        List<CellData> cellDataList = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell != null && !cell.toString().trim().isEmpty()) {
                    CellData cellData = new CellData(cell.getRowIndex(), cell.getColumnIndex());
                    cellData.setValue(cell.toString());
                    if (cell.getCellType() == CellType.FORMULA) {
                        cellData.setFormula(cell.getCellFormula());
                    }
                    cellDataList.add(cellData);
                }
            }
        }
        return cellDataList;
    }

    // Step 2: Process sheet into structured and unstructured data
    private static ProcessedSheet processSheet(List<CellData> cells) {
        // Convert cells to points for clustering
        List<DoublePoint> points = new ArrayList<>();
        for (CellData cell : cells) {
            points.add(new DoublePoint(new double[]{cell.getRowIndex(), cell.getColumnIndex()}));
        }

        // Apply DBSCAN to detect clusters (structured regions)
        DBSCANClusterer<DoublePoint> dbscan = new DBSCANClusterer<>(1.5, 2); // Adjust epsilon and minPoints
        List<Cluster<DoublePoint>> clusters = dbscan.cluster(points);

        // Identify structured tables
        List<Table> tables = new ArrayList<>();
        Set<DoublePoint> clusteredPoints = new HashSet<>();
        for (Cluster<DoublePoint> cluster : clusters) {
            Table table = new Table();
            for (DoublePoint point : cluster.getPoints()) {
                clusteredPoints.add(point);
                cells.stream()
                        .filter(c -> c.getRowIndex() == (int) point.getPoint()[0] && c.getColumnIndex() == (int) point.getPoint()[1])
                        .findFirst().ifPresent(table::addCell);
            }
            tables.add(table);
        }

        // Identify unstructured data (remaining points)
        List<UnstructuredData> unstructuredData = new ArrayList<>();
        for (CellData cell : cells) {
            DoublePoint point = new DoublePoint(new double[]{cell.getRowIndex(), cell.getColumnIndex()});
            if (!clusteredPoints.contains(point)) {
                unstructuredData.add(new UnstructuredData(cell));
            }
        }

        return new ProcessedSheet(tables, unstructuredData);
    }

    @Data
    static class CellData {
        private final int rowIndex;
        private final int columnIndex;
        private String value;
        private String formula;

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
                    '}';
        }
    }

    // Helper class to represent a table
    static class Table {
        private final List<CellData> cells = new ArrayList<>();

        public void addCell(CellData cell) {
            cells.add(cell);
        }

        @Override
        public String toString() {
            return "Table{" +
                    "cells=" + cells +
                    '}';
        }
    }

    // Helper class to represent unstructured data
    record UnstructuredData(CellData cellData) {
            @Override
            public String toString() {
                return "UnstructuredData{" +
                        "cellData=" + cellData +
                        '}';
            }
        }

    // Helper class to represent processed sheet data
    record ProcessedSheet(List<Table> tables, List<UnstructuredData> unstructuredData) {

    }
}

