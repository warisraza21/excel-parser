package com.example.excel_parser;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.ml.clustering.Cluster;
import org.apache.commons.math3.ml.clustering.DBSCANClusterer;
import org.apache.commons.math3.ml.clustering.DoublePoint;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.util.*;

@Slf4j
public class ExcelProcessor  {

    public static void test(ProcessedSheet processedSheet) throws IOException {
        // Initialize Jackson ObjectMapper
        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.enable(SerializationFeature.INDENT_OUTPUT); // Enable pretty printing

        if (processedSheet.tables() != null && !processedSheet.tables.isEmpty()) {
            for (Table table : processedSheet.tables()) {
                String json = objectMapper.writeValueAsString(table);
                log.info("Structured Table:\n{}", json);
            }
        }

        if (processedSheet.unstructuredData() != null && !processedSheet.unstructuredData().isEmpty()) {
            for (UnstructuredData data : processedSheet.unstructuredData()) {
                String json = objectMapper.writeValueAsString(data);
                log.info("Unstructured Data:\n{}", json);
            }
        }
    }


    // Step 2: Process sheet into structured and unstructured data
    public static ProcessedSheet processSheet(List<CellData> cells) {
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
    static class CellData implements Serializable{
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
    @Data
    static class Table implements Serializable{
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
    record UnstructuredData(CellData cellData) implements Serializable{
        @Override
        public String toString() {
            return "UnstructuredData{" +
                    "cellData=" + cellData +
                    '}';
        }
    }

    // Helper class to represent processed sheet data
    record ProcessedSheet(List<Table> tables, List<UnstructuredData> unstructuredData) implements Serializable{

    }
}

