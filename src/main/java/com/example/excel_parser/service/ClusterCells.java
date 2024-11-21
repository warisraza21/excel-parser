package com.example.excel_parser.service;

import com.example.excel_parser.dtos.CellRangeInfo;
import com.example.excel_parser.model.CellData;
import com.example.excel_parser.model.ProcessedSheet;
import com.example.excel_parser.model.TableData;
import com.example.excel_parser.model.NonTableData;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.ml.clustering.Cluster;
import org.apache.commons.math3.ml.clustering.DBSCANClusterer;
import org.apache.commons.math3.ml.clustering.DoublePoint;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;


import java.util.*;

@Slf4j
public class ClusterCells {

    // Step 2: Process sheet into structured and unstructured data
    public static ProcessedSheet clusterCellsData(List<CellData> cells) {
        // Convert cells to points for clustering
        List<DoublePoint> points = new ArrayList<>();
        for (CellData cell : cells) {
            points.add(new DoublePoint(new double[]{cell.getRowIndex(), cell.getColumnIndex()}));
        }

        // Apply DBSCAN to detect clusters (structured regions)
        DBSCANClusterer<DoublePoint> dbscan = new DBSCANClusterer<>(1.5, 2);
        List<Cluster<DoublePoint>> clusters = dbscan.cluster(points);

        // Identify structured tableData
        List<TableData> tables = new ArrayList<>();
        Set<DoublePoint> clusteredPoints = new HashSet<>();
        for (Cluster<DoublePoint> cluster : clusters) {
            TableData table = new TableData();
            for (DoublePoint point : cluster.getPoints()) {
                clusteredPoints.add(point);
                cells.stream()
                        .filter(c -> c.getRowIndex() == (int) point.getPoint()[0] && c.getColumnIndex() == (int) point.getPoint()[1])
                        .findFirst().ifPresent(table::addCell);
            }
            if(!table.getCells().isEmpty()){
                calculateBoundary(table);
            }
            tables.add(table);
        }

        // Identify unstructured data (remaining points)
        List<NonTableData> nonTableData = new ArrayList<>();
        for (CellData cell : cells) {
            DoublePoint point = new DoublePoint(new double[]{cell.getRowIndex(), cell.getColumnIndex()});
            if (!clusteredPoints.contains(point)) {
                nonTableData.add(new NonTableData(cell));
            }
        }

        return new ProcessedSheet(tables, nonTableData);
    }
    private static void calculateBoundary(TableData table) {
        int minRow = Integer.MAX_VALUE, maxRow = Integer.MIN_VALUE;
        int minCol = Integer.MAX_VALUE, maxCol = Integer.MIN_VALUE;

        for (CellData cell : table.getCells()) {
            int rowIndex = cell.getRowIndex();
            int colIndex = cell.getColumnIndex();

            // Update boundaries
            minRow = Math.min(minRow, rowIndex);
            maxRow = Math.max(maxRow, rowIndex);
            minCol = Math.min(minCol, colIndex);
            maxCol = Math.max(maxCol, colIndex);
        }

        // Set rowCount and columnCount in the TableData
        table.setRowCount(maxRow - minRow + 1); // Inclusive range
        table.setColumnCount(maxCol - minCol + 1); // Inclusive range

        // Use CellReference for boundary representation
        CellReference startCell = new CellReference(minRow, minCol);
        CellReference endCell = new CellReference(maxRow, maxCol);
        table.setBoundaries(new CellRangeInfo(startCell.formatAsString(), endCell.formatAsString()));
    }

}

