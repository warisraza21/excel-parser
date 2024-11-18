package com.example.excel_parser;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Slf4j
public class ExcelReader1 {

    public static void processExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            log.info("Processing file: {}", filePath);

            for (Sheet sheet : workbook) {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;

                log.info("\nSheet Name: {}", xssfSheet.getSheetName());

                // Detect Tables
                detectTables(xssfSheet);

                // Detect Pivot Tables
                detectPivotTables(xssfSheet, workbook);

                // Detect Charts
                detectCharts(xssfSheet);

                // Detect Unstructured Data
                detectUnstructuredData(xssfSheet);
            }

        } catch (IOException e) {
            log.error("Error processing Excel file", e);
        }
    }

    private static void detectTables(XSSFSheet sheet) {
        List<XSSFTable> tables = sheet.getTables();
        if (tables.isEmpty()) {
            log.info("No tables found in sheet: {}", sheet.getSheetName());
        } else {
            log.info("Found {} table(s) in sheet: {}", tables.size(), sheet.getSheetName());
            for (XSSFTable table : tables) {
                log.info("Table Name: {}", table.getName());
                log.info("Display Name: {}", table.getDisplayName());
                log.info("Table Area: {}", table.getArea());
            }
        }
    }

    private static void detectPivotTables(XSSFSheet sheet, Workbook workbook) {
        List<XSSFPivotTable> pivotTables = sheet.getPivotTables();
        if (pivotTables.isEmpty()) {
            log.info("No pivot tables found in sheet: {}", sheet.getSheetName());
        } else {
            log.info("Found {} pivot table(s) in sheet: {}", pivotTables.size(), sheet.getSheetName());
            for (XSSFPivotTable pivotTable : pivotTables) {
                AreaReference sourceArea = new AreaReference(
                        pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                        workbook.getSpreadsheetVersion()
                );
                log.info("Pivot Table Source Area: {}", sourceArea.formatAsString());
            }
        }
    }

    private static void detectCharts(XSSFSheet sheet) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null || drawing.getCharts().isEmpty()) {
            log.info("No charts found in sheet: {}", sheet.getSheetName());
        } else {
            List<XSSFChart> charts = drawing.getCharts();
            log.info("Found {} chart(s) in sheet: {}", charts.size(), sheet.getSheetName());
            for (XSSFChart chart : charts) {
                log.info("Chart Title: {}", chart.getPackagePart().getPartName().getName() != null ? chart.getTitle() : "Untitled");
            }
        }
    }

    private static void detectUnstructuredData(XSSFSheet sheet) {
        log.info("Detecting unstructured data in sheet: {}", sheet.getSheetName());

        Set<CellAddress> reservedCells = new HashSet<>();

        // Mark all cells within tables
        for (XSSFTable table : sheet.getTables()) {
            reservedCells.addAll(getCellsInArea(table.getArea(), sheet));
        }

        // Mark all cells within pivot table sources
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            AreaReference pivotArea = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    sheet.getWorkbook().getSpreadsheetVersion()
            );
            reservedCells.addAll(getCellsInArea(pivotArea, sheet));
        }

        List<String> unstructuredData = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (!reservedCells.contains(cell.getAddress()) && !isCellEmpty(cell)) {
                    unstructuredData.add(
                            String.format("Cell [%s]: %s", cell.getAddress().formatAsString(), getCellValue(cell))
                    );
                }
            }
        }

        if (unstructuredData.isEmpty()) {
            log.info("No unstructured data found in sheet: {}", sheet.getSheetName());
        } else {
            log.info("Unstructured Data Found:");
            for (String data : unstructuredData) {
//                log.info(data);
            }
        }
    }

    private static List<CellAddress> getCellsInArea(AreaReference area, XSSFSheet sheet) {
        List<CellAddress> cells = new ArrayList<>();
        CellReference[] cellReferences = area.getAllReferencedCells();
        for (CellReference cellRef : cellReferences) {
            Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
            if (cell != null) {
                cells.add(cell.getAddress());
            }
        }
        return cells;
    }

    private static boolean isCellEmpty(Cell cell) {
        if (cell == null) return true;
        if (cell.getCellType() == CellType.BLANK) return true;
        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) return true;
        return false;
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return DateUtil.isCellDateFormatted(cell)
                        ? cell.getDateCellValue().toString()
                        : String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
