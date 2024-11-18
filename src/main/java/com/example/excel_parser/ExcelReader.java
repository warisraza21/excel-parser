package com.example.excel_parser;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@Slf4j
public class ExcelReader {

    public static void testExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            log.info("Total sheets: {}", workbook.getNumberOfSheets());

            for (Sheet sheet : workbook) {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;

                log.info("Sheet Name : {}", xssfSheet.getSheetName());
                log.info("Sheet : {} contains total table(s) : {}", xssfSheet.getTables().size());
                log.info("Sheet : {} contains total pivot table(s) : {}", xssfSheet.getPivotTables().size());

                // Reserved areas
                List<AreaReference> reservedAreas = new ArrayList<>();

                // Process normal tables
                for (XSSFTable table : xssfSheet.getTables()) {
                    log.info("Processing Table: {}", table.getName());
                    log.info("Table Display Name: {}", table.getDisplayName());

                    AreaReference areaReference = table.getArea();
                    reservedAreas.add(areaReference); // Add table area to reserved areas
                    log.info("Area of table: {}", areaReference);

                    CellReference[] cellReferences = areaReference.getAllReferencedCells();

                    StringBuilder tableData = new StringBuilder("Table Data:\n");
                    for (CellReference cellRef : cellReferences) {
                        Row row = xssfSheet.getRow(cellRef.getRow());
                        if (row != null) {
                            Cell cell = row.getCell(cellRef.getCol());
                            if (cell != null) {
                                tableData.append(getCellValue(cell)).append("\t");
                            }
                        }
                        if (cellRef.getCol() == areaReference.getLastCell().getCol()) {
                            tableData.append("\n");
                        }
                    }
                    log.info("\n{}", tableData);
                }

                for (XSSFPivotTable pivotTable : xssfSheet.getPivotTables()) {
                    log.info("Processing Pivot Table");

                    AreaReference sourceArea = new AreaReference(
                            pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                            workbook.getSpreadsheetVersion()
                    );
                    reservedAreas.add(sourceArea);
                    log.info("Pivot Table Area: {}", sourceArea);

                    CellReference[] cellReferences = sourceArea.getAllReferencedCells();

                    StringBuilder pivotSourceData = new StringBuilder("Pivot Table Data:\n");
                    for (CellReference cellRef : cellReferences) {
                        Row row = xssfSheet.getRow(cellRef.getRow());
                        if (row != null) {
                            Cell cell = row.getCell(cellRef.getCol());
                            if (cell != null) {
                                pivotSourceData.append(getCellValue(cell)).append("\t");
                            }
                        }
                        if (cellRef.getCol() == sourceArea.getLastCell().getCol()) {
                            pivotSourceData.append("\n");
                        }
                    }
                    log.info("\n{}", pivotSourceData);
                }

                log.info("Processing Unstructured Data");
//                for (Row row : xssfSheet) {
//                    for (Cell cell : row) {
//                        if (!isWithinReservedArea(cell, reservedAreas)) {
//                            log.info("Unstructured Data - Cell [{}]: {}", new CellReference(cell).formatAsString(), getCellValue(cell));
//                        }
//                    }
//                }

                XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();

                if (drawing != null) {
                    List<XSSFChart> charts = drawing.getCharts();
                    for (XSSFChart chart : charts) {
                        String title = (chart.getTitle() != null ? chart.getTitle().toString(): "Untitled");
                        log.info("Chart Title : {}", title);
                    }
                }
            }
        } catch (IOException e) {
            log.error("Error processing Excel file", e);
        }
    }

    private static boolean isWithinReservedArea(Cell cell, List<AreaReference> reservedAreas) {
        for (AreaReference area : reservedAreas) {
            if (isWithinArea(cell, area)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isWithinArea(Cell cell, AreaReference area) {
        CellReference start = area.getFirstCell();
        CellReference end = area.getLastCell();

        int row = cell.getRowIndex();
        int col = cell.getColumnIndex();

        return row >= start.getRow() && row <= end.getRow() &&
                col >= start.getCol() && col <= end.getCol();
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}