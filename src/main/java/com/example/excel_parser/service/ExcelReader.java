package com.example.excel_parser.service;

import com.example.excel_parser.model.CellData;
import com.example.excel_parser.model.ProcessedSheet;
import com.example.excel_parser.model.TableData;
import com.example.excel_parser.model.UnstructuredData;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
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
    public static final ObjectMapper objectMapper = new ObjectMapper();

    static {
        objectMapper.enable(SerializationFeature.INDENT_OUTPUT);
    }

    public static void testExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            log.info("Total sheets: {}", workbook.getNumberOfSheets());

            for (Sheet sheet : workbook) {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;

                log.info("Sheet Name : {}", xssfSheet.getSheetName());

                List<AreaReference> reservedAreas = getReservedAreas(xssfSheet, workbook.getSpreadsheetVersion());
                detectCharts(xssfSheet);

                List<CellData> list = new ArrayList<>();
                for (Row row : xssfSheet) {
                    for (Cell cell : row) {
                        if (!isWithinReservedArea(cell, reservedAreas)) {
                            CellData cellData = extractCellData(cell);
                            if(cellData != null) list.add(cellData);
                        }
                    }
                }

                if(!list.isEmpty()){
                    ProcessedSheet processedSheet =  ClusterCells.clusterCellsData(list);
                    printClusteredData(processedSheet);
                }
            }
        } catch (IOException e) {
            log.error("Error processing Excel file", e);
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
                log.info("Chart Title: {}", chart.getPackagePart().getPartName().getName());
            }
        }
    }

    private static List<AreaReference> getReservedAreas(XSSFSheet sheet, SpreadsheetVersion spreadsheetVersion) {
        List<AreaReference> reservedAreas = new ArrayList<>();

        // Process normal tableData
        for (XSSFTable table : sheet.getTables()) {
            AreaReference areaReference = table.getArea();
            printTableData(sheet, areaReference, table.getName());
            reservedAreas.add(areaReference);
            log.info("Table: {} added to reserved areas [{}:{}]", table.getName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
        }

        // Process pivot tableData
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            AreaReference areaReference = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    spreadsheetVersion
            );
            printTableData(sheet, areaReference, pivotTable.getCTPivotTableDefinition().getName());
            reservedAreas.add(areaReference);
            log.info("Pivot Table of Sheet : {} added to reserved areas [{}:{}]", sheet.getSheetName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
        }
        return reservedAreas;
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

    private static CellData extractCellData(Cell cell) {
        CellData cellData = null;
        if (cell != null && !cell.toString().trim().isEmpty()) {
            cellData = new CellData(cell.getRowIndex(), cell.getColumnIndex());
            cellData.setValue(cell.toString());
            if (cell.getCellType() == CellType.FORMULA) {
                cellData.setFormula(cell.getCellFormula());
            }
        }
        return cellData;
    }

    public static void printClusteredData(ProcessedSheet processedSheet) throws JsonProcessingException {
        if (processedSheet.tableData() != null && !processedSheet.tableData().isEmpty()) {
            for (TableData table : processedSheet.tableData()) {
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

    public static void printTableData(XSSFSheet sheet, AreaReference areaReference, String tableName) {
        CellReference[] cellReferences = areaReference.getAllReferencedCells();

        StringBuilder tableData = new StringBuilder("Table Data:\n");
        for (CellReference cellRef : cellReferences) {
            Row row = sheet.getRow(cellRef.getRow());
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
        log.info("{} {}",tableName,tableData);
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
