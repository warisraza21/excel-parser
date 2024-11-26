package com.example.excel_parser.service;

import com.example.excel_parser.model.*;
import com.example.excel_parser.utils.DataTypeUtils;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Slf4j
public class ExcelReader {
    public static final ObjectMapper objectMapper = new ObjectMapper();
    public static FormulaEvaluator formulaEvaluator;

    static {
        objectMapper.enable(SerializationFeature.INDENT_OUTPUT);
    }

    public static void testExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            log.info("Total sheets: {}", workbook.getNumberOfSheets());

            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (Sheet sheet : workbook) {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;
                getNamedRangesInfo(workbook, sheet.getSheetName());

                log.info("SheetInfo Name : {}", xssfSheet.getSheetName());

                Set<AreaReference> reservedAreas = getAreaReferenceForTableAndPivotTable(xssfSheet, workbook.getSpreadsheetVersion());
                reservedAreas.forEach(areaReference -> log.info("Reserved Area : {}",areaReference));

                int rowCount = xssfSheet.getDimension().getLastRow() + 1;
                int colCount = xssfSheet.getDimension().getLastColumn() + 1;

                boolean[][] visited = new boolean[rowCount][colCount];
                markVisitedForReservedArea(visited,reservedAreas);


                List<CellData> list = new ArrayList<>();
                for (Row row : xssfSheet) {
                    for (Cell cell : row) {
                        if (!isWithinReservedArea(cell, reservedAreas)) {
                            CellData cellData = extractCellData(cell);
                            if(cellData != null) list.add(cellData);
                        }
                    }
                }

                if(!list.isEmpty()) {
                    ProcessedSheet processedSheet = ClusterCells.clusterCellsData(list, xssfSheet);
                    processedSheet.tableData().forEach(tableData -> {
                        Set<AreaReference> rangesAreaReference = RectangleDetector.getRangesAreaReference(xssfSheet,visited,tableData.getAreaReference());
                        rangesAreaReference.forEach(areaReference -> log.info("Range Area : {}",areaReference));
                    });
                }
            }
        } catch (IOException e) {
            log.error("Error processing Excel file", e);
        }
    }

    private static Set<AreaReference> getAreaReferenceForTableAndPivotTable(XSSFSheet sheet, SpreadsheetVersion spreadsheetVersion) {
        Set<AreaReference> reservedAreas = new HashSet<>();

        // Process normal tableData
        for (XSSFTable table : sheet.getTables()) {
            AreaReference areaReference = table.getArea();
            reservedAreas.add(areaReference);
        }

        // Process pivot tableData
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            AreaReference areaReference = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    spreadsheetVersion
            );
            reservedAreas.add(areaReference);
        }
        return reservedAreas;
    }

    // Helper method to get source table name
    private static String getSourceTableName(CTPivotCacheDefinition pivotCacheDef) {
        if (pivotCacheDef.getCacheSource() != null && pivotCacheDef.getCacheSource().isSetWorksheetSource()) {
            return pivotCacheDef.getCacheSource().getWorksheetSource().getRef();
        }
        return "No source table name found";
    }

    // Helper method to get column names
    private static void getColumnNames(CTPivotCacheDefinition pivotCacheDef) {
        if (pivotCacheDef.getCacheFields() != null) {
            CTCacheFields cacheFields = pivotCacheDef.getCacheFields();
            for (CTCacheField field : cacheFields.getCacheFieldArray()) {
                log.info(" - {}", field.getName());
            }
        } else {
            log.info("No columns found.");
        }
    }

    // Helper method to extract parent sheet name
    private static String getParentSheetName(CTPivotCacheDefinition pivotCacheDef) {
        if (pivotCacheDef.getCacheSource() != null && pivotCacheDef.getCacheSource().isSetWorksheetSource()) {
            CTWorksheetSource worksheetSource = pivotCacheDef.getCacheSource().getWorksheetSource();

            // Check if a sheet name is explicitly defined
            if (worksheetSource.isSetSheet()) {
                return worksheetSource.getSheet();
            }
        }
        return "No sheet name found";
    }

    private static boolean isWithinReservedArea(Cell cell, Set<AreaReference> reservedAreas) {
        for (AreaReference area : reservedAreas) {
            if (isWithinArea(cell, area.getFirstCell(), area.getLastCell())) {
                return true;
            }
        }
        return false;
    }

    private static boolean isWithinArea(Cell cell, CellReference start, CellReference end) {
        int row = cell.getRowIndex();
        int col = cell.getColumnIndex();

        return row >= start.getRow() && row <= end.getRow() &&
                col >= start.getCol() && col <= end.getCol();
    }

    private static CellData extractCellData(Cell cell) {
        CellData cellData = null;
        if (cell != null && !cell.toString().trim().isEmpty()) {
            cellData = new CellData(cell.getRowIndex(), cell.getColumnIndex());
            if (cell.getCellType() == CellType.FORMULA) {
                cellData.setFormula(cell.getCellFormula());
                cellData.setValue(DataTypeUtils.getCellValue(formulaEvaluator.evaluateInCell(cell)));
                cellData.setDataType(CellType.FORMULA.name());
            } else {
                cellData.setValue(DataTypeUtils.getCellValue(cell));
                DataTypeUtils.setCellDataType(cell, cellData);
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

        if (processedSheet.nonTableData() != null && !processedSheet.nonTableData().isEmpty()) {
            for (NonTableData data : processedSheet.nonTableData()) {
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
                    if (cell.getCellType() == CellType.FORMULA) {
                        tableData.append(DataTypeUtils.getCellValue(formulaEvaluator.evaluateInCell(cell))).append("\t");
                    } else {
                        tableData.append(DataTypeUtils.getCellValue(cell));
                    }
                }
            }
            if (cellRef.getCol() == areaReference.getLastCell().getCol()) {
                tableData.append("\n");
            }
        }
        log.info("{} {}", tableName, tableData);
    }

    private static void getNamedRangesInfo(Workbook workbook, String sheetName) {
        XSSFSheet sheet = (XSSFSheet) workbook.getSheet(sheetName);
        //To get named and normal ranges
        for (DataValidation validation : sheet.getDataValidations()) {
            DataValidationConstraint constraint = validation.getValidationConstraint();
            String formula = constraint.getFormula1();

            if (formula != null) {
                if (formula.contains(":")) {
                    log.info("Formula is a normal range: {}", formula);

                    // Parse and print the range
                    AreaReference areaRef = new AreaReference(formula, workbook.getSpreadsheetVersion());
                    CellReference firstCell = areaRef.getFirstCell();
                    CellReference lastCell = areaRef.getLastCell();

                    log.info("Range Details:");
                    log.info("SheetInfo: {}", firstCell.getSheetName());
                    log.info("Start Cell: {}", firstCell.formatAsString());
                    log.info("End Cell: {}", lastCell.formatAsString());
                } else {
                    log.info("Formula might be a named range: {}", formula);

                    // Check if the formula is a named range
                    if (formula.startsWith("=")) {
                        formula = formula.substring(1); // Remove '=' for named range lookup
                    }

                    Name namedRange = workbook.getName(formula);
                    if (namedRange != null) {
                        log.info("Named Range Found: {}", namedRange.getNameName());

                        // Resolve the named range to a range
                        AreaReference areaRef = new AreaReference(namedRange.getRefersToFormula(), workbook.getSpreadsheetVersion());
                        CellReference firstCell = areaRef.getFirstCell();
                        CellReference lastCell = areaRef.getLastCell();

                        log.info("Named Range Details:");
                        log.info("SheetInfo-: {}", firstCell.getSheetName());
                        log.info("Start -Cell: {}", firstCell.formatAsString());
                        log.info("End -Cell: {}", lastCell.formatAsString());
                    } else {
                        log.info("Formula is not a named range or a valid range.");
                    }
                }
            } else {
                log.info("Validation does not use a formula.");
            }
        }
    }

    public static void markVisitedForReservedArea(boolean[][] visited, Set<AreaReference> reservedAreas) {
        for (AreaReference area : reservedAreas) {
            // Get the top-left and bottom-right corners of the area
            CellReference firstCell = area.getFirstCell();
            CellReference lastCell = area.getLastCell();

            // Get row and column indices
            int startRow = firstCell.getRow();
            int endRow = lastCell.getRow();
            int startCol = firstCell.getCol();
            int endCol = lastCell.getCol();

            // Mark the corresponding cells in the visited array
            for (int row = startRow; row <= endRow; row++) {
                for (int col = startCol; col <= endCol; col++) {
                    visited[row][col] = true;
                }
            }
        }
    }

    public static List<CellCoordinate> getUnvisitedCells(boolean[][] visited) {
        List<CellCoordinate> unvisitedCells = new ArrayList<>();

        for (int row = 0; row < visited.length; row++) {
            for (int col = 0; col < visited[row].length; col++) {
                if (!visited[row][col]) {
                    // Create a CellCoordinate object for this unvisited cell
                    unvisitedCells.add(new CellCoordinate(row, col));
                }
            }
        }

        return unvisitedCells;
    }

}