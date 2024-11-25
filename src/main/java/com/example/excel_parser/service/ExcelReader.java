package com.example.excel_parser.service;

import com.example.excel_parser.model.CellData;
import com.example.excel_parser.model.ProcessedSheet;
import com.example.excel_parser.model.TableData;
import com.example.excel_parser.model.NonTableData;
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
import java.util.stream.Collectors;

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

                Set<AreaReference> reservedAreas = getReservedAreas(xssfSheet, workbook.getSpreadsheetVersion());
                detectCharts(xssfSheet);

                List<CellData> list = new ArrayList<>();
                for (Row row : xssfSheet) {
                    for (Cell cell : row) {
                        if (!isWithinReservedArea(cell, reservedAreas)) {
                            CellData cellData = extractCellData(cell);
                            if (cellData != null) list.add(cellData);
                        }
                    }
                }
                processUnNamedRanges(xssfSheet,list);
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

    private static Set<AreaReference> getReservedAreas(XSSFSheet sheet, SpreadsheetVersion spreadsheetVersion) {
        Set<AreaReference> reservedAreas = new HashSet<>();

        // Process normal tableData
        for (XSSFTable table : sheet.getTables()) {
            AreaReference areaReference = table.getArea();
            printTableData(sheet, areaReference, table.getName());
            reservedAreas.add(areaReference);
            log.info("Table: {} added to reserved areas [{}:{}]", table.getName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
        }

        // Process pivot tableData
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            CTPivotCacheDefinition pivotCacheDef = pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition();

            // Extract source table name
            String sourceTableName = getSourceTableName(pivotCacheDef);
            log.info("Source Table Name: {}", sourceTableName);
            String parentSheetName = getParentSheetName(pivotCacheDef);
            log.info("Parent Table SheetInfo Name: {}", parentSheetName);

            // Extract column names
            log.info("Columns:");
            getColumnNames(pivotCacheDef);
            AreaReference areaReference = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    spreadsheetVersion
            );
            printTableData(sheet, areaReference, pivotTable.getCTPivotTableDefinition().getName());
            reservedAreas.add(areaReference);

            log.info("Pivot Table of SheetInfo : {} added to reserved areas [{}:{}]", sheet.getSheetName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
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
            if (isWithinArea(cell, area.getFirstCell(),area.getLastCell())) {
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

    private static void processUnNamedRanges(XSSFSheet xssfSheet, List<CellData> list) throws JsonProcessingException {
        if (!list.isEmpty()) {
            ProcessedSheet processedSheet = ClusterCells.clusterCellsData(list,xssfSheet);

            Set<AreaReference> unNamedAreaReference = processedSheet.tableData().stream().map(TableData::getAreaReference).collect(Collectors.toSet());

            for (TableData tableData : processedSheet.tableData()) {
                List<int[][]> boundaryCellCoordinateList = getFirstRowBoundaryLineList(xssfSheet, tableData.getFirstCell(), tableData.getLastCell());


                for (int[][] line: boundaryCellCoordinateList){
                    int firstCol = line[0][1];
                    int lastCol = line[1][1];

                    int firstRow = line[0][0];
                    int lastRow = tableData.getLastCell()[0];

                    int lastRowIndex = getLastRowIndex(xssfSheet,firstRow,lastRow,firstCol,lastCol);

                    log.info("Boundary : [ {}, {}] : [ {}, {}]",firstRow,firstCol,lastRowIndex,lastCol);

                }
            }
            printClusteredData(processedSheet);
        }
    }

    private static List<int[][]> getFirstRowBoundaryLineList(XSSFSheet sheet, int[] firstCellIndex, int[] lastCellIndex) {
        int firstRow = firstCellIndex[0];
        int firstCol = firstCellIndex[1];

        int lastRow = lastCellIndex[0];
        int lastCol = lastCellIndex[1];

        // Map to store min and max column indices for each row
        Map<Integer, int[]> rowToColIndicesMap = new HashMap<>();

        Row row = sheet.getRow(firstRow);
        for (int colIndex = firstCol; colIndex <= lastCol; colIndex++) {
            Cell cell = row.getCell(colIndex);
            if (cell == null || cell.getCellType() == CellType.BLANK) {
                int[] coordinate = getBoundaryCellCoordinate(sheet, firstRow, lastRow, colIndex);
                if (coordinate[0] != -1 && coordinate[1] != -1) {
                    int rowIndex = coordinate[0];
                    int colIndexFound = coordinate[1];

                    // Update min and max colIndex for the row
                    if (!rowToColIndicesMap.containsKey(rowIndex)) {
                        rowToColIndicesMap.put(rowIndex, new int[]{colIndexFound, colIndexFound}); // [minCol, maxCol]
                    } else {
                        int[] colIndices = rowToColIndicesMap.get(rowIndex);
                        colIndices[0] = Math.min(colIndices[0], colIndexFound); // Update minCol
                        colIndices[1] = Math.max(colIndices[1], colIndexFound); // Update maxCol
                    }
                }
            }
        }

        // Prepare the result list
        List<int[][]> lineEndpointList = new ArrayList<>();

        for (Map.Entry<Integer, int[]> entry : rowToColIndicesMap.entrySet()) {
            int rowIndex = entry.getKey();
            int[] colIndices = entry.getValue();

            // Endpoints of the row line
            int[] firstPoint = new int[]{rowIndex, colIndices[0]};
            int[] secondPoint = new int[]{rowIndex, colIndices[1]};

            // Check for consistency on the left and right of the line
            int leftEndpoint = checkLeftConsistency(sheet, rowIndex, colIndices[0], firstCol);
            int rightEndpoint = checkRightConsistency(sheet, rowIndex, colIndices[1], lastCol);

            int[][] lineEndPoint = new int[2][2];
            lineEndPoint[0] = firstPoint;
            lineEndPoint[1] = secondPoint;

            // If data is consistent, store the endpoints
            if (leftEndpoint != -1 && rightEndpoint != -1) {
                //set line first and last point y co-ordinate
                firstPoint[1] = leftEndpoint;
                secondPoint[1] = rightEndpoint;
            }

            lineEndpointList.add(lineEndPoint);
        }

        return lineEndpointList;
    }

    private static int checkLeftConsistency(XSSFSheet sheet, int rowIndex, int colStartIndex, int firstCol) {
        int latestColIndex = colStartIndex;  // Track the last valid column index

        // Iterate over the columns to the left of the line start
        for (int colIndex = colStartIndex - 1; colIndex >= firstCol; colIndex--) {
            Cell cell = sheet.getRow(rowIndex).getCell(colIndex);

            if (cell == null || cell.getCellType() == CellType.BLANK) {
                return colIndex + 1; // Return the current column index if the cell is empty
            }

            if (DataTypeUtils.checkDataTypeRow(DataTypeUtils.getCellValue(cell))) {
                return -1; // Return -1 if the cell is not a String
            }

            latestColIndex = colIndex;  // Update the latest valid column index
        }

        return latestColIndex; // Return the last valid column index
    }

    private static int checkRightConsistency(XSSFSheet sheet, int rowIndex, int colEndIndex, int lastCol) {
        int latestColIndex = colEndIndex;  // Track the last valid column index

        // Iterate over the columns to the right of the line end
        for (int colIndex = colEndIndex + 1; colIndex <= lastCol; colIndex++) {
            Cell cell = sheet.getRow(rowIndex).getCell(colIndex);

            if (cell == null || cell.getCellType() == CellType.BLANK) {
                return colIndex - 1; // Return the current column index if the cell is empty
            }

            if (DataTypeUtils.checkDataTypeRow(DataTypeUtils.getCellValue(cell))) {
                return -1; // Return -1 if the cell is not a String
            }

            latestColIndex = colIndex;  // Update the latest valid column index
        }

        return latestColIndex; // Return the last valid column index
    }

    private static int[] getBoundaryCellCoordinate(XSSFSheet sheet, int rowStart, int rowEnd, int colIndex) {
        int[] coordinate = new int[]{-1, -1};
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null && (cell.getCellType() != CellType.BLANK)) {
                    coordinate[0] = rowIndex;
                    coordinate[1] = colIndex;
                    break;
                }
            }
        }
        return coordinate;
    }


    public static int getLastRowIndex(Sheet sheet, int rowStartIndex, int rowEndIndex, int colStartIndex, int colEndIndex) {
        int []lastRowIndexes = new int[colEndIndex - colStartIndex + 1];
        int idx = 0;
        // Traverse each column within the specified range
        for (int colIndex = colStartIndex; colIndex <= colEndIndex; colIndex++) {
            String initialDataType = null; // To store the initial column data type
            int maxRowIdx = -1;
            for (int rowIndex = rowStartIndex + 1; rowIndex <= rowEndIndex; rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    Cell cell = row.getCell(colIndex);
                    if(cell != null) {
                        String cellValue = DataTypeUtils.getCellValue(cell); // Extract cell value as string

                        if (cellValue != null && !cellValue.trim().isEmpty()) {
                            String currentDataType = DataTypeUtils.detectDataType(cellValue);

                            if (initialDataType == null) {
                                initialDataType = currentDataType;
                            } else if (!initialDataType.equals(currentDataType)) {
                                break;
                            }
                        }
                        maxRowIdx = rowIndex;
                    }
                }
            }
            lastRowIndexes[idx++] = maxRowIdx;
        }

        int minRow = Integer.MAX_VALUE;
        for(int item : lastRowIndexes) minRow = Math.min(minRow,item);
        return minRow == Integer.MAX_VALUE ? -1 : minRow; // Return -1 if no valid rows are found
    }

}