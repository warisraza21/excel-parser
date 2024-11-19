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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
                setHiddenHeader(xssfSheet, filePath);
                getRangesInfo(workbook, sheet.getSheetName());

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
            CTPivotCacheDefinition pivotCacheDef = pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition();

            // Extract source table name
            String sourceTableName = getSourceTableName(pivotCacheDef);
            log.info("Source Table Name: {}", sourceTableName);
            String parentSheetName = getParentSheetName(pivotCacheDef);
            log.info("Parent Table Sheet Name: {}", parentSheetName);

            // Extract column names
            log.info("Columns:");
            getColumnNames(pivotCacheDef);
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
            if (cell.getCellType() == CellType.FORMULA) {
                cellData.setFormula(cell.getCellFormula());
                cellData.setValue(getCellValue(formulaEvaluator.evaluateInCell(cell)));
            } else {
                cellData.setValue(getCellValue(cell));
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
                    if (cell.getCellType() == CellType.FORMULA) {
                        tableData.append(getCellValue(formulaEvaluator.evaluateInCell(cell))).append("\t");
                    } else {
                        tableData.append(getCellValue(cell));
                    }
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

    private static void setHiddenHeader(XSSFSheet sheet, String outputFilePath) {
        for (XSSFTable table : sheet.getTables()) {
            CTTable ctTable = table.getCTTable();

            // Check if the header row is hidden
            long headerRowCount = ctTable.getHeaderRowCount();
            if (headerRowCount == 0) {
                log.info("Header row is hidden. Attempting to un hide...");

                // Get the starting cell reference for the table
                CellReference startCellRef = table.getStartCellReference();
                int headerRowIndex = startCellRef.getRow() - 1; // Adjust for hidden header
                int startColumnIndex = startCellRef.getCol();

                // Retrieve the table column names from CTTable
                if (ctTable.getTableColumns() != null) {
                    // Use custom method to write header row
                    writeHiddenHeader(sheet, headerRowIndex, ctTable, startColumnIndex, outputFilePath);
                } else {
                    log.info("No columns found in the table.");
                }
            } else {
                log.info("Header row is already visible.");
            }
        }
    }


    private static void getRangesInfo(Workbook workbook,String sheetName){
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
                    log.info("Sheet: {}", firstCell.getSheetName());
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
                        log.info("Sheet-: {}", firstCell.getSheetName());
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

    private static void writeHiddenHeader(XSSFSheet sheet, int headerRowIndex, CTTable ctTable,
                                          int startColumnIndex, String outputFilePath) {
        try {
            // Populate the header row with column names
            Row headerRow = sheet.getRow(headerRowIndex);
            if (headerRow == null) {
                headerRow = sheet.createRow(headerRowIndex);
            }

            int colIndex = startColumnIndex;
            for (CTTableColumn column : ctTable.getTableColumns().getTableColumnList()) {
                Cell cell = headerRow.createCell(colIndex, CellType.STRING);
                cell.setCellValue(column.getName()); // Write the column name into the cell
                colIndex++;
            }

            // Set the row height to default to make it visible
            headerRow.setZeroHeight(false);
            headerRow.setHeight(sheet.getDefaultRowHeight());

            // Write changes to the output file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                sheet.getWorkbook().write(fos);
                log.info("Hidden header row written successfully to file: {}", outputFilePath);
            }
        } catch (IOException e) {
            log.error("Error while writing hidden header row to file.", e);
        }
    }

}