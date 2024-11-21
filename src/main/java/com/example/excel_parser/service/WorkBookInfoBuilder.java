package com.example.excel_parser.service;

import com.example.excel_parser.dtos.*;
import com.example.excel_parser.model.CellData;
import com.example.excel_parser.model.ProcessedSheet;
import com.example.excel_parser.model.TableData;
import jdk.jshell.spi.SPIResolutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

@Slf4j
public class WorkBookInfoBuilder {

    public static FormulaEvaluator formulaEvaluator;

    private static String generateUniqueId() {
        return UUID.randomUUID().toString();
    }

    public static WorkBookInfo buildWorkBookInfo(String filePath){

        WorkBookInfo workBookInfo = null;
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

//            FileChannel fileChannel = fis.getChannel();
//            long fileSizeInBytes = fileChannel.size();
//
//            // Convert to KB and MB for readability
//            double fileSizeInKB = fileSizeInBytes / 1024.0;
            POIXMLProperties properties = workbook.getProperties();
            POIXMLProperties.CoreProperties coreProperties = properties.getCoreProperties();

            // Retrieve the last modified time and last modifier
            Date lastModifiedDate = coreProperties.getModified();
            Date createdAt = coreProperties.getCreated();

            log.info("Total sheets: {}", workbook.getNumberOfSheets());

            workBookInfo = new WorkBookInfo();
            workBookInfo.setId(UUID.randomUUID().toString());
            workBookInfo.setName(filePath.substring(filePath.lastIndexOf("/") + 1));
            workBookInfo.setExcelVersion(workbook.getSpreadsheetVersion().toString());
            workBookInfo.setCreatedAt(createdAt);
            workBookInfo.setLastModifiedAt(lastModifiedDate);
            workBookInfo.setSheetCount(workbook.getNumberOfSheets());
            workBookInfo.setFileSize(String.valueOf("fileSizeInKB"));
//            workBookInfo.setProtected(workbook.isHidden());

            List<SheetInfo> sheetInfoList = new ArrayList<>();

            int index = 0;
            for (Sheet sheet : workbook) {

                XSSFSheet xssfSheet = (XSSFSheet) sheet;
                SheetInfo sheetInfo = processSheet(xssfSheet, index);

                sheetInfo.setSheetIndex(index);

                log.info("SheetInfo Name : {}", xssfSheet.getSheetName());


                sheetInfoList.add(sheetInfo);
                index++;
            }

            workBookInfo.setSheets(sheetInfoList);
        } catch (IOException e) {
            log.error("Error processing Excel file", e);
        }
        return workBookInfo;
    }

    private static SheetInfo processSheet(XSSFSheet sheet, int sheetIndex) {
        SheetInfo sheetInfo = new SheetInfo();
        sheetInfo.setId(generateUniqueId());
        sheetInfo.setName(sheet.getSheetName());
        sheetInfo.setSheetIndex(sheetIndex);
        sheetInfo.setVisibility(!sheet.isSheetLocked());

        // Extract tables
        List<TableInfo> tables = extractTables(sheet);
        sheetInfo.setTables(tables);
        sheetInfo.setTableCounts(tables.size());

        // Extract pivot tables
        List<PivotTableInfo> pivotTables = extractPivotTables(sheet);
        sheetInfo.setPivotTables(pivotTables);
        sheetInfo.setPivotTableCount(pivotTables.size());

        // Extract charts
        List<ChartInfo> charts = extractCharts(sheet);
        sheetInfo.setCharts(charts);
        sheetInfo.setChartCounts(charts.size());

        // Extract named ranges
        List<NamedRangeInfo> namedRanges = extractNamedRanges(sheet);
        sheetInfo.setNamedRanges(namedRanges);

        List<NonTableInfo> nonTables = extractNonTables(sheet);
        sheetInfo.setNonTables(nonTables);
        sheetInfo.setNonTableCount(nonTables.size());

        return sheetInfo;
    }

    private static List<TableInfo> extractTables(XSSFSheet sheet) {
        List<TableInfo> tables = new ArrayList<>();

            sheet.getTables().forEach(table -> {
                TableInfo tableInfo = createTableInfoObject(table);
                tableInfo.setBoundaries(getCellRangeInfo(table.getArea()));
                tableInfo.setColumns(getTableColumns(table));
                tables.add(tableInfo);
            });

        return tables;
    }

    private static List<NonTableInfo> extractNonTables(XSSFSheet sheet) {

        List<AreaReference> reservedAreas = getReservedAreas(sheet,sheet.getWorkbook().getSpreadsheetVersion());

        List<CellData> list = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (!isWithinReservedArea(cell, reservedAreas)) {
                    CellData cellData = extractCellData(cell);
                    if(cellData != null) list.add(cellData);
                }
            }
        }


        List<NonTableInfo> nonTables = new ArrayList<>();
        if(!list.isEmpty()){
            ProcessedSheet processedSheet =  ClusterCells.clusterCellsData(list);

            processedSheet.tableData().forEach(table -> {
                NonTableInfo nonTableInfo = createNonTableInfo(table);
                nonTableInfo.setColumns(getNonTableColumns(table));
                nonTables.add(nonTableInfo);
            });
        }

        return nonTables;
    }


    private static List<ChartInfo> extractCharts(XSSFSheet sheet) {
        List<ChartInfo> charts = new ArrayList<>();
        // Add chart extraction logic using Apache POI
        return charts;
    }

    private static List<PivotTableInfo> extractPivotTables(XSSFSheet sheet) {
        List<PivotTableInfo> pivotTables = new ArrayList<>();
        SpreadsheetVersion spreadsheetVersion = sheet.getWorkbook().getSpreadsheetVersion();
        sheet.getPivotTables().forEach(pivotTable -> {
            PivotTableInfo pivotTableInfo = createPivotTableInfoObject(pivotTable,spreadsheetVersion);
            pivotTables.add(pivotTableInfo);
        });
        return pivotTables;
    }

    private static List<NamedRangeInfo> extractNamedRanges(XSSFSheet sheet) {
        List<NamedRangeInfo> namedRanges = new ArrayList<>();
        // Add named range extraction logic if available
        return namedRanges;
    }

    private static String getFileSize(String filePath) {
        // Logic to calculate file size
        return "1.2 MB"; // Placeholder
    }




    private static TableInfo createTableInfoObject(XSSFTable table){
        TableInfo tableInfo = new TableInfo();
        tableInfo.setId(UUID.randomUUID().toString());
        tableInfo.setName(table.getName());
        tableInfo.setRowCount(table.getRowCount());
        tableInfo.setColumnCount(table.getColumnCount());
        return tableInfo;
    }

    private static NonTableInfo createNonTableInfo(TableData table) {
        NonTableInfo nonTableInfo = new NonTableInfo();
        nonTableInfo.setId(UUID.randomUUID().toString());
        nonTableInfo.setRowCount(table.getRowCount());
        nonTableInfo.setColumnCount(table.getColumnCount());
        nonTableInfo.setBoundaries(table.getBoundaries());
        return nonTableInfo;
    }


    private static PivotTableInfo createPivotTableInfoObject(XSSFPivotTable pivotTable, SpreadsheetVersion spreadsheetVersion) {
        PivotTableInfo pivotTableInfo = new PivotTableInfo();
        pivotTableInfo.setId(UUID.randomUUID().toString());
        pivotTableInfo.setName(pivotTable.getCTPivotTableDefinition().getName());

        AreaReference areaReference = new AreaReference(
                pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                spreadsheetVersion
        );

        CellRangeInfo cellRangeInfo = new CellRangeInfo(
                areaReference.getFirstCell().formatAsString(),
                areaReference.getLastCell().formatAsString()
        );
        pivotTableInfo.setBoundaries(cellRangeInfo);

        // Fetch the source and parent information
        pivotTableInfo.setSource(getPivotTableSource(pivotTable));
        return pivotTableInfo;
    }

    public static CellRangeInfo getCellRangeInfo(AreaReference areaReference) {
        CellReference startCell = areaReference.getFirstCell();
        CellReference endCell = areaReference.getLastCell();
        return new CellRangeInfo(startCell.formatAsString(), endCell.formatAsString());
    }

    public static List<ColumnInfo> getTableColumns(XSSFTable table) {
        List<ColumnInfo> columns = new ArrayList<>();

        if (table != null) {
            // Get the reference for the header row
            CellReference startCellRef = table.getStartCellReference();
            int headerRowIndex = startCellRef.getRow();

            // Get the header row
            Sheet sheet = table.getXSSFSheet();
            Row headerRow = sheet.getRow(headerRowIndex);

            if (headerRow != null) {
                // Iterate through the columns in the table range
                int startCol = startCellRef.getCol();
                int endCol = table.getEndCellReference().getCol();

                for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                    Cell headerCell = headerRow.getCell(colIndex);

                    ColumnInfo columnInfo = new ColumnInfo();
                    columnInfo.setName(headerCell != null ? headerCell.getStringCellValue() : "");
                    columnInfo.setType(detectColumnType(sheet, colIndex, headerRowIndex + 1)); // Infer type from data
                    columnInfo.setDerived(false); // Update this logic based on your requirements
                    columnInfo.setFormula(getColumnFormula(sheet, colIndex));
                    columnInfo.setDataValidation(getDataValidation(sheet, colIndex));
                    columnInfo.setFormatting(extractCellFormatting(headerCell));

                    columns.add(columnInfo);
                }
            }
        }
        return columns;
    }

    private static List<ColumnInfo> getNonTableColumns(TableData table) {
        List<ColumnInfo> columnInfos = new ArrayList<>();

        // Extract the boundaries to identify the first row
        CellRangeInfo boundaries = table.getBoundaries();
        String startCell = boundaries.getStartCell();
        String endCell = boundaries.getEndCell();

        // Parse the start and end cell references
        CellReference startRef = new CellReference(startCell);
        CellReference endRef = new CellReference(endCell);

        // Get the first row index
        int firstRow = startRef.getRow();

        // Iterate over the cells in the first row (header row)
        for (CellData cell : table.getCells()) {
            if (cell.getRowIndex() == firstRow) {
                ColumnInfo columnInfo = new ColumnInfo();
                columnInfo.setName(String.valueOf(cell.getValue())); // Header name
                columnInfo.setType("String"); // Default type, can be enhanced
                columnInfo.setDerived(false); // Assume non-derived unless specified
                columnInfo.setFormula(null);

                columnInfos.add(columnInfo);
            }
        }

        return columnInfos;
    }


    private static String detectColumnType(Sheet sheet, int colIndex, int startRow) {
        // Inspect data rows in the column to infer the type
        for (int rowIndex = startRow; rowIndex < sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            return "String";
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return "Date";
                            }
                            return "Numeric";
                        case BOOLEAN:
                            return "Boolean";
                        case FORMULA:
                            return "Formula";
                        default:
                            break;
                    }
                }
            }
        }
        return "Unknown";
    }

    private static Object getColumnFormula(Sheet sheet, int colIndex) {
        // Placeholder: Logic to retrieve formula applied to a column if relevant
        return null;
    }

    private static DataValidationRules getDataValidation(Sheet sheet, int colIndex) {
        DataValidationRules rules = new DataValidationRules();
        // Placeholder: Extract data validation rules if available
        // Example: Retrieve validation from sheet.getDataValidationHelper()
        return rules;
    }

    public static CellFormatting extractCellFormatting(Cell cell) {
        CellFormatting formatting = new CellFormatting();

        if (cell != null) {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle style = cell.getCellStyle();

            if (style != null) {
                // Extract font weight
                XSSFFont font = (XSSFFont) workbook.getFontAt(style.getFontIndexAsInt());
                formatting.setFontWeight(font.getBold() ? "Bold" : "Normal");

                    XSSFColor fontColor = font.getXSSFColor();
                    formatting.setFontColor(fontColor != null ? getHexColor(fontColor) : "#000000");

                    XSSFCellStyle xssfCellStyle = (XSSFCellStyle) style;
                    XSSFColor backgroundColor = xssfCellStyle.getFillForegroundXSSFColor();
                    formatting.setBackgroundColor(backgroundColor != null ? getHexColor(backgroundColor) : "#FFFFFF");

            }
        }

        return formatting;
    }

    private static PivotTableSource getPivotTableSource(XSSFPivotTable pivotTable) {
        PivotTableSource pivotTableSource = new PivotTableSource();
        CTPivotCacheDefinition pivotCacheDef = pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition();

        // Extract source table or range
        if (pivotCacheDef.getCacheSource() != null && pivotCacheDef.getCacheSource().isSetWorksheetSource()) {
            CTCacheSource cacheSource = pivotCacheDef.getCacheSource();
            CTWorksheetSource worksheetSource = cacheSource.getWorksheetSource();

            String tableName = worksheetSource.isSetName() ? worksheetSource.getName() : null;
            String sheetName = worksheetSource.isSetSheet() ? worksheetSource.getSheet() : null;

            if (tableName != null) {
                pivotTableSource.setTableName(tableName);
                log.info("Pivot table source table: " + tableName);
            } else if (worksheetSource.isSetRef()) {
                String range = worksheetSource.getRef();
                pivotTableSource.setTableName("No table name (range: " + range + ")");
                log.info("Pivot table source range: " + range);
            }

            if (sheetName != null) {
                pivotTableSource.setTableSheetName(sheetName);
                log.info("Pivot table source sheet: " + sheetName);
            } else {
                pivotTableSource.setTableSheetName("No sheet name found");
            }
        } else {
            pivotTableSource.setTableName("No table or range found");
            pivotTableSource.setTableSheetName("No sheet name found");
            log.info("Cache source is not set or invalid for this pivot table.");
        }

        // Log parent sheet and table information
        String parentSheetName = getParentSheetName(pivotCacheDef);
        pivotTableSource.setTableSheetName(parentSheetName);
        log.info("Parent sheet name: " + parentSheetName);

        return pivotTableSource;
    }

    private static List<AreaReference> getReservedAreas(XSSFSheet sheet, SpreadsheetVersion spreadsheetVersion) {

        List<AreaReference> reservedAreas = new ArrayList<>();

        // Process normal tableData
        for (XSSFTable table : sheet.getTables()) {
            AreaReference areaReference = table.getArea();
            reservedAreas.add(areaReference);
            log.info("Table: {} added to reserved areas [{}:{}]", table.getName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
        }

        // Process pivot tableData
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            AreaReference areaReference = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    spreadsheetVersion
            );
            reservedAreas.add(areaReference);
            log.info("Pivot Table of SheetInfo : {} added to reserved areas [{}:{}]", sheet.getSheetName(), areaReference.getFirstCell().formatAsString(), areaReference.getLastCell().formatAsString());
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
            if (cell.getCellType() == CellType.FORMULA) {
                cellData.setFormula(cell.getCellFormula());
                cellData.setValue(getCellValue(formulaEvaluator.evaluateInCell(cell)));
            } else {
                cellData.setValue(getCellValue(cell));
            }
        }
        return cellData;
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

    private static String getParentSheetName(CTPivotCacheDefinition pivotCacheDef) {
        if (pivotCacheDef.getCacheSource() != null && pivotCacheDef.getCacheSource().isSetWorksheetSource()) {
            CTWorksheetSource worksheetSource = pivotCacheDef.getCacheSource().getWorksheetSource();

            if (worksheetSource.isSetSheet()) {
                return worksheetSource.getSheet();
            }
        }
        return "No parent sheet found";
    }

    private static String getHexColor(XSSFColor color) {
        byte[] rgb = color.getRGB();
        if (rgb != null) {
            return String.format("#%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
        }
        return "#000000"; // Default to black if color is not available
    }
}
