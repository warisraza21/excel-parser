package com.example.excel_parser.service;

import com.example.excel_parser.dtos.*;
import com.aspose.cells.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@Slf4j
public class WorkbookInfoBuilderV2 {

    private static String generateUniqueId() {
        return UUID.randomUUID().toString();
    }

    public static WorkbookDTO buildWorkbookInfo(String filePath) {
        WorkbookDTO workbookDTO = new WorkbookDTO();
        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(filePath);
            Workbook workbook = new Workbook(filePath);
            WorksheetCollection worksheets = workbook.getWorksheets();

            workbookDTO.setId(generateUniqueId());
            workbookDTO.setName(filePath.substring(filePath.lastIndexOf("/") + 1));
            workbookDTO.setExcelVersion(getExcelVersion(workbook.getFileFormat())); // Placeholder for actual version if needed
            workbookDTO.setCreatedAt(null); // Add logic if creation time is available
            workbookDTO.setLastModifiedAt(null); // Add logic if modification time is available
            workbookDTO.setSheetCount(worksheets.getCount());
            workbookDTO.setFileSize("Unknown"); // Placeholder, update with logic to calculate file size
            workbookDTO.setProtected(workbook.isWorkbookProtectedWithPassword());
            workbookDTO.setNamedRanges(extractNamedRanges(workbook));



            List<SheetDTO> sheetDTOList = new ArrayList<>();

            for (int i = 0; i < worksheets.getCount(); i++) {
                Worksheet worksheet = worksheets.get(i);
                List<NonTableInfo> nonTableInfos = WorkBookInfoBuilder.extractNonTables(xssfWorkbook.getSheetAt(i));
                SheetDTO sheetDTO = processSheet(worksheet, i,nonTableInfos);
                sheetDTOList.add(sheetDTO);
            }

            workbookDTO.setSheets(sheetDTOList);

        } catch (Exception e) {
            log.error("Error processing Excel file with Aspose", e);
        }
        return workbookDTO;
    }

    private static SheetDTO processSheet(Worksheet worksheet, int sheetIndex, List<NonTableInfo> nonTableInfoList) {
        SheetDTO sheetDTO = new SheetDTO();
        sheetDTO.setId(generateUniqueId());
        sheetDTO.setName(worksheet.getName());
        sheetDTO.setSheetIndex(sheetIndex);
        sheetDTO.setVisibility(worksheet.isVisible() ? "Visible" : "Hidden");

        sheetDTO.setTabularData(extractTables(worksheet));
        sheetDTO.setPivotTables(extractPivotTables(worksheet));
        sheetDTO.setCharts(extractCharts(worksheet));


        List<MiscellaneousDTO> miscellaneousDTOList = new ArrayList<>();
        for (NonTableInfo nonTableInfo : nonTableInfoList) {
            // Convert CellRangeInfo to BoundariesDTO
            BoundariesDTO boundariesDTO = new BoundariesDTO(
                    nonTableInfo.getBoundaries().getStartCell(),
                    nonTableInfo.getBoundaries().getEndCell()
            );

            MiscellaneousDTO miscellaneousDTO = new MiscellaneousDTO();
            miscellaneousDTO.setBoundaries(boundariesDTO);

            // Convert sample row data to RowDTO, focusing on formula
            List<RowDTO> rows = new ArrayList<>();
            List<List<String>> sampleRowData = nonTableInfo.getSampleRowData();
            if (sampleRowData != null) {
                for (int rowIndex = 0; rowIndex < sampleRowData.size(); rowIndex++) {
                    List<String> rowData = sampleRowData.get(rowIndex);

                    List<CellDTO> cells = new ArrayList<>();
                    for (int columnIndex = 0; columnIndex < rowData.size(); columnIndex++) {
                        // Build CellDTO with formula-specific data
                        CellDTO cellDTO = new CellDTO();
                        cellDTO.setRowIndex(rowIndex);
                        cellDTO.setColumnIndex(columnIndex);
                        cellDTO.setValue(rowData.get(columnIndex));
                        cellDTO.setValueType("String"); // Adjust based on actual type if needed
                        cellDTO.setFormula(createFormulaDTO(nonTableInfo.getColumns().get(columnIndex)));
                        cellDTO.setFormatting(createFormattingDTO(nonTableInfo.getColumns().get(columnIndex).getFormatting()));

                        cells.add(cellDTO);
                    }

                    // Assume the first row is the header
                    RowDTO rowDTO = new RowDTO(rowIndex, rowIndex == 0, null, cells);
                    rows.add(rowDTO);
                }
            }

            miscellaneousDTO.setRows(rows);
            miscellaneousDTOList.add(miscellaneousDTO);
        }

        sheetDTO.setMiscellaneous(miscellaneousDTOList);

        return sheetDTO;
    }



    private static List<TabularDataDTO> extractTables(Worksheet worksheet) {
        List<TabularDataDTO> tabularDataList = new ArrayList<>();
        ListObjectCollection tables = worksheet.getListObjects();

        for (int i = 0; i < tables.getCount(); i++) {
            ListObject table = tables.get(i);

            TabularDataDTO tableDTO = new TabularDataDTO();
            tableDTO.setId(generateUniqueId());
            tableDTO.setName(table.getDisplayName());
            tableDTO.setType("Table");
            tableDTO.setColumnsCount(table.getListColumns().getCount());
            tableDTO.setRowsCount(table.getEndRow() - table.getStartRow() + 1);
            tableDTO.setBoundaries(new BoundariesDTO(
                    CellsHelper.cellIndexToName(table.getStartRow(), table.getStartColumn()),
                    CellsHelper.cellIndexToName(table.getEndRow(), table.getEndColumn())
            ));

            tableDTO.setRows(new ArrayList<>()); // Add row extraction logic if necessary
            tableDTO.setTotalRow(null); // Add total row information if required

            tabularDataList.add(tableDTO);
        }
        return tabularDataList;
    }

    private static List<PivotTableDTO> extractPivotTables(Worksheet worksheet) {
        List<PivotTableDTO> pivotTableList = new ArrayList<>();
        PivotTableCollection pivotTables = worksheet.getPivotTables();

        for (int i = 0; i < pivotTables.getCount(); i++) {
            PivotTable pivotTable = pivotTables.get(i);

            PivotTableDTO pivotTableDTO = new PivotTableDTO();
            pivotTableDTO.setId(generateUniqueId());
            pivotTableDTO.setName(pivotTable.getName());

            CellArea pivotRange = pivotTable.getTableRange1();
            pivotTableDTO.setBoundaries(new BoundariesDTO(
                    CellsHelper.cellIndexToName(pivotRange.StartRow, pivotRange.StartColumn),
                    CellsHelper.cellIndexToName(pivotRange.EndRow, pivotRange.EndColumn)
            ));


            // Set Pivot Data Source
            Object dataSource = pivotTable.getDataSource();
            if (dataSource != null) {
                String[] sourceArray = (String[]) dataSource;
                if (sourceArray.length > 0) {
                    pivotTableDTO.setSources(processSourceData(worksheet.getWorkbook().getWorksheets(), sourceArray[0]));
                }
            }

            // Fields
            FieldsDTO fieldsDTO = new FieldsDTO();
            fieldsDTO.setRowFields(extractPivotFields(pivotTable.getRowFields()));
            fieldsDTO.setColumnFields(extractPivotFields(pivotTable.getColumnFields()));
            fieldsDTO.setValueFields(extractValueFields(pivotTable.getDataFields()));
            fieldsDTO.setFilterFields(extractFilterFields(pivotTable.getPageFields()));
            pivotTableDTO.setFields(fieldsDTO);

            // Options
            PivotOptionsDTO optionsDTO = new PivotOptionsDTO();
            GrandTotalsDTO grandTotalsDTO = new GrandTotalsDTO();
            grandTotalsDTO.setRows(pivotTable.getRowGrand());
            grandTotalsDTO.setColumns(pivotTable.getColumnGrand());
            optionsDTO.setGrandTotals(grandTotalsDTO);
            pivotTableDTO.setOptions(optionsDTO);

            pivotTableList.add(pivotTableDTO);
        }
        return pivotTableList;
    }

    private static List<FieldDTO> extractPivotFields(PivotFieldCollection pivotFields) {
        List<FieldDTO> fieldList = new ArrayList<>();
        for (int i = 0; i < pivotFields.getCount(); i++) {
            PivotField field = pivotFields.get(i);
            FieldDTO fieldDTO = new FieldDTO();
            fieldDTO.setName(field.getName());
            fieldDTO.setSortOrder("Ascending"); // Add logic to fetch sort order if available
            fieldList.add(fieldDTO);
        }
        return fieldList;
    }

    private static List<ValueFieldDTO> extractValueFields(PivotFieldCollection pivotFields) {
        List<ValueFieldDTO> valueFieldList = new ArrayList<>();
        for (int i = 0; i < pivotFields.getCount(); i++) {
            PivotField field = pivotFields.get(i);

            // Extract the aggregation function dynamically if available
            String function = getAggregationFunction(field);

            // Retrieve format (if available)
            String format = field.getNumberFormat() != null ? field.getNumberFormat() : "General";

            // Create DTO
            ValueFieldDTO valueFieldDTO = new ValueFieldDTO();
            valueFieldDTO.setName(field.getName());
            valueFieldDTO.setFunction(function);
            valueFieldDTO.setFormat(format);

            valueFieldList.add(valueFieldDTO);
        }
        return valueFieldList;
    }

    private static String getAggregationFunction(PivotField field) {
        switch (field.getFunction()) {
            case ConsolidationFunction.SUM:
                return "Sum";
            case ConsolidationFunction.AVERAGE:
                return "Average";
            case ConsolidationFunction.COUNT:
                return "Count";
            case ConsolidationFunction.MIN:
                return "Min";
            case ConsolidationFunction.MAX:
                return "Max";
            case ConsolidationFunction.PRODUCT:
                return "Product";
            default:
                return "Unknown"; // Default if the function is unrecognized
        }
    }

    private static List<FilterFieldDTO> extractFilterFields(PivotFieldCollection pivotFields) {
        List<FilterFieldDTO> filterFieldList = new ArrayList<>();
        for (int i = 0; i < pivotFields.getCount(); i++) {
            PivotField field = pivotFields.get(i);

            // Extract criteria (selected items) for the filter
            List<String> criteria = new ArrayList<>();
            if (field.getItems() != null) {
                for (int j = 0; j < field.getItems().length; j++) {
                    Object item = field.getItems()[j];
                    if (item != null) {
                        criteria.add(item.toString());
                    }
                }
            }

            // Create FilterFieldDTO object
            FilterFieldDTO filterFieldDTO = new FilterFieldDTO();
            filterFieldDTO.setName(field.getName());
            filterFieldDTO.setCriteria(criteria);

            filterFieldList.add(filterFieldDTO);
        }
        return filterFieldList;
    }



    private static List<ChartDTO> extractCharts(Worksheet worksheet) {
        List<ChartDTO> chartList = new ArrayList<>();
        ChartCollection charts = worksheet.getCharts();

        for (int i = 0; i < charts.getCount(); i++) {
            Chart chart = charts.get(i);

            ChartDTO chartDTO = new ChartDTO();
            chartDTO.setId(generateUniqueId());
            chartDTO.setName(chart.getTitle() != null ? chart.getTitle().getText() : "Untitled");
            chartDTO.setType(getChartTypeName(chart.getType()));
            chartDTO.setBoundaries(new BoundariesDTO(
                    chart.getChartObject().getUpperLeftRow() + "," + chart.getChartObject().getUpperLeftColumn(),
                    chart.getChartObject().getLowerRightRow() + "," + chart.getChartObject().getLowerRightColumn()
            ));

            // Legends
            List<LegendDTO> legendList = new ArrayList<>();
            SeriesCollection seriesCollection = chart.getNSeries();
            for (int j = 0; j < seriesCollection.getCount(); j++) {
                Series series = seriesCollection.get(j);
                LegendDTO legendDTO = new LegendDTO();
                legendDTO.setName(series.getName());
                legendDTO.setColor(series.getArea().getForegroundColor().toString());
                legendList.add(legendDTO);
            }
            chartDTO.setLegends(legendList);

            // Title and Axes
            ChartTitleDTO chartTitleDTO = new ChartTitleDTO();
            chartTitleDTO.setText(chart.getTitle() != null ? chart.getTitle().getText() : "Chart Title");
            chartTitleDTO.setFontSize(12);
            chartTitleDTO.setBold(true);
            chartDTO.setTitle(chartTitleDTO);

            AxesDTO axesDTO = new AxesDTO();
            axesDTO.setXAxis(new AxisDTO("Category", "String"));
            axesDTO.setYAxis(new AxisDTO("Values", "Numeric"));
            chartDTO.setAxes(axesDTO);

            chartList.add(chartDTO);
        }
        return chartList;
    }

    private static String getChartTypeName(int chartType) {
        switch (chartType) {
            case ChartType.COLUMN:
                return "Column";
            case ChartType.BAR:
                return "Bar";
            case ChartType.LINE:
                return "Line";
            case ChartType.PIE:
                return "Pie";
            case ChartType.AREA:
                return "Area";
            case ChartType.SCATTER:
                return "Scatter";
            case ChartType.DOUGHNUT:
                return "Doughnut";
            case ChartType.RADAR:
                return "Radar";
            case ChartType.SURFACE_3_D,ChartType.SURFACE_CONTOUR,ChartType.SURFACE_CONTOUR_WIREFRAME,ChartType.SURFACE_WIREFRAME_3_D:
                return "Surface";
            case ChartType.BUBBLE:
                return "Bubble";
            default:
                return "Unknown";
        }
    }


    private static List<NamedRangeDTO> extractNamedRanges(Workbook workbook) {
        List<NamedRangeDTO> namedRanges = new ArrayList<>();
        NameCollection names = workbook.getWorksheets().getNames();

        for (int i = 0; i < names.getCount(); i++) {
            Name name = names.get(i);
            NamedRangeDTO namedRangeDTO = new NamedRangeDTO();
            namedRangeDTO.setName(name.getText());

            String refersTo = name.getRefersTo();
            if (refersTo == null || refersTo.isEmpty()) {
                namedRangeDTO.setBoundary("Invalid or Undefined");
                namedRangeDTO.setSheetName("Unknown");
                namedRangeDTO.setRange("Unknown");
                namedRangeDTO.setCreatedFromTable(false);
            } else {
                namedRangeDTO.setBoundary(refersTo);

                // Extract sheet name and range
                if (refersTo.contains("!")) {
                    String sheetName = extractSheetName(refersTo);
                    String range = extractTableRange(refersTo);

                    namedRangeDTO.setSheetName(sheetName);
                    namedRangeDTO.setRange(range);

                    // Check if the range is part of a table
                    Worksheet sheet = workbook.getWorksheets().get(sheetName);
                    if (sheet != null) {
                        namedRangeDTO.setCreatedFromTable(isRangePartOfTable(sheet, range));
                    } else {
                        namedRangeDTO.setCreatedFromTable(false); // Sheet not found
                    }
                } else {
                    namedRangeDTO.setSheetName("Unknown");
                    namedRangeDTO.setRange(refersTo); // Could be a workbook-level reference
                    namedRangeDTO.setCreatedFromTable(false);
                }
            }

            namedRanges.add(namedRangeDTO);
        }

        return namedRanges;
    }



    private static String getExcelVersion(int fileFormatType) {
        switch (fileFormatType) {
            case FileFormatType.XLSX:
                return "Excel 2007 or later (XLSX)";
            case FileFormatType.XLSM:
                return "Excel 2007 or later with Macros (XLSM)";
            case FileFormatType.XLTX:
                return "Excel 2007 or later Template (XLTX)";
            case FileFormatType.XLT:
                return "Excel 97-2003 Template (XLT)";
            case FileFormatType.CSV:
                return "CSV File (Version not applicable)";
            default:
                return "Unknown Format";
        }
    }

    private static FormulaDTO createFormulaDTO(ColumnInfo columnInfo) {
        if (columnInfo.getFormula() == null) {
            return null;
        }

        // Create the FormulaDTO
        FormulaDTO formulaDTO = new FormulaDTO();
        formulaDTO.setStructured(true);
        formulaDTO.setExpression(columnInfo.getFormula().toString());

        // Extract operands dynamically
        List<OperandDTO> operands = new ArrayList<>();
        // Assuming the formula operands can be deduced dynamically from columnInfo
        if (columnInfo.getFormula() instanceof List<?>) {
            List<?> operandDetails = (List<?>) columnInfo.getFormula();
            for (Object operandDetail : operandDetails) {
                if (operandDetail instanceof OperandInfo) { // Assuming OperandInfo contains operand details
                    OperandInfo operandInfo = (OperandInfo) operandDetail;
                    operands.add(new OperandDTO(
                            operandInfo.getSource(),
                            operandInfo.getSourceType(),
                            operandInfo.getRowIndex(),
                            operandInfo.getColumnIndex(),
                            operandInfo.getColumnName()
                    ));
                }
            }
        }

        formulaDTO.setOperands(operands);
        return formulaDTO;
    }


    private static FormattingDTO createFormattingDTO(CellFormatting formatting) {
        if (formatting == null) {
            return null;
        }

        return new FormattingDTO(
                formatting.getFontWeight(),
                formatting.getFontColor(),
                formatting.getBackgroundColor()
        );
    }

    private static SourceDTO processSourceData(WorksheetCollection worksheets, String source) {
        SourceDTO sourceDTO = new SourceDTO();

        if (source.contains("!")) {
            // Extract sheet name and range
            String sheetName = extractSheetName(source);
            String range = extractTableRange(source);

            sourceDTO.setSheetName(sheetName);
            sourceDTO.setSource(source);
            sourceDTO.setSourceType("Range");

            if (!sheetName.equals("Unknown") && !range.equals("Unknown")) {
                Worksheet sourceSheet = worksheets.get(sheetName);
                if (sourceSheet != null) {
                    BoundariesDTO boundaries = addRangeInfo(sourceSheet, range);
                    sourceDTO.setBoundaries(boundaries);
                } else {
                    throw new IllegalArgumentException("Sheet not found: " + sheetName);
                }
            } else {
                throw new IllegalArgumentException("Invalid sheet name or range: " + source);
            }
        } else {
            sourceDTO.setSource(source);
            sourceDTO.setSourceType("Table");
            String sheetName = findSheetByTable(worksheets, source);
            sourceDTO.setSheetName(sheetName);

            if (!sheetName.equals("Unknown")) {
                Worksheet sourceSheet = worksheets.get(sheetName);
                if (sourceSheet != null) {
                    ListObject table = sourceSheet.getListObjects().get(source);
                    if (table != null) {
                        String startCell = CellsHelper.cellIndexToName(table.getStartRow(), table.getStartColumn());
                        String endCell = CellsHelper.cellIndexToName(table.getEndRow(), table.getEndColumn());
                        sourceDTO.setBoundaries(new BoundariesDTO(startCell, endCell));
                    } else {
                        throw new IllegalArgumentException("Table not found: " + source);
                    }
                } else {
                    throw new IllegalArgumentException("Sheet not found for table: " + source);
                }
            } else {
                throw new IllegalArgumentException("Table does not belong to any sheet: " + source);
            }
        }

        return sourceDTO;
    }

    private static String extractSheetName(String source) {
        if (source.contains("!")) {
            return source.split("!")[0].replace("'", ""); // Remove single quotes
        }
        return "Unknown";
    }

    private static String extractTableRange(String source) {
        if (source.contains("!")) {
            return source.split("!")[1]; // Extract the range after '!'
        }
        return "Unknown";
    }

    private static String findSheetByTable(WorksheetCollection worksheets, String tableName) {
        for (int i = 0; i < worksheets.getCount(); i++) {
            Worksheet sheet = worksheets.get(i);
            ListObjectCollection tables = sheet.getListObjects();
            for (int j = 0; j < tables.getCount(); j++) {
                if (tables.get(j).getDisplayName().equals(tableName)) {
                    return sheet.getName(); // Return the parent sheet name
                }
            }
        }
        return "Unknown";
    }

    private static BoundariesDTO addRangeInfo(Worksheet sheet, String range) {
        Range rangeObj = sheet.getCells().createRange(range);

        // Create boundaries DTO
        String startCell = CellsHelper.cellIndexToName(rangeObj.getFirstRow(), rangeObj.getFirstColumn());
        String endCell = CellsHelper.cellIndexToName(rangeObj.getFirstRow() + rangeObj.getRowCount() - 1,
                rangeObj.getFirstColumn() + rangeObj.getColumnCount() - 1);
        return new BoundariesDTO(startCell, endCell);
    }

    private static boolean isRangePartOfTable(Worksheet sheet, String range) {
        ListObjectCollection tables = sheet.getListObjects();
        Range rangeObj = sheet.getCells().createRange(range);

        for (int i = 0; i < tables.getCount(); i++) {
            ListObject table = tables.get(i);
            Range tableRange = sheet.getCells().createRange(
                    table.getStartRow(),
                    table.getStartColumn(),
                    table.getEndRow() - table.getStartRow() + 1,
                    table.getEndColumn() - table.getStartColumn() + 1
            );

            if (rangeObj.getFirstRow() >= tableRange.getFirstRow() &&
                    rangeObj.getFirstColumn() >= tableRange.getFirstColumn() &&
                    rangeObj.getFirstRow() + rangeObj.getRowCount() <= tableRange.getFirstRow() + tableRange.getRowCount() &&
                    rangeObj.getFirstColumn() + rangeObj.getColumnCount() <= tableRange.getFirstColumn() + tableRange.getColumnCount()) {
                return true;
            }
        }
        return false;
    }




}
