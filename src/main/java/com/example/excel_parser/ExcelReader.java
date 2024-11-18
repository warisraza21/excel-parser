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

                List<AreaReference> reservedAreas = getReservedAreas(xssfSheet, workbook);

                log.info("Processing Unstructured Data");
                List<ExcelProcessor.CellData> list = new ArrayList<>();
                for (Row row : xssfSheet) {
                    for (Cell cell : row) {
                        if (!isWithinReservedArea(cell, reservedAreas)) {
                            ExcelProcessor.CellData cellData = extractCellData(cell);
                            if(cellData != null) list.add(cellData);
                        }
                    }
                }

                if(!list.isEmpty()){
                    ExcelProcessor.ProcessedSheet processedSheet =  ExcelProcessor.processSheet(list);
                    ExcelProcessor.test(processedSheet);
                }
            }
        } catch (IOException e) {
            log.error("Error processing Excel file", e);
        }
    }

    private static List<AreaReference> getReservedAreas(XSSFSheet sheet, Workbook workbook) {
        List<AreaReference> reservedAreas = new ArrayList<>();

        // Process normal tables
        for (XSSFTable table : sheet.getTables()) {
            AreaReference areaReference = table.getArea();
            reservedAreas.add(areaReference);
            log.info("Table: {} added to reserved areas {}", table.getName(), areaReference);
        }

        // Process pivot tables
        for (XSSFPivotTable pivotTable : sheet.getPivotTables()) {
            AreaReference areaReference = new AreaReference(
                    pivotTable.getCTPivotTableDefinition().getLocation().getRef(),
                    workbook.getSpreadsheetVersion()
            );
            reservedAreas.add(areaReference);
            log.info("Pivot Table with parent name {} added to reserved areas {}", pivotTable.getCTPivotTableDefinition().getName(), areaReference);
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

    private static ExcelProcessor.CellData extractCellData(Cell cell) {
        ExcelProcessor.CellData cellData = null;
        if (cell != null && !cell.toString().trim().isEmpty()) {
            cellData = new ExcelProcessor.CellData(cell.getRowIndex(), cell.getColumnIndex());
            cellData.setValue(cell.toString());
            if (cell.getCellType() == CellType.FORMULA) {
                cellData.setFormula(cell.getCellFormula());
            }
        }
        return cellData;
    }
}
