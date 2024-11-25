package com.example.excel_parser.utils;

import com.example.excel_parser.model.CellData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DataTypeUtils {
    public static boolean checkDataTypeRow(String input) {
        // Check all data types in sequence
        if (isInteger(input)) return true;
        if (isLong(input)) return true;
        if (isFloat(input)) return true;
        if (isDouble(input)) return true;
        if (isBigDecimal(input)) return true;
        if (isBoolean(input)) return true;
        if (isCharacter(input)) return true;
        return isDate(input);
    }

    public static String detectDataType(String input) {
        if (isInteger(input) || isLong(input) || isFloat(input) || isDouble(input) || isBigDecimal(input)) return "Number";
        if (isBoolean(input)) return "Boolean";
        if (isDate(input)) return "Date";
        return "String";
    }

    public static boolean isInteger(String input) {
        try {
            Integer.parseInt(input);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static boolean isLong(String input) {
        try {
            Long.parseLong(input);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static boolean isFloat(String input) {
        try {
            Float.parseFloat(input);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static boolean isDouble(String input) {
        try {
            Double.parseDouble(input);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static boolean isBigDecimal(String input) {
        try {
            new BigDecimal(input);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static boolean isBoolean(String input) {
        return input.equalsIgnoreCase("true") || input.equalsIgnoreCase("false");
    }

    public static boolean isCharacter(String input) {
        return input != null && input.length() == 1;
    }

    public static boolean isDate(String input) {
        String[] dateFormats = {
                "yyyy-MM-dd", // Example: 2024-11-24
                "MM/dd/yyyy", // Example: 11/24/2024
                "dd-MM-yyyy", // Example: 24-11-2024
                "yyyy/MM/dd", // Example: 2024/11/24
                "MM-dd-yyyy", // Example: 11-24-2024
                "dd/MM/yyyy", // Example: 24/11/2024
                "yyyy.MM.dd", // Example: 2024.11.24
                "MM.dd.yyyy"  // Example: 11.24.2024
        };

        for (String format : dateFormats) {
            if (tryParseDate(input, format)) {
                return true;
            }
        }
        return false;
    }

    public static boolean tryParseDate(String input, String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        sdf.setLenient(false); // Prevents invalid dates like "2024-02-30"
        try {
            Date date = sdf.parse(input);
            return date != null;
        } catch (ParseException e) {
            return false;
        }
    }

    public static String getCellValue(Cell cell) {
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

    public static void setCellDataType(Cell cell, CellData cellData) {
        switch (cell.getCellType()) {
            case STRING:
                cellData.setDataType(CellType.STRING.name());
                break;
            case NUMERIC:
                cellData.setDataType(CellType.NUMERIC.name());
                break;
            case BOOLEAN:
                cellData.setDataType(CellType.BOOLEAN.name());
                break;
            case FORMULA:
                cellData.setDataType(CellType.FORMULA.name());
                break;
            default:
                cellData.setDataType(CellType._NONE.name());
                break;
        }
    }

    // Utility method to retrieve cell values from the sheet
    public static String getCellValue(XSSFSheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return null;
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
