package com.example.excel_parser.service;

import com.example.excel_parser.utils.DataTypeUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

@Slf4j
public class RectangleDetector {

    public static Set<AreaReference> getRangesAreaReference(XSSFSheet sheet, boolean[][] visited,AreaReference areaReference) {
        Set<AreaReference> areaReferences = new HashSet<>();

        List<int[][]> rectangles = getRangesBoundary(sheet,visited,areaReference);

        for (int[][] rectangle : rectangles) {
            // corners point
            int topLeftRow = rectangle[0][0],bottomRightRow = rectangle[1][0];
            int topLeftCol = rectangle[0][1],bottomRightCol = rectangle[1][1];

            // Create CellReferences
            CellReference topLeft = new CellReference(topLeftRow, topLeftCol);
            CellReference bottomRight = new CellReference(bottomRightRow, bottomRightCol);

            // Create AreaReference
            AreaReference newArea = new AreaReference(topLeft, bottomRight, sheet.getWorkbook().getSpreadsheetVersion());
            areaReferences.add(newArea);
        }

        return areaReferences;
    }

    public static List<int[][]> getRangesBoundary(XSSFSheet sheet,boolean[][] visited, AreaReference areaReference) {
        CellReference firstCell = areaReference.getFirstCell();
        CellReference lastCell = areaReference.getLastCell();

        int startRowIdx = firstCell.getRow();
        int endRowIdx = lastCell.getRow();
        int startColIdx = firstCell.getCol();
        int endColIdx = lastCell.getCol();

        List<int[][]> rectangles = new ArrayList<>();

        // Traverse each row
        for (int rowIdx = startRowIdx; rowIdx <= endRowIdx; rowIdx++) {
            int startColIdxRect = -1;

            // Detect continuous non-empty cells in each row
            for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
                String cellValue = DataTypeUtils.getCellValue(sheet, rowIdx, colIdx);
                if (cellValue != null && !cellValue.isEmpty() && !visited[rowIdx - startRowIdx][colIdx - startColIdx]) {
                    if (startColIdxRect == -1) {
                        startColIdxRect = colIdx; // Start of a potential rectangle block
                    }
                } else if (startColIdxRect != -1 && !visited[rowIdx - startRowIdx][colIdx - startColIdx]) {
                    // End of a continuous block in this row
                    rectangles.addAll(detectVerticalRectangles(sheet, visited, rowIdx, startColIdxRect, colIdx - 1, endRowIdx, startRowIdx, startColIdx));
                    startColIdxRect = -1;
                }
            }
            // Handle case when the row ends with a continuous block
            if (startColIdxRect != -1) {
                rectangles.addAll(detectVerticalRectangles(sheet, visited, rowIdx, startColIdxRect, endColIdx, endRowIdx, startRowIdx, startColIdx));
            }
        }

        return rectangles;
    }

    // Detect vertical rectangles based on continuous blocks
    private static List<int[][]> detectVerticalRectangles(XSSFSheet sheet, boolean[][] visited, int startRowIdx, int startColIdx, int endColIdx, int globalEndRowIdx, int globalStartRowIdx, int globalStartColIdx) {
        List<int[][]> rectangles = new ArrayList<>();
        int currRowIdx = startRowIdx + 1;

        String[] prevRowValue = null;
        // Traverse vertically from the current row downwards
        while (currRowIdx <= globalEndRowIdx) {
            if (prevRowValue == null) prevRowValue = getRowData(sheet, currRowIdx, startColIdx, endColIdx,visited);
            else {
                String[] currRowValue = getRowData(sheet, currRowIdx, startColIdx, endColIdx,visited);
                if (currRowValue != null) {
                    if (!compareRowDataTypes(prevRowValue, currRowValue)) {
                        break;
                    }
                    prevRowValue = currRowValue;
                }else break;
            }
            currRowIdx++;
        }

        // If a valid rectangle is found, mark the cells as visited and add the rectangle
        if (currRowIdx > startRowIdx) {
            for (int rowIdx = startRowIdx; rowIdx < currRowIdx; rowIdx++) {
                for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
                    visited[rowIdx - globalStartRowIdx][colIdx - globalStartColIdx] = true;
                }
            }
            int[][] boundary = new int[2][2];

            int lastRowIdx = narrowDownLastRowBoundary(sheet,startRowIdx,startColIdx,endColIdx,currRowIdx - 1);

            // get first corner
            boundary[0][0] = startRowIdx;
            boundary[0][1] = startColIdx;

            //get last corner
            boundary[1][0] = lastRowIdx;
            boundary[1][1] = endColIdx;

            rectangles.add(boundary);
        }

        return rectangles;
    }



    // New method to extract the data of a row within a specified column range (startColIdx to endColIdx)
    private static String[] getRowData(XSSFSheet sheet, int rowIdx, int startColIdx, int endColIdx, boolean[][] visited) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return null;

        // Create an array for the specific column range
        String[] rowData = new String[endColIdx - startColIdx + 1];

        // Iterate through each column within the range and get the value
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
            // Check if the cell is marked as visited
            if (visited[rowIdx][colIdx]) {
                return null; // Stop processing as the row is invalid
            }

            Cell cell = row.getCell(colIdx);
            if (cell != null) {
                rowData[colIdx - startColIdx] = DataTypeUtils.getCellValue(sheet, rowIdx, colIdx);
            }
        }
        return rowData;
    }


    public static boolean compareRowDataTypes(String[] prevArr, String[] currArr) {

        // Compare the data types of prevArr and currArr for non-empty values
        for (int i = 0; i < prevArr.length; i++) {
            String prevValue = prevArr[i];
            String currValue = currArr[i];

            // If currValue is not empty, check the data types
            if (currValue != null && !currValue.isEmpty()) {
                if (prevValue != null && !prevValue.isEmpty()) {
                    // Compare the data types between prevArr and currArr at index i
                    if (!DataTypeUtils.detectDataType(prevValue).equals(DataTypeUtils.detectDataType(currValue))) {
                        return false; // Data type mismatch
                    }
                }
            }
        }

        // If max data types match or currArr is entirely empty/null, return true
        return true;
    }

    private static int narrowDownLastRowBoundary(XSSFSheet sheet, int startRowIdx, int startColIdx, int endColIdx, int currentEndRowIdx) {
        // Start from the currentEndRowIdx and check upwards
        for (int rowIdx = currentEndRowIdx; rowIdx >= startRowIdx; rowIdx--) {
            boolean allEmptyOrNull = true;

            // Check if all cells in the current row (from startColIdx to endColIdx) are empty or null
            for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
                String cellValue = DataTypeUtils.getCellValue(sheet, rowIdx, colIdx);

                if (cellValue != null && !cellValue.isEmpty()) {
                    // If any cell is not empty, we break and consider this row as the last boundary
                    allEmptyOrNull = false;
                    break;
                }
            }

            // If this row is not empty, update the last row boundary and break the loop
            if (!allEmptyOrNull) {
                return rowIdx; // This is the new last row boundary
            }
        }

        // If all rows are empty/null, return the original currentEndRowIdx
        return currentEndRowIdx;
    }


}
