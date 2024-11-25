package com.example.excel_parser.service;

import com.example.excel_parser.utils.DataTypeUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;

public class TableSegregation {

    private static Set<String> reservedBoundaries = new HashSet<>(); // Set to track reserved boundaries

    public static List<int[][]> segregateTablesRecursively(XSSFSheet sheet, int[] firstCellIndex, int[] lastCellIndex) {
        int firstRow = firstCellIndex[0];
        int firstCol = firstCellIndex[1];
        int lastRow = lastCellIndex[0];
        int lastCol = lastCellIndex[1];

        List<int[][]> tables = new ArrayList<>();

        // Recursively identify tables within the cluster range
        tables.addAll(getFirstRowBoundaryLineList(sheet, firstCellIndex, lastCellIndex));

        // After identifying a table, we recursively call the function for the next set of rows.
        int nextRowStart = getLastRowIndex(sheet, firstRow, lastRow, firstCol, lastCol) + 1;
        if (nextRowStart <= lastRow) {
            tables.addAll(segregateTablesRecursively(sheet, new int[]{nextRowStart, firstCol}, new int[]{lastRow, lastCol}));
        }

        return tables;
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

            // Reserve this boundary for future calls
            String boundaryKey = rowIndex + "-" + colIndices[0] + "-" + rowIndex + "-" + colIndices[1];
            reservedBoundaries.add(boundaryKey);
        }

        return lineEndpointList;
    }

    private static int checkLeftConsistency(XSSFSheet sheet, int rowIndex, int colStartIndex, int firstCol) {
        int latestColIndex = colStartIndex;

        // Iterate over the columns to the left of the line start
        for (int colIndex = colStartIndex - 1; colIndex >= firstCol; colIndex--) {
            Cell cell = sheet.getRow(rowIndex).getCell(colIndex);

            if (cell == null || cell.getCellType() == CellType.BLANK) {
                return colIndex + 1; // Return the current column index if the cell is empty
            }

            if (DataTypeUtils.checkDataTypeRow(DataTypeUtils.getCellValue(cell))) {
                return -1; // Return -1 if the cell is not a String
            }

            latestColIndex = colIndex;
        }

        return latestColIndex;
    }

    private static int checkRightConsistency(XSSFSheet sheet, int rowIndex, int colEndIndex, int lastCol) {
        int latestColIndex = colEndIndex;

        // Iterate over the columns to the right of the line end
        for (int colIndex = colEndIndex + 1; colIndex <= lastCol; colIndex++) {
            Cell cell = sheet.getRow(rowIndex).getCell(colIndex);

            if (cell == null || cell.getCellType() == CellType.BLANK) {
                return colIndex - 1; // Return the current column index if the cell is empty
            }

            if (DataTypeUtils.checkDataTypeRow(DataTypeUtils.getCellValue(cell))) {
                return -1; // Return -1 if the cell is not a String
            }

            latestColIndex = colIndex;
        }

        return latestColIndex;
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

    public static int getLastRowIndex(XSSFSheet sheet, int rowStartIndex, int rowEndIndex, int colStartIndex, int colEndIndex) {
        int[] lastRowIndexes = new int[colEndIndex - colStartIndex + 1];
        int idx = 0;
        for (int colIndex = colStartIndex; colIndex <= colEndIndex; colIndex++) {
            String initialDataType = null;
            int maxRowIdx = -1;
            for (int rowIndex = rowStartIndex + 1; rowIndex <= rowEndIndex; rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        String cellValue = DataTypeUtils.getCellValue(cell);

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
        for (int item : lastRowIndexes) minRow = Math.min(minRow, item);
        return minRow == Integer.MAX_VALUE ? -1 : minRow;
    }
}

