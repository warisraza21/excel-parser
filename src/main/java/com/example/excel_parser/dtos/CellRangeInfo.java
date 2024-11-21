package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class CellRangeInfo {
    private String startCell;
    private String endCell;
}

