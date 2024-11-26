package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class NamedRangeDTO {
    private String name;
    private String boundary;
    private String sheetName;
    private String range;
    private boolean createdFromTable;
}

