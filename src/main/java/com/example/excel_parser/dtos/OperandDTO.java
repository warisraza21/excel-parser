package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class OperandDTO {
    private String source;
    private String sourceType;
    private int rowIndex;
    private int columnIndex;
    private String columnName;


}
