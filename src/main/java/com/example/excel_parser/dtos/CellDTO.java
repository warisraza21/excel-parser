package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class CellDTO {
    private int rowIndex;
    private int columnIndex;
    private String value;
    private String valueType;
    private FormulaDTO formula;
    private MergedDTO merged;
    private DataValidationDTO dataValidation;
    private FormattingDTO formatting;


}
