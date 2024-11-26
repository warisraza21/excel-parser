package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class FormulaDTO {
    private boolean structured;
    private String expression;
    private List<OperandDTO> operands;


}
