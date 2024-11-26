package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ColumnAggregationDTO {
    private String columnName;
    private String aggregationFunction;
    private double calculatedValue;


}
