package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ChartSourceDTO {
    private String source;
    private String sourceType;
    private BoundariesDTO boundaries;
    private List<ChartColumnDTO> columns;


}
