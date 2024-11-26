package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class TabularDataDTO {
    private String id;
    private String name;
    private String type;
    private int columnsCount;
    private int rowsCount;
    private BoundariesDTO boundaries;
    private List<RowDTO> rows;
    private TotalRowDTO totalRow;


}
