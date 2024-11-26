package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class RowDTO {
    private int index;
    private boolean isHeaderRow;
    private MergedDTO merged;
    private List<CellDTO> cells;


}
