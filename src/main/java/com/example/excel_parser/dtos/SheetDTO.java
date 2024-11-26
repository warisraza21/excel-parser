package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class SheetDTO {
    private String id;
    private String name;
    private int sheetIndex;
    private String visibility;
    private List<TabularDataDTO> tabularData;
    private List<PivotTableDTO> pivotTables;
    private List<ChartDTO> charts;
    private List<MiscellaneousDTO> miscellaneous;


}
