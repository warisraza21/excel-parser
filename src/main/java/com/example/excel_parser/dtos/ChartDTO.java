package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ChartDTO {
    private String id;
    private String name;
    private String type;
    private BoundariesDTO boundaries;
    private ChartSourceDTO source;
    private List<LegendDTO> legends;
    private ChartTitleDTO title;
    private AxesDTO axes;


}
