package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class ChartInfo {
    private String id;
    private String name;
    private String type;
    private CellRangeInfo boundaries;
    private ChartSourceData source;
    private List<ChartLegend> legends;
    private ChartTitle title;
    private ChartAxes axes;
}

