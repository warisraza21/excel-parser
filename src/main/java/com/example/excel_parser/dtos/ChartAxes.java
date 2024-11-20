package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class ChartAxes {
    private ChartAxis xAxis;
    private ChartAxis yAxis;
}

@Data
class ChartAxis {
    private String label;
    private String dataType;
}
