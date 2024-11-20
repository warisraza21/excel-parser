package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class PivotTableFields {
    private List<PivotTableField> rowFields;
    private List<PivotTableField> columnFields;
    private List<PivotTableValueField> valueFields;
    private List<PivotTableFilterField> filterFields;
}

@Data
class PivotTableField {
    private String name;
    private String sortOrder;
}

@Data
class PivotTableValueField {
    private String name;
    private String function;
    private String format;
}

@Data
class PivotTableFilterField {
    private String name;
    private List<String> criteria;
}
