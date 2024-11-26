package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class PivotTableDTO {
    private String id;
    private String name;
    private BoundariesDTO boundaries;
    private SourceDTO sources;
    private FieldsDTO fields;
    private PivotOptionsDTO options;
    private List<RowDTO> rows;


}
