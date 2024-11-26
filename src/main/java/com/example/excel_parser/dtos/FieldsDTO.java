package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class FieldsDTO {
    private List<FieldDTO> rowFields;
    private List<FieldDTO> columnFields;
    private List<ValueFieldDTO> valueFields;
    private List<FilterFieldDTO> filterFields;


}
