package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class DataValidationRules {
    private String type;  // E.g., List, Range, etc.
    private List<String> criteria;
}

