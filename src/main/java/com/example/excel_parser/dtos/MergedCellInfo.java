package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class MergedCellInfo {
    private String range;
    private String mergeType;  // E.g., Vertical or Horizontal
}

