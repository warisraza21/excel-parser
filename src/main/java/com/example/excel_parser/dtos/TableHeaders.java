package com.example.excel_parser.dtos;

import lombok.Data;

@Data
public class TableHeaders {
    private HeaderStyle rowHeaders;
    private HeaderStyle columnHeaders;
}


@Data
class HeaderStyle {
    private String fontStyle;  // Bold, Italic, etc.
    private int fontSize;
    private String alignment;  // Center, Left, Right
    private String color;
}

