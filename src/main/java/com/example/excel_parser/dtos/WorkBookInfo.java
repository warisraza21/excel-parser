package com.example.excel_parser.dtos;

import lombok.Data;

import java.util.List;

@Data
public class WorkBookInfo {
    private String id;
    private String name;
    private String excelVersion;
    private String createdAt;
    private String lastModifiedAt;
    private int sheetCount;
    private String fileSize;
    private boolean isProtected;
    private List<SheetInfo> sheets;
}

