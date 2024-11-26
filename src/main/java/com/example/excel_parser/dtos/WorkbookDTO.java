package com.example.excel_parser.dtos;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class WorkbookDTO {
    private String id;
    private String name;
    private String excelVersion;
    private String createdAt;
    private String lastModifiedAt;
    private int sheetCount;
    private String fileSize;
    private boolean isProtected;
    private List<SheetDTO> sheets;

    private List<NamedRangeDTO> namedRanges;

}

