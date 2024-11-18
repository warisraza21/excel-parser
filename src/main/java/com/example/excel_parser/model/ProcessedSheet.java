package com.example.excel_parser.model;


import java.io.Serializable;
import java.util.List;

public record ProcessedSheet(List<TableData> tableData, List<UnstructuredData> unstructuredData) implements Serializable {

}
