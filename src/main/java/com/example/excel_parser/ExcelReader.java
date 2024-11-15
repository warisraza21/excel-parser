package com.example.excel_parser;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;


@Slf4j
public class ExcelReader {

    public static void testExcel(){
        String filePath = "src/main/resources/static/test.xlsx";
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            log.info("Total sheets: {}" , workbook.getNumberOfSheets());

            for(Sheet sheet : workbook){
                XSSFSheet xssfSheet = (XSSFSheet) sheet;
                log.info("Sheet Name : {}",xssfSheet.getSheetName());
                log.info("Sheet : {} contains total table : {}", sheet.getSheetName(),xssfSheet.getTables().size());
                log.info("Sheet : {} contains total pivot table : {}", sheet.getSheetName(),xssfSheet.getPivotTables().size());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}