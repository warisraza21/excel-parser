package com.example.excel_parser;

import com.example.excel_parser.service.ExcelReader;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelParserApplication {

	public static void main(String[] args) {

		String filePath = "src/main/resources/static/test.xlsx";
		SpringApplication.run(ExcelParserApplication.class, args);
		ExcelReader.testExcel(filePath);
	}
}
