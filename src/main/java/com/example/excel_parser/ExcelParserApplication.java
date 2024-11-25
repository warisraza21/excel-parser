package com.example.excel_parser;

import com.example.excel_parser.service.ExcelReader;
import com.example.excel_parser.service.WorkBookInfoBuilder;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@Slf4j
@SpringBootApplication
public class ExcelParserApplication {

	public static final ObjectMapper objectMapper = new ObjectMapper();

	public static void main(String[] args) throws JsonProcessingException {


		String filePath = "src/main/resources/static/test.xlsx";
		SpringApplication.run(ExcelParserApplication.class, args);
		ExcelReader.testExcel(filePath);
//		log.info("WorkBookInfo : {}", objectMapper.writeValueAsString(WorkBookInfoBuilder.buildWorkBookInfo(filePath)));
	}
}
