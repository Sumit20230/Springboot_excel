package com.example.Excel_springboot.service;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@Service
public class ExcelService {

    public void readExcelFile(MultipartFile file) throws Exception {
        InputStream inputStream = file.getInputStream();
        Workbook workbook = new XSSFWorkbook(inputStream);

        Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
        for (Row row : sheet) {
            StringBuilder rowData = new StringBuilder();
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING -> rowData.append(cell.getStringCellValue()).append(" | ");
                    case NUMERIC -> rowData.append(cell.getNumericCellValue()).append(" | ");
                    case BOOLEAN -> rowData.append(cell.getBooleanCellValue()).append(" | ");
                    default -> rowData.append("UNKNOWN | ");
                }
            }
            System.out.println(rowData); // Print row data to the log
        }
        workbook.close();
    }
}
