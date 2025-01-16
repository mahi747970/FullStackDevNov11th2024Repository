package com.sgtesting.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class City {
    public static void main(String[] args)
    {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Cities");
        String[] cities = {
                "New York", "London", "Tokyo", "Paris", "Berlin",
                "Sydney", "Los Angeles", "Madrid", "Rome", "Delhi",
                "Toronto", "Dubai", "Shanghai", "Moscow", "Seoul",
                "Singapore", "Amsterdam", "Hong Kong", "Bangkok", "Istanbul"
        };

        try {
            Row row = sheet.createRow(9);

            for (int i = 0; i < cities.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(cities[i]);
            }
            String filePath = "D:\\ExcelForJava\\Assignment4_Cities.xlsx";

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("Excel file created successfully at: " + filePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
