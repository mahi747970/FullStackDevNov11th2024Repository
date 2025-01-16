package com.sgtesting.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Country {
    public static void main(String[] args) {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Countries");

        String[] countries = {
                "India", "USA", "Canada", "Germany", "France",
                "China", "Japan", "Brazil", "Russia", "Australia",
                "Mexico", "Italy", "South Korea", "Spain", "Indonesia",
                "Turkey", "UK", "South Africa", "Netherlands", "Saudi Arabia"
        };

        try {

            for (int i = 0; i < countries.length; i++) {
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(i);
                cell.setCellValue(countries[i]);
            }
            String filePath = "D:\\ExcelForJava\\Assignment1_Countries.xlsx";

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
