package com.sgtesting.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Fruits {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Fruits");
        String[] fruits = {
                "Apple", "Banana", "Mango", "Grapes", "Orange",
                "Pineapple", "Strawberry", "Peach", "Watermelon", "Blueberry",
                "Papaya", "Cherry", "Kiwi", "Lemon", "Pomegranate",
                "Avocado", "Pear", "Plum", "Coconut", "Guava"
        };

        try {

            Row row = sheet.createRow(0);
            for (int i = 0; i < fruits.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(fruits[i]);
            }
            String filePath = "D:\\ExcelForJava\\Assignment3_Fruits.xlsx";

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("Excel file created successfully at: " + filePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally
        {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
