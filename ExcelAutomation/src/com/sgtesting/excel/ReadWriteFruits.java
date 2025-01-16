package com.sgtesting.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteFruits
{
    public static void main(String[] args)
    {
        String filePath = "D:\\ExcelForJava\\Fruits.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis))
        {
            Sheet sheet1 = workbook.getSheetAt(0);

            String[] fruits = new String[20];
            for (int i = 0; i < 20; i++) {
                Row row = sheet1.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        fruits[i] = cell.getStringCellValue();
                    }
                }
            }
            Sheet sheet2 = workbook.createSheet("Sheet2");

            Row row = sheet2.createRow(4);

            for (int i = 0; i < fruits.length; i++)
            {
                Cell cell = row.createCell(i);
                cell.setCellValue(fruits[i]);
            }
            try (FileOutputStream fos = new FileOutputStream(filePath))
            {
                workbook.write(fos);
                System.out.println("Excel file updated successfully at: " + filePath);
            }
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
}
