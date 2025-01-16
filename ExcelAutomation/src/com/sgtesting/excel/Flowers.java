package com.sgtesting.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Flowers {
    public static void main(String[] args)
    {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Flowers");
        String[] flowers = {
                "Rose", "Tulip", "Sunflower", "Daisy", "Lily",
                "Orchid", "Jasmine", "Tulip", "Lavender", "Chrysanthemum",
                "Magnolia", "Marigold", "Violet", "Carnation", "Peony",
                "Geranium", "Daffodil", "Iris", "Camellia", "Aster"
        };

        try {
            for (int i = 0; i < flowers.length; i++) {
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(flowers[i]);
            }


            String filePath = "D:\\ExcelForJava\\Assignment2_Flowers.xlsx";

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
