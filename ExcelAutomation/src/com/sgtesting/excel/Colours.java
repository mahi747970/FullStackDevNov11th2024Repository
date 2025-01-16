package com.sgtesting.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Colours {
    public static void main(String[] args) {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Colors");

        String[] colors = {
                "Red", "Blue", "Green", "Yellow", "Pink",
                "Orange", "Purple", "Brown", "Black", "White",
                "Gray", "Violet", "Indigo", "Cyan", "Magenta",
                "Maroon", "Beige", "Turquoise", "Lavender", "Gold"
        };

        try {
            for (int i = 0; i < colors.length; i++) {
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(4);  // Column index 4 corresponds to the 5th column (Column E)
                cell.setCellValue(colors[i]);
            }
            String filePath = "D:\\ExcelForJava\\Assignment5_Colors.xlsx";

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
