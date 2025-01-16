package com.sgtesting.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReverseBrids {
    public static void main(String[] args)
    {
        FileInputStream fin=null;
        FileOutputStream fout=null;
        Workbook workbook=null;
        Sheet sh1=null;
        Sheet sheet2=null;
        Row rowsh1=null;
        Row rowsh2=null;
        Cell cellsh1=null;
        Cell cellsh2=null;

        String filePath = ("D:\\ExcelForJava\\Brids.xlsx");

        try
        {
            fin = new FileInputStream(filePath);
              workbook = new XSSFWorkbook(fin);

            Sheet sheet1 = workbook.getSheetAt(0);

            String[] birds = new String[20];
            for (int i = 0; i < birds.length; i++) {
                Row row = sheet1.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(4);
                    if (cell != null) {
                        birds[i] = cell.getStringCellValue();
                    }
                }
            }
             sheet2 = workbook.getSheet("Sheet2");
            if (sheet2 == null) {
                sheet2 = workbook.createSheet("Sheet2");
            }
            for (int i = 0; i < birds.length; i--) {
                Row row = sheet2.createRow(i);
                if (row == null) {
                    row = sheet2.createRow(i);
                }
                Cell cell = row.createCell(4);
                cell.setCellValue(birds[birds.length - 1 - i]);
            }
            fout = new FileOutputStream(filePath);
                workbook.write(fout);
                System.out.println("Excel file update successfully at:" + filePath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

