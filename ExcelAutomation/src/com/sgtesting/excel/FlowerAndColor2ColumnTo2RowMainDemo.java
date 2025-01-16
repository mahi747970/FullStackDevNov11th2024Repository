package com.sgtesting.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class FlowerAndColor2ColumnTo2RowMainDemo {
    public static void main(String[] args) {
        columnToRow();
    }

    private static void columnToRow(){
        FileInputStream fin = null;
        FileOutputStream fout = null;
        Workbook workbook = null;
        Sheet sheet1 = null;
        Sheet sheet2 = null;
        Row rowsh1 = null;
        Row rowsh2 = null;
        Cell cellsh1 = null;
        Cell cellsh2 = null;
        try{
             fin = new FileInputStream("D:\\ExcelForJava\\flowercolor.xlsx");
             workbook = new XSSFWorkbook(fin);

               sheet1 =  workbook.getSheet("Sheet1");
               sheet2 = workbook.getSheet("Sheet2");
              if(sheet2 == null){
                  sheet2 = workbook.createSheet("Sheet2");
              }

              int rowCount = sheet1.getPhysicalNumberOfRows();
            for (int i = 0; i < rowCount; i++) {
                rowsh1 = sheet1.getRow(i);


                int cellCount = rowsh1.getPhysicalNumberOfCells();
                for (int j = 0; j < cellCount; j++) {

                    rowsh2 = sheet2.getRow(j+3);
                    if(rowsh2 == null){
                        rowsh2 = sheet2.createRow(j+3);
                    }

                    cellsh1 = rowsh1.getCell(j);

                    cellsh2 = rowsh2.getCell(i);
                    if(cellsh2 == null){
                        cellsh2 = rowsh2.createCell(i);
                    }

                    String dataOfCellSh1 = cellsh1.getStringCellValue();

                    cellsh2.setCellValue(dataOfCellSh1);

                }

            }

            fout = new FileOutputStream("D:\\ExcelForJava\\flowercolor.xlsx");
            workbook.write(fout);

        }catch(Exception e){
            e.printStackTrace();
        }finally {
            try {
                workbook.close();
                fin.close();
                fout.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }
}