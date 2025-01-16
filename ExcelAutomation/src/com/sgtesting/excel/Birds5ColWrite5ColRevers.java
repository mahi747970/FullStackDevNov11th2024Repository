package com.sgtesting.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Birds5ColWrite5ColRevers
{
    public static void main(String[] args)
    {
        FileInputStream fin= null;
        FileOutputStream fout= null;
        Workbook workbook=null;
        Sheet sheet1= null;
        Sheet sheet2= null;
        Row rowsh1 = null;
        Row rowsh2 = null;
        Cell cellsh1 = null;
        Cell cellsh2 = null;
        try
        {
            fin = new FileInputStream("D:\\ExcelForJava\\Birds5ColWrite5ColRevers.xlsx");
            workbook =new XSSFWorkbook(fin);
            sheet1 =workbook.getSheet("Sheet1");
            sheet2 =workbook.getSheet("Sheet2");
            if(sheet2 == null)
            {
                sheet2=workbook.createSheet("Sheet2");
            }
            int rowCount= sheet1.getPhysicalNumberOfRows();
            int k=0;
            for(int  i= rowCount-1; i>=0; i--)
            {
                
                rowsh1 = sheet1.getRow(i);

                rowsh2 = sheet2.getRow(k);
                if (rowsh2 == null)
                {
                    rowsh2 = sheet2.createRow(k);

                }
                k++;
                    cellsh1 = rowsh1.getCell(4);

                    cellsh2 = rowsh2.getCell(4);
                    if (cellsh2 == null) {
                        cellsh2 = rowsh2.createCell(4);

                        String dataOfCellSh1 = cellsh1.getStringCellValue();
                        cellsh2.setCellValue(dataOfCellSh1);

                    }



            }

            fout = new FileOutputStream("D:\\ExcelForJava\\Birds5ColWrite5ColRevers.xlsx");
            workbook.write(fout);
        } catch (Exception e)
        {
            e.printStackTrace();
        }
        try
        {

        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
