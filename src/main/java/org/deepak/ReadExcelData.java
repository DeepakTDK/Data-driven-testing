package org.deepak;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcelData {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\Testdata\\data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int lastrow = sheet.getLastRowNum();
        int lastcell = sheet.getRow(1).getLastCellNum();

        for(int i=0; i<=lastrow; i++){
            XSSFRow currrow = sheet.getRow(i);
            for(int j=0; j<lastcell; j++){
                XSSFCell currcell = currrow.getCell(j);
                System.out.print(currcell.toString()+" ");
            }
            System.out.println("");
        }
        workbook.close();
        file.close();

    }




}
