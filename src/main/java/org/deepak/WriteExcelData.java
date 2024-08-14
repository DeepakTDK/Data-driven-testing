package org.deepak;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelData {
    public static void main(String[] args) throws IOException {
        FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\Testdata\\data1.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Data");

        XSSFRow row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("java");
        row1.createCell(1).setCellValue(10);
        row1.createCell(2).setCellValue("automation");

        XSSFRow row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("python");
        row2.createCell(1).setCellValue(19);
        row2.createCell(2).setCellValue("devops");

        workbook.write(file);
        workbook.close();
        file.close();

        System.out.println("File created");

    }
}
