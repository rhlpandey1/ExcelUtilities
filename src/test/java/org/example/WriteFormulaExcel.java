package org.example;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFormulaExcel {

    @Test
    public void writeToExcelUsingFormula() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Formula Demo");
        XSSFRow row=sheet.createRow(0);
        row.createCell(0).setCellValue(100);
        row.createCell(1).setCellValue(200);
        row.createCell(2).setCellValue(300);
        row.createCell(3).setCellFormula("A1+B1+C1");
        FileOutputStream fos=new FileOutputStream("demo_formula_write.xlsx");
        workbook.write(fos);
    }
}
