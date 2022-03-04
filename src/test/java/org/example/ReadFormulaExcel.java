package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadFormulaExcel {
    String filePath="demo_formula.xlsx";
    FileInputStream fis=new FileInputStream(filePath);
    XSSFWorkbook xssfWorkbook=new XSSFWorkbook(fis);
    XSSFSheet sheet=xssfWorkbook.getSheet("Sheet1");

    public ReadFormulaExcel() throws IOException {
    }

    @Test
    public void readFormulaExcel(){
        int noOfRows=sheet.getPhysicalNumberOfRows();
        int noOfColumns=sheet.getRow(1).getLastCellNum();
        System.out.println("noOfColumns "+noOfColumns);
        for(int i=0;i<noOfRows;i++){
            XSSFRow row=sheet.getRow(i);
            for(int j=0;j<noOfColumns;j++){
                XSSFCell cell=row.getCell(j);
                switch (cell.getCellType()){
                    case NUMERIC :
                    case FORMULA:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;

                }
                System.out.print("|");
            }
            System.out.println();
        }
    }
}
