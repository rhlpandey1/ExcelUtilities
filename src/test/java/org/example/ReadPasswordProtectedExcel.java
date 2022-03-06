package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadPasswordProtectedExcel {
    @Test
    public void readProtectedFile() throws IOException {
        FileInputStream fis=new FileInputStream("demow_pp.xlsx");
        String password="test123";
        XSSFWorkbook workbook=(XSSFWorkbook)WorkbookFactory.create(fis,password);
        XSSFSheet sheet=workbook.getSheet("Sheet1");
        int rowNum=sheet.getPhysicalNumberOfRows();
        int colNum=sheet.getRow(0).getLastCellNum();
        System.out.println(rowNum+" "+colNum);
        for(int i=0;i<rowNum;i++){
            XSSFRow row=sheet.getRow(i);
            for(int j=0;j<colNum;j++){
                XSSFCell cell=row.getCell(j);
                switch (cell.getCellType()){
                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue());
                    case STRING -> System.out.print(cell.getStringCellValue());
                    case NUMERIC -> System.out.print(cell.getNumericCellValue());
                    default -> System.out.println("NaN");
                }
                System.out.print("|");
            }
            System.out.println();
        }
        workbook.close();
        fis.close();
    }
    @Test
    public void readProtectedFileIterator() throws IOException {
        FileInputStream fis=new FileInputStream("demow_pp.xlsx");
        String password="test123";
        XSSFWorkbook workbook=(XSSFWorkbook)WorkbookFactory.create(fis,password);
        XSSFSheet sheet=workbook.getSheet("Sheet1");
        Iterator<Row> it=sheet.iterator();
        while(it.hasNext()){
            XSSFRow row=(XSSFRow) it.next();
            Iterator<Cell> cellIterator=row.cellIterator();
            while (cellIterator.hasNext()){
                XSSFCell cell=(XSSFCell) cellIterator.next();
                switch (cell.getCellType()){
                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue());
                    case STRING -> System.out.print(cell.getStringCellValue());
                    case NUMERIC -> System.out.print(cell.getNumericCellValue());
                    default -> System.out.println("NaN");
                }
                System.out.print("|");
            }
            System.out.println();
        }
        workbook.close();
        fis.close();
    }
}
