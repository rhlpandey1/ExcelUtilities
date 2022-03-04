package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcelTest {
    String filePath="demo.xlsx";
    FileInputStream fis=new FileInputStream(filePath);
    XSSFWorkbook xssfWorkbook=new XSSFWorkbook(fis);
    XSSFSheet sheet=xssfWorkbook.getSheet("Sheet1");

    public ReadExcelTest() throws IOException {

    }

    @Test
    public void readExcelUsingForLoop() throws IOException {
        int noOfRows=sheet.getPhysicalNumberOfRows();
        int noOfColumns=sheet.getRow(1).getLastCellNum();
        System.out.println("noOfColumns "+noOfColumns);
        for(int i=0;i<noOfRows;i++){
            XSSFRow  row=sheet.getRow(i);
            for(int j=0;j<noOfColumns;j++){
                XSSFCell cell=row.getCell(j);
                switch (cell.getCellType()){
                    case STRING:
                        System.out.print(cell.getStringCellValue()+" ");
                        break;
                    case NUMERIC:
                        System.out.print((int)cell.getNumericCellValue()+" ");
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue()+" ");
                        break;
                }
            }
            System.out.println();
        }

    }
    @Test
    public void readExcelDataUsingIterator(){
        Iterator iterator=sheet.iterator();
        while(iterator.hasNext()){
            XSSFRow row= (XSSFRow) iterator.next();
            Iterator cellIterator= row.cellIterator();
            while (cellIterator.hasNext()){
                XSSFCell cell= (XSSFCell) cellIterator.next();
                switch (cell.getCellType()){
                    case STRING:
                        System.out.print(cell.getStringCellValue()+" ");
                        break;
                    case NUMERIC:
                        System.out.print((int)cell.getNumericCellValue()+" ");
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue()+" ");
                        break;
                }
            }
            System.out.println();
        }
    }

}
