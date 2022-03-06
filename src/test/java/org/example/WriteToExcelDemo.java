package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class WriteToExcelDemo {

    @Test
    public void writeUsingForLoop() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("My Info");
        Object [][]myData={{"Id","Name","City"},
                            {101,"Rahul","BLR"},
                            {201,"Deba","BLR"},
                            {301,"Subarna","KOL"}
                          };
        //using for loop
        int rows=myData.length;
        int cols=myData[0].length;
        System.out.println(rows+" "+cols);
        for(int r=0;r<rows;r++){
            XSSFRow row=sheet.createRow(r);
            for(int c=0;c<cols;c++)
            {
                XSSFCell cell=row.createCell(c);
                Object value=myData[r][c];
                if(value instanceof String)
                    cell.setCellValue(value.toString());
                else if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
                else if(value instanceof Integer)
                    cell.setCellValue((Integer) value);
            }
        }
        FileOutputStream fos=new FileOutputStream("demow.xlsx");
        workbook.write(fos);
        fos.close();
    }
    @Test
    public void writeUsingForEachLoop() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("My Info");
        Object [][]myData={{"Id","Name","City"},
                {101,"Rahul","BLR"},
                {201,"Deba","BLR"},
                {301,"Subarna","KOL"}
        };
        //using for each loop
        int rowCount=0;
        for(Object[] data:myData){//getting 1D array from the 2D Object array
            XSSFRow row=sheet.createRow(rowCount++);
            int columnCount=0;
            for(Object value:data){
                XSSFCell cell=row.createCell(columnCount++);
                if(value instanceof String)
                    cell.setCellValue(value.toString());
                else if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
                else if(value instanceof Integer)
                    cell.setCellValue((Integer) value);
            }
        }
        FileOutputStream fos=new FileOutputStream("demow.xlsx");
        workbook.write(fos);
        fos.close();
    }
    @Test
    public void writeUsingArrayList() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("My Info");
       /* Object [][]myData={{"Id","Name","City"},
                {101,"Rahul","BLR"},
                {201,"Deba","BLR"},
                {301,"Subarna","KOL"}
        };*/
        ArrayList<Object[]> myData=new ArrayList<>();
        myData.add(new Object[]{"Id","Name","City"});
        myData.add(new Object[]{101,"Rahul","BLR"});
        myData.add(new Object[]{201,"Deba","BLR"});
        myData.add(new Object[]{301,"Subarna","KOL"});

        //using for each loop
        int rowCount=0;
        for(Object[] data:myData){//getting 1D array from the 2D Object array
            XSSFRow row=sheet.createRow(rowCount++);
            int columnCount=0;
            for(Object value:data){
                XSSFCell cell=row.createCell(columnCount++);
                if(value instanceof String)
                    cell.setCellValue(value.toString());
                else if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
                else if(value instanceof Integer)
                    cell.setCellValue((Integer) value);
            }
        }
        FileOutputStream fos=new FileOutputStream("demow.xlsx");
        workbook.write(fos);
        fos.close();
    }

}
