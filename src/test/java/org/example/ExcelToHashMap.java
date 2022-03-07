package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelToHashMap {

    @Test
    public void useHashMap() throws IOException {
        String filePath="ExcelToHashMap.xlsx";
        FileInputStream fis=new FileInputStream(filePath);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet sheet=workbook.getSheet("hashmap");
        int rows=sheet.getPhysicalNumberOfRows();
        Map<String,String> data=new LinkedHashMap<>();
        for(int i=0;i<rows;i++){
            int noOfCells=sheet.getRow(0).getLastCellNum();
            for(int j=0;j<noOfCells;j++){
                String key=sheet.getRow(0).getCell(j).getStringCellValue();
                String value=sheet.getRow(i).getCell(j).getStringCellValue();
                data.put(key,value);
            }
            }

        for(Map.Entry<String,String> entry:data.entrySet()){
            System.out.println(entry.getKey()+"|"+entry.getValue());
        }

    }
}
