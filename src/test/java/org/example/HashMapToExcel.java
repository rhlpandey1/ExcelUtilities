package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.TreeMap;

public class HashMapToExcel {

    @Test
    public void useHashMap() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("hashmap");
        Map<String,String> data=new LinkedHashMap<>();
        data.put("Id","Name");
        data.put("101","Rahul");
        data.put("102","Deba");
        data.put("103","Sub");
        data.put("104","King");

        int rowNo=0;
        for(Map.Entry<String,String> entry: data.entrySet()){
            XSSFRow row=sheet.createRow(rowNo++);
            XSSFCell cell1=row.createCell(0);
            cell1.setCellValue(entry.getKey());
            XSSFCell cell2=row.createCell(1);
            cell2.setCellValue(entry.getValue());
        }
        String filePath="HashMapToExcel.xlsx";
        FileOutputStream fos=new FileOutputStream(filePath);
        workbook.write(fos);
        workbook.close();
        fos.close();
    }
}
