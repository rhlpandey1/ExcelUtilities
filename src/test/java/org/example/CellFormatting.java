package org.example;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CellFormatting {

    @Test
    public void formatCell() throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("formatter");
        XSSFRow row=sheet.createRow(1);

        //background color
        XSSFCellStyle style=workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        XSSFCell cell=row.createCell(1);
        cell.setCellValue("Rahul");
        cell.setCellStyle(style);


        //foreground color
        style=workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        cell=row.createCell(2);
        cell.setCellValue("Pandey");
        cell.setCellStyle(style);

        String filePath="formattedExcel.xlsx";
        FileOutputStream fos=new FileOutputStream(filePath);
        workbook.write(fos);

        workbook.close();
        fos.close();
    }
}
