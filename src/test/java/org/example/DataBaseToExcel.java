package org.example;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class DataBaseToExcel {

    @Test
    public void dbToExcel() throws ClassNotFoundException, SQLException, IOException {
        Connection con= DriverManager.getConnection(
                "jdbc:mysql://localhost:3306/exceldemo","root","Nayan@1990");
        Statement st=con.createStatement();
        ResultSet rs=st.executeQuery("SELECT * FROM exceldemo.EmpDetails");
      /*  while(rs.next()){
            System.out.print(rs.getString(1));
            System.out.print("|");
            System.out.print(rs.getString(2));
            System.out.print("|");
            System.out.print(rs.getString(3));
            System.out.println();
        }*/
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("dbDemo");
        XSSFRow row=sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("Name");
        row.createCell(2).setCellValue("org");
        int i=1;
        while (rs.next()){
            double id=rs.getDouble("id");
            String name=rs.getString("Name");
            String org=rs.getString("org");
            row=sheet.createRow(i++);
            row.createCell(0).setCellValue(id);
            row.createCell(1).setCellValue(name);
            row.createCell(2).setCellValue(org);
        }
        FileOutputStream fos=new FileOutputStream("demo_dbTo_Excel.xlsx");
        workbook.write(fos);
        con.close();
    }
}
