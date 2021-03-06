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
    Connection con;
    FileOutputStream fos;
    XSSFWorkbook workbook;
    @Test
    public void dbToExcel(){
        try{
            con= DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/exceldemo","root","Nayan@1990");
            Statement st=con.createStatement();
            ResultSet rs=st.executeQuery("SELECT * FROM exceldemo.EmpDetails");
            workbook=new XSSFWorkbook();
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
            fos=new FileOutputStream("demo_dbTo_Excel.xlsx");
            workbook.write(fos);
        }catch (SQLException sqe){
            System.out.println(sqe.getMessage());
        }catch (IOException ie){
            System.out.println(ie.getMessage());
        }
       finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                con.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}
