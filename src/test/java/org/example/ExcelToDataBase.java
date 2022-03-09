package org.example;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class ExcelToDataBase {
    Connection con;
    XSSFWorkbook workbook;
    FileInputStream fis;
    @Test
    public void dbToExcel(){
        try{
            con= DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/exceldemo","root","Nayan@1990");
            Statement st=con.createStatement();
            //Create a new table 'MyDetails' in database
            String sql="CREATE TABLE exceldemo.MyDetails(id decimal(4,0),Name VARCHAR(45),org VARCHAR(45))";
            st.execute(sql);
            //Read data from excel
            fis=new FileInputStream("demo_dbTo_Excel.xlsx");
            workbook=new XSSFWorkbook(fis);
            XSSFSheet sheet=workbook.getSheetAt(0);
            int rows=sheet.getPhysicalNumberOfRows();

            for(int i=1;i<rows;i++){
                XSSFRow row=sheet.getRow(i);
                double id=row.getCell(0).getNumericCellValue();
                String name=row.getCell(1).getStringCellValue();
                String org=row.getCell(2).getStringCellValue();
                sql="insert into exceldemo.MyDetails values ('"+id+"','"+name+"','"+org+"')";
                st.execute(sql);
                st.execute("commit");
            }
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
                fis.close();
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
