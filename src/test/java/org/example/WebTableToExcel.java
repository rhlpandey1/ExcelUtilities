package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.IOException;
import java.time.Duration;

public class WebTableToExcel {
    WebDriver driver;
    @BeforeClass
    public void setUp(){
        WebDriverManager.chromedriver().setup();
        driver=new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        driver.manage().window().maximize();
        driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
    }
    @Test()
    public void testWebTable() throws IOException {
        String path="countries.xlsx";
        ExcelUtility utility=new ExcelUtility(path);
        //write headers
        String[] headers={"Country","Region","Population","Percentage of the world","Date","Source","Notes"};
        for(int i=0;i<headers.length;i++){
            utility.setCellData("Sheet1",0,i,headers[i]);
        }
        WebElement table=driver.findElement(By.xpath("//table[contains(@class,'wikitable')]/tbody"));
        //rows present in Web Table
        int rowNum=table.findElements(By.xpath("tr")).size();
        for(int i=2;i<=rowNum;i++){
            for(int j=1;j<=headers.length;j++){
                String text=table.findElement(By.xpath("tr["+i+"]/td["+j+"]")).getText();
                System.out.println(text);
                utility.setCellData("Sheet1",i,j-1,text);
            }
        }

    }
    @AfterClass
    public void tearDown(){
        driver.quit();
    }
}
