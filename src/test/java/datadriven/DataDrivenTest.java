package datadriven;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.example.ExcelUtility;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;
import java.time.Duration;

public class DataDrivenTest {
    WebDriver driver;
    @BeforeClass
    public void setUp(){
        WebDriverManager.chromedriver().setup();
        driver=new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        driver.manage().window().maximize();
        driver.get("https://admin-demo.nopcommerce.com/login");
    }
    @Test(dataProvider = "loginData")
    public void testLogin(String user,String pwd,String expected){
        try{
            System.out.println(user+" "+pwd+" "+expected);
            driver.findElement(By.id("Email")).clear();
            driver.findElement(By.id("Email")).sendKeys(user);
            driver.findElement(By.id("Password")).clear();
            driver.findElement(By.id("Password")).sendKeys(pwd);
            driver.findElement(By.cssSelector("[class*='login-button']")).click();
            String expectedTitle="Dashboard / nopCommerce administration";
            if(expected.equalsIgnoreCase("valid")){
                Assert.assertTrue(driver.findElement(By.cssSelector("[class*='brand-image-xl']")).isDisplayed(),"Verification failed after login to get the brand logo");
                Assert.assertEquals(driver.getTitle(),expectedTitle,"Verification failed for getting the title");
            }
            else if(expected.equalsIgnoreCase("invalid")){
                Assert.assertNotEquals(driver.getTitle(),expectedTitle,"Verification failed for getting the title");
            }
        }finally {
            try{
                WebElement logout=driver.findElement(By.linkText("Logout"));
                if(logout.isDisplayed())
                    logout.click();
            }catch (Exception e){
                System.out.println("Logout is not appeared");
            }
        }

    }

    @DataProvider(name="loginData")
    public Object[][] getData() throws IOException {
       /* Object[][] loginData={{"admin@yourstore.com","admin","valid"},
                              {"admin@yourstore.com","admn","invalid"},
                              {"amd@yourstore.com","admin","invalid"},
                              {"amd@yourstore.com","adm","invalid"}
                            };*/
        //get data from excel
        String path="LoginTestData.xlsx";
        String sheetName="Sheet1";
        ExcelUtility utility=new ExcelUtility(path);
        int noOfRows=utility.getRowCount(sheetName);
        int noOfColumns=utility.getCellCount(sheetName,0);
        Object[][] loginData=new Object[noOfRows][noOfColumns];
        for(int i=1;i<=noOfRows;i++){
            for (int j=0;j<noOfColumns;j++){
                loginData[i-1][j]=utility.getCellData(sheetName,i,j);
            }
        }
        return loginData;
    }
    @AfterClass
    public void tearDown(){
        driver.quit();
    }

}
