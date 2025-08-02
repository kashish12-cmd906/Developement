package pck;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class FokusLogin {
	WebDriver driver;

	  @BeforeMethod

	    public void setup() {

		  System.setProperty("webdriver.chrome.driver", ".\\Driver\\chromedriver.exe");
			driver = new ChromeDriver();
			
	    }



@Test(dataProvider = "testData")





public void FokusLoginPage( String Email, String Password) throws InterruptedException {



driver.get("https://fokus.shop/account/login");


driver.manage().window().maximize();
Thread.sleep(2000);
//Alert prompt = driver.switchTo().alert();

//prompt.sendKeys("This Is Softwaretestinghelp");

//prompt.accept();
driver.switchTo().frame("relative=parent");
driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/img")).click();

Thread.sleep(2000);
Alert prompt = driver.switchTo().alert();
prompt.accept();

//driver.get("https://fokus.shop/account/login/Mobile.asp?filename=Mobilejs_prompt");
//driver.manage().window().maximize();
//Thread.sleep(2000);
//driver.switchTo().frame("relative=parent");
//driver.findElement(By.xpath("//*[@id=\\\"root\\\"]/div/div/div[1]/img")).click();
//Alert prompt = driver.switchTo().alert();
//prompt.sendKeys("7505981348");
//prompt.accept();





WebElement email = driver.findElement(By.xpath("//*[@id=\"CustomerEmail\"]"));
WebElement password = driver.findElement(By.xpath("//*[@id=\"CustomerPassword\"]"));


email.sendKeys(Email);



//This is the highlighter if you want to use

JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;

jsExecutor.executeScript("arguments[0].style.background='Lime'", email);

Thread.sleep(2000);
password.sendKeys(Password);


//This is the highlighter if you want to use

jsExecutor.executeScript("arguments[0].style.background='Lime'", password); 

driver.findElement(By.xpath("//*[@id=\\\"CustomerPassword\\\"]")).click();

Thread.sleep(2000);



driver.findElement(By.xpath("/html/body/main/div/div[1]/div[1]/form/div[2]/button")).click();



String actualText = driver.findElement(By.xpath("/html/body/main/div/div[1]/div[1]/form/div[2]/button")).getText();


//assertEquals(expectedText, actualText, "Success Message mismatch");



}





@AfterMethod



public void TearDown() {



driver.quit();



}



@DataProvider(name = "testData")
public Object[][] getData() throws IOException {
FileInputStream fis = new FileInputStream(new File("testData.xlsx"));
Workbook workbook = new XSSFWorkbook(fis);
org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
int rowCount = ((org.apache.poi.ss.usermodel.Sheet) sheet).getPhysicalNumberOfRows();
Object[][] data = new Object[rowCount - 1][2]; // Two fields
Iterator<Row> rows = sheet.iterator();
rows.next(); // Skip header row
int i = 0;
while (rows.hasNext()) {
Row row = rows.next();
Cell email = row.getCell(0); // Column A
Cell password = row.getCell(1);// Column B
data[i][0] = email != null ? email.toString() : "";
data[i][1] = password != null ? password.toString() : "";
i++;
}
workbook.close();   
fis.close();
return data;
}}

