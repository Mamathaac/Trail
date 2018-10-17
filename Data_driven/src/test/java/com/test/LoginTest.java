package com.test;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.utility.ExcellUtility;

public class LoginTest {
	
	WebDriver driver;
	
	@BeforeTest
	
	public void setUp()
	{
System.setProperty("webdriver.chrome.driver","C:\\Users\\mamatha.a.c\\Downloads\\chromedriver_win32\\chromedriver.exe" );
		
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		
	driver.get("https://www.phptravels.net/admin");
	
		
	}
	
	@Test
	
public void login() throws InterruptedException 
{
		
		WebElement emailbox = null;
		WebElement pwdbox = null;
		WebElement loginbutton=null;
		
		//button[@type='submit']
		///html/body/div/form[1]/button
		
		String excelpath = "C:\\Users\\mamatha.a.c\\eclipse-workspace\\Data_driven\\src\\test\\resources\\Credentials.xlsx";
		
		ExcellUtility exu = new ExcellUtility(excelpath);
		
		String sheetName = "admin";
		int rowCount = exu.getNumberOfRows(sheetName);
		int colCount = exu.getNumberOfCos(sheetName);
		
		String email = null;
		String pwd = null;
		
		for(int i=1; i<rowCount; i++) 
		{
			email = exu.getCellData(sheetName, i, 0);
			pwd = exu.getCellData(sheetName, i, 1);
			
			
			 emailbox = driver.findElement(By.name("email"));
			 pwdbox = driver.findElement(By.name("password"));
			 loginbutton = driver.findElement(By.xpath("//button[@type='submit']"));
			
			
			emailbox.sendKeys(email);
			pwdbox.sendKeys(pwd);
			loginbutton.click();
			Thread.sleep(6000);
			
			String pageTitle = driver.getTitle();
			if(pageTitle.equals("Dashboard"))
			{
				
				
			
				exu.setCellData(sheetName, i, 2, "Pass");
				driver.findElement(By.xpath("//a[contains(text(),'Log Out')]")).click();
			}
			else
			{
				
				exu.setCellData(sheetName, i, 2, "Fail");
				emailbox.clear();
				pwdbox.clear();
			}

			
		}
		
		String destinationPath = "C:\\Users\\mamatha.a.c\\Desktop\\result.xlsx";
		exu.writeAndSaveExcel(destinationPath);
		
		
}
	
//	@AfterTest
//	public void tearDown()
//	{
//		
//		driver.close();
//	}

}
