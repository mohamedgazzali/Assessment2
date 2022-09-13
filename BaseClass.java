package testcase;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	
	
		
		
		// TODO Auto-generated method stub
		
		public ChromeDriver driver;
		
		public String filename;
		
		@Parameters ({"url","username","password"})
		
		@BeforeMethod
		public void precondition(String url, String username, String password) {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(url);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.findElement(By.id("username")).sendKeys(username);
		driver.findElement(By.id("password")).sendKeys(password);
		driver.findElement(By.className("decorativeSubmit")).click();
		driver.findElement(By.linkText("CRM/SFA")).click();
		driver.findElement(By.linkText("Leads")).click();
		
		}
		

		@AfterMethod
		public void postcondition() {
			driver.close();
		}
		
		@DataProvider(name="fetchData")
		public String [][] readData()throws IOException
		{
			
			return ReadExcel.ReadExcelData(filename);
			
		}

}
