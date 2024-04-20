package com.poly.test_resultExcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class SaveTestNGtoResultExcel {
	WebDriver driver;
	String url = "https://id.zing.vn/";
	String username = "huunghia8844";
	String password = "Huunghia123";
	
	Map<String, Object[]> testNGResult;
	
	HSSFWorkbook workBook;
	HSSFSheet sheet;

	@Test(priority = 0)
	public void launchWebsite() {
		try {
			driver.get(url);
			driver.manage().window().maximize();
			testNGResult.put("2", new Object[] {"1d", "Navigate to demo website", "Site gets opened", "Pass"});
		} catch (Exception e) {
			// TODO: handle exception
			testNGResult.put("2", new Object[] {"1d", "Navigate to demo website", "Site gets opened", "Fail"});
			Assert.assertTrue(false);
		}
	}
	
	@Test(priority = 1)
	public void fillLoginDetail() {
		try {
			WebElement user = driver.findElement(By.id("login_account"));
			user.sendKeys(username);
			
			WebElement pass = driver.findElement(By.id("login_pwd"));
			pass.sendKeys(password);
			
			Thread.sleep(1000);
			
			testNGResult.put("3", new Object[] {"2d", "Fill login form data(username and password)", "Login form gets filled", "Pass"});
			
		} catch (Exception e) {
			// TODO: handle exception
			testNGResult.put("3", new Object[] {"2d", "Fill login form data(username and password)", "Login form gets filled", "Fail"});
			Assert.assertTrue(false);
		}
	}
	
	@Test(priority = 2)
	public void doLogin() {
		try {
			WebElement submit = driver.findElement(By.className("zidsignin_btn"));
			submit.click();
			
			testNGResult.put("4", new Object[] {"3d", "Click on button login", "Login success", "Pass"});
			
		} catch (Exception e) {
			// TODO: handle exception
			testNGResult.put("4", new Object[] {"3d", "Click on button login", "Login success", "Fail"});
			Assert.assertTrue(false);
		}
	}

	@BeforeClass
	public void suiteSetUp() {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver-win64\\chromedriver.exe");
		
		workBook = new HSSFWorkbook();
		sheet = workBook.createSheet("TestNGResult");

		testNGResult = new LinkedHashMap<String, Object[]>();
		testNGResult.put("1", new Object[] {"Test Step No", "Action", "Expected", "Actual"});
		try {
			driver = new ChromeDriver();
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		} catch (Exception e) {
			// TODO: handle exception
			throw new IllegalStateException("Can't start web driver");
		}
	}

	@AfterClass
	public void afterClass() {
		Set<String> setKey =  testNGResult.keySet();
		int rownum = 0;
		for (String key : setKey) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = testNGResult.get(key);
			int cellnum = 0;
			
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof Date)
					cell.setCellValue((Date) obj);
				if(obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				if(obj instanceof String)
					cell.setCellValue((String) obj);
				if(obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}
		
		try {
			FileOutputStream fos = new FileOutputStream(new File("SaveTestNGResultExcel.xls"));
			workBook.write(fos);
			fos.close();
			System.out.println("Successfully saved selenium webdriver testNG result excel!");
		} catch (FileNotFoundException e) {
			// TODO: handle exception
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
//		driver.close();
//		driver.quit();
	}

}
