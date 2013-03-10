package au.com.asx;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class StockPriceSearchResult {
	
	@Test
	public void priceSearch(){
		
		WebDriver driver = new FirefoxDriver();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get("http://www.asx.com.au/");
		
		driver.findElement(By.id("ASXCodes")).sendKeys("CBA");
		driver.findElement(By.className("pricesearchbutton")).click();
		
		String headerXpath = null;
		String dataXpath = null;
		int headercolumnCounter = 0;
		int headerRowCounter = 0;
		int dataColumnCounter = 0;
		int dataRowCounter = 1;
		
		headerXpath = "//table[@class='datatable']/tbody/tr[1]/th";
		dataXpath = "//table[@class='datatable']/tbody/tr[2]/td";
		
		List<WebElement> headerValues = driver.findElements(By.xpath(headerXpath));
		for (WebElement headerValue: headerValues){
			System.out.println(headerValue.getText());
			writeToExcel("Sheet1", headerValue, headercolumnCounter++, headerRowCounter);
		}
		
		List<WebElement> dataValues = driver.findElements(By.xpath(dataXpath));
		for (WebElement dataValue: dataValues){
			//System.out.println(dataValue.getText());
			writeToExcel("Sheet1", dataValue, dataColumnCounter++, dataRowCounter);
		}
	}
	
	public void writeToExcel(String sheetName, WebElement value, int headercolumnCounter, int headerRowCounter){
		
		try {
			FileInputStream fis = new FileInputStream("D:\\MyWorkSpace\\SiteScraper\\StockPriceSearchResult.xls");
			Workbook wb = WorkbookFactory.create(fis);
			Sheet s = wb.getSheet(sheetName);
			
			Row r = s.getRow(headerRowCounter);
			if (r == null)
		        r = s.createRow(headerRowCounter);
			
		    Cell cell = r.getCell(headercolumnCounter);
		    
	        if (cell == null)
		        cell = r.createCell(headercolumnCounter);
	        
		    cell.setCellValue(value.getText());

			FileOutputStream fos = new FileOutputStream("D:\\MyWorkSpace\\SiteScraper\\StockPriceSearchResult.xls");
			wb.write(fos);
			fos.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}