package com.qa.hs.keyword.engine;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.qa.hs.keyword.base.Base;

public class KeyWordEngine 
{
	
	public WebDriver driver;
	public Properties prop;
	public Base base;
	public WebElement element;
	
	public static Workbook wb;
	public static Sheet sh;
	
	public final String SCENARIO_SHEET_PATH = "C:\\Users\\Priya Borade\\git\\AutomationPriyanka\\TestGit\\src\\main\\java\\com\\qa\\hs\\keyword\\scenarios\\hubspot_Scenarios.xlsx" ;

	public void startExecution(String sheetName)
	{
		String locatorName= null;
		String locatorValue =  null;
		
		FileInputStream file = null;
		
		try 
		{
			file= new FileInputStream(SCENARIO_SHEET_PATH);
		} catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		}
		
		try {
			wb= WorkbookFactory.create(file);
		} catch (EncryptedDocumentException e) {	
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		sh = wb.getSheet(sheetName);
		int k=0;
		
		
		for(int i=0;		i< sh.getLastRowNum();	 i++)
		{
			try{
			String locatorColValue=  sh.getRow(i+1).getCell(k+1).toString().trim(); //id=username
			if(!locatorColValue.equalsIgnoreCase("NA"))
			{
				locatorName = locatorColValue.split("=")[0].trim(); // id
				locatorValue =  locatorColValue.split("=")[1].trim(); //username
			}
			String action= sh.getRow(i+1).getCell(k+2).toString().trim();
			String value= sh.getRow(i+1).getCell(k+3).toString().trim();
			
			switch (action) {
			case "launch browser":
				base = new Base();
				prop= base.init_properties();
					if(value.isEmpty() || value.equals("NA"))
					{
						driver=  base.init_driver(prop.getProperty("browser"));
					}else
					{
						driver = base.init_driver(value);
					}
				break;

			case "enter url":
				if(value.isEmpty() || value.equals("NA"))
				{
					driver.get(prop.getProperty("url"));
				}else
				{
					driver.get(value);
				}
			break;
				
			case "quit":
				driver.quit();
			break;
		default:
				break;
			}
			
			switch (locatorName) {
			case "id":
				element =   driver.findElement(By.id(locatorValue));
				if(action.equalsIgnoreCase("sendkeys")	)
				{ 
					element.clear();
					element.sendKeys(value);
				}else if (action.equalsIgnoreCase("click")) {
					element.click();
				}
				locatorName = null;
				break;

			default:
				break;
			}
			
		} 
			catch(Exception e){
			e.printStackTrace();
			}
		
		
	}
	
}
}
