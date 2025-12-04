 package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		//path of the excel sheet
		File f = new File(System.getProperty("user.dir")+"\\excelSheets\\datas.xlsx");
		//convert file to object 
		FileInputStream fi = new FileInputStream(f);
		//get the workbook
		Workbook wb = new XSSFWorkbook(fi); 
		
		//get the sheet 
		Sheet sheet = wb.getSheet("login");
		
		int rowcount = sheet.getPhysicalNumberOfRows(); //1<4, 
		 
		for(int i =1; i<rowcount;i++) //1<4, 2<4
		{
			Row row = sheet.getRow(i);//1 row; 2 row
			int cellcount = row.getPhysicalNumberOfCells();// cell value =2
			for(int j =0; j<cellcount; j++)//0<2; 1<2, 2<2
			{
				Cell cell = row.getCell(j); //0th cell, 1st cell
				DataFormatter format = new DataFormatter();
				String data = format.formatCellValue(cell);
				System.out.println(data + "");
			}
			
		}
		wb.close();
		//get the row
	//	Row row = sheet.getRow(1);
		//get the cell
		//Cell cell = row.getCell(1);
		//get the data
		//String data = cell.getStringCellValue();
	//	DataFormatter format = new DataFormatter();
	//	String data = format.formatCellValue(cell);
	//	System.out.println(data);
		//close the wb
	//	wb.close();
		
		//WebDriver driver = new ChromeDriver();
		//driver.get("https://www.google.com/");
		//driver.findElement(By.xpath("//textarea[@title='Search']")).sendKeys(data);
		

	}

}
