package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPrj {

	public static void main(String[] args) throws IOException {
		//path of the excel sheet
		File f = new File(System.getProperty("user.dir")+"\\excelSheets\\Excelprj.xlsx");
		//convert file to object 
		FileInputStream fi = new FileInputStream(f);
		//get the workbook
		Workbook wb = new XSSFWorkbook(fi); 
		Sheet sheet = wb.createSheet("Login page");
		
		//Row 0
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("UserName");
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Password");
		
		
		
		//row1
		Row row2 = sheet.createRow(1);
		Cell cell2 = row2.createCell(0);
		cell2.setCellValue("Bharathi");
		Cell cell3 = row2.createCell(1);
		cell3.setCellValue("Bharathi@123");
		
		//row2
		Row row3 = sheet.createRow(2);
		Cell cell4 = row3.createCell(0);
		cell4.setCellValue("Admin");
		Cell cell5 = row3.createCell(1);
		cell5.setCellValue("admin123");
		
		
		
		
		
		
		
		FileOutputStream fo = new FileOutputStream(f);
		wb.write(fo);
		wb.close();
		System.out.println("Completed");

	}

}
