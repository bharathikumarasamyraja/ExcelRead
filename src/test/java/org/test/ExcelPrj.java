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
		
		Sheet sheet2 = wb.createSheet("Personal Details");
		//row0
		Row row4 = sheet2.createRow(0);
		Cell cell6 = row4.createCell(0);
		cell6.setCellValue("First Name");
		
		//row1
		Row row5 = sheet2.createRow(1);	
		Cell cell22 = row5.createCell(0);
		cell22.setCellValue("Bharathi");
		
		
		
		//row1 Personal Details
		
		
		
		Cell cell24 = row5.createCell(1);
		cell24.setCellValue("Kumar");
		
		
		Cell cell25 = row5.createCell(2);
		cell25.setCellValue("Raja");
		
		
		Cell cell26 = row5.createCell(3);
		cell26.setCellValue("Bharathi");
		
		
		Cell cell27 = row5.createCell(4);
		cell27.setCellValue("123456");
		
		
		Cell cell28 = row5.createCell(5);
		cell28.setCellValue("20-11-2030");
		
		
		Cell cell29 = row5.createCell(6);
		cell29.setCellValue("1234567");
		
		
		Cell cell30 = row5.createCell(7);
		cell30.setCellValue("123456789");
		
		
		Cell cell31 = row5.createCell(8);
		cell31.setCellValue("Indian");
		
		
		Cell cell32 = row5.createCell(9);
		cell32.setCellValue("Married");
		
		
		Cell cell33 = row5.createCell(10);
		cell33.setCellValue("30-11-1990");
		
		
		Cell cell34 = row5.createCell(11);
		cell34.setCellValue("Female");
		
		
		Cell cell35 = row5.createCell(12);
		cell35.setCellValue("yes");
		
		
		Cell cell36 = row5.createCell(13);
		cell36.setCellValue("yes");
		
		
		Cell cell37 = row5.createCell(14);
		cell37.setCellValue("A+");
		
//=====================================================================	
		
		//row2 personal details
		
				Row row6 = sheet2.createRow(2);
				Cell cell23 = row6.createCell(0);
				cell23.setCellValue("Samrithi");
				
				Cell cell38 = row6.createCell(1);
				cell38.setCellValue("Bala");
				
				Cell cell39 = row6.createCell(2);
				cell39.setCellValue("karthi");
				
				Cell cell40 = row6.createCell(3);
				cell40.setCellValue("Samrithi");
				
				Cell cell41 = row6.createCell(4);
				cell41.setCellValue("258963");
				
				Cell cell42 = row6.createCell(5);
				cell42.setCellValue("21-11-2030");
				
				Cell cell43 = row6.createCell(6);
				cell43.setCellValue("789456");
				
				Cell cell44 = row6.createCell(7);
				cell44.setCellValue("123456789");
				
				Cell cell45 = row6.createCell(8);
				cell45.setCellValue("Indian");
				
				Cell cell46 = row6.createCell(9);
				cell46.setCellValue("Single");
				
				Cell cell47 = row6.createCell(10);
				cell47.setCellValue("29-03-2023");
				
				Cell cell48 = row6.createCell(11);
				cell48.setCellValue("Female");
				
				Cell cell49 = row6.createCell(12);
				cell49.setCellValue("No");
				
				Cell cell50 = row6.createCell(13);
				cell50.setCellValue("No");
				
				Cell cell51 = row6.createCell(14);
				cell51.setCellValue("B+");
	//=================================================================================			
				
	//Header created	
		
		//---------------------------------------------------------------------
		Cell cell7 = row4.createCell(1);
		cell7.setCellValue("Middle Name");
		
		Cell cell8 = row4.createCell(2);
		cell8.setCellValue("Last Name");
		
		Cell cell9 = row4.createCell(3);
		cell9.setCellValue("Nick Name");
		
		Cell cell10 = row4.createCell(4);
		cell10.setCellValue("Driver License ID");
		
		Cell cell11 = row4.createCell(5);
		cell11.setCellValue("License Expiry Date");
		
		Cell cell12 = row4.createCell(6);
		cell12.setCellValue("SSN number");
		
		Cell cell13 = row4.createCell(7);
		cell13.setCellValue("SIN number");
		
		Cell cell14 = row4.createCell(8);
		cell14.setCellValue("Nationality");
		
		Cell cell15 = row4.createCell(9);
		cell15.setCellValue("Marital Status");
		
		Cell cell16 = row4.createCell(10);
		cell16.setCellValue("Date of Birth");
		
		Cell cell17 = row4.createCell(11);
		cell17.setCellValue("Gender");
		
		Cell cell18 = row4.createCell(12);
		cell18.setCellValue("Military Service");
		
		Cell cell19 = row4.createCell(13);
		cell19.setCellValue("Smoker");
		
		Cell cell20 = row4.createCell(14);
		cell20.setCellValue("Blood Type");
		
		Cell cell21 = row.createCell(15);
		cell21.setCellValue("Test Field");
	//---------------------------------------------------------------------------------	
	
			
		FileOutputStream fo = new FileOutputStream(f);//file object created
		wb.write(fo);
		wb.close();
		System.out.println("Completed");

	}

}
