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

public class ExcelCreate {

	public static void main(String[] args) throws IOException {
		//path of the excel sheet
				File f = new File(System.getProperty("user.dir")+"\\excelSheets\\datas.xlsx");
				//convert file to object 
				FileInputStream fi = new FileInputStream(f);
				//get the workbook
				Workbook wb = new XSSFWorkbook(fi); 
				Sheet sheet = wb.createSheet("demo");
				Row row = sheet.createRow(3);
				Cell cell = row.createCell(3);
				cell.setCellValue("World");
				FileOutputStream fo = new FileOutputStream(f);
				wb.write(fo);
				wb.close();
				System.out.println("Completed");

	}

}
