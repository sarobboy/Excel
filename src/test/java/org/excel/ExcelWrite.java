package org.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws Throwable {
		// TODO Auto-generated method stub
		
File f = new File("C:\\Users\\sarav\\eclipse-workspace\\Excel\\ExcelSelenium\\src\\test\\resources\\AvengersEnd.xlsx");
		
		Workbook w = new XSSFWorkbook();
		Sheet s1 = w.createSheet("Excel");
		Row r = s1.createRow(0);
		Cell c = r.createCell(0);
		c.setCellValue("I am Inevitable");
		
		FileOutputStream f1 = new FileOutputStream(f);
		//to write in the workbook
		w.write(f1);
		
		
		
		
		
		
		

	}

}
