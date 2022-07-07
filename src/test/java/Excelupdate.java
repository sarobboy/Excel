import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelupdate {

	public static void main(String[] args) throws Throwable {
		// TODO Auto-generated method stub

		File f = new File("C:\\Users\\sarav\\eclipse-workspace\\Excel\\ExcelSelenium\\src\\test\\resources\\AvengersEnd.xlsx");
		FileInputStream f2 = new FileInputStream(f);
		Workbook w = new XSSFWorkbook (f2);
		Sheet s = w.getSheet("Excel");
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		int celltype = c.getCellType();
		if(celltype == 1) {
			String value = c.getStringCellValue();
			if(value.equals("I am Inevitable")) {
				c.setCellValue("I am IronMan");
			}
		}
		FileOutputStream f1 = new FileOutputStream(f);
		w.write(f1);
		
		
		
		
		
		
		
		

	}
	

}
