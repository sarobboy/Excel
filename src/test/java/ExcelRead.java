import java.io.File;
import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead { //22.6

	public static void main(String[] args) throws Throwable {
		// TODO Auto-generated method stub

		File f = new File("C:\\Users\\sarav\\eclipse-workspace\\Excel\\ExcelSelenium\\src\\test\\resources\\Avengers.xlsx");
		
		FileInputStream f1 = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(f1);
		
		Sheet s = w.getSheet("Sheet1");
		for(int i = 0 ; i<s.getPhysicalNumberOfRows(); i++) {
		Row r = s.getRow(i);
			
		for(int j=0; j<r.getPhysicalNumberOfCells(); j++) {
			Cell c = r.getCell(j);
		
			int celltype = c.getCellType(); //words - 1 / num,date - 0
			if(celltype == 1) {
				String words = c.getStringCellValue();
				System.out.println(words);
			}
			else if (celltype ==0) {
				if(DateUtil.isCellDateFormatted(c)) {
					Date d = c.getDateCellValue();
					SimpleDateFormat sd = new SimpleDateFormat ("MM/dd/yyyy");
					String value = sd.format(d); //change from date to string
					System.out.println(value);
				}
				else {
				double d =	c.getNumericCellValue();
				// double can't be converted to string so, changing/Typecasting it to long
				Long l = (long)d;
				//using string.valueOf(l)>> long is converted to string
				String num = String.valueOf(l);
				System.out.println(num);
				}
			}
			
			}
		}
	}

}
