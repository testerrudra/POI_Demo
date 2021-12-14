package xl_operation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Count_No_of_row_xl {

	public static void main(String[] args) throws IOException  {
		// Script to Count No. of Rows in a XlSheet
	
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		Sheet ws1 = wb.getSheet("LoginData");
		Sheet ws2 = wb.getSheet("EmpData");
		
		int sheet1_rowcount =  ws1.getLastRowNum();
		int sheet2_rowcount =  ws2.getLastRowNum();
		
		System.out.println("Sheet1 No. of Rows: "+ sheet1_rowcount);
		System.out.println("Sheet2 No. of Rows: "+ sheet2_rowcount);
		
		
		wb.close();
		fi.close();
	}

}
