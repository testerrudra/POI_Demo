package xl_operation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Handle_nullpointer_exception {

	public static void main(String[] args) throws IOException {
		//Script to handle "null pointer exception" that occurs when no data in XlSheet Cell
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb=new XSSFWorkbook(fi);
		Sheet ws=wb.getSheet("LoginData");
		Row r= ws.getRow(1);
		Cell c=r.getCell(3);
		
		String data;
		try 
		{
			data= c.getStringCellValue();
			System.out.println(c);
			
		} catch (Exception e) 
		
		{
			data="";
			System.out.println("No data available");
		}
		
		wb.close();
		fi.close();
	}

}
