package xl_operation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_data_from_a_row {

	public static void main(String[] args) throws IOException {
		// Script to Read data from XlSheet Cells of a particular row
		
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb= new XSSFWorkbook(fi);
		Sheet ws=wb.getSheet("EmpData");
		Row r = ws.getRow(1);
		
		Cell c1,c2, c3,c4;
		c1= r.getCell(0);
		c2= r.getCell(1);
		c3= r.getCell(2);
		c4= r.getCell(3);
		
		String empno=c1.getStringCellValue();
		String empname=c2.getStringCellValue();
		double salary=c3.getNumericCellValue();
		boolean working=c4.getBooleanCellValue();
		
		System.out.println(empno+ "   " +empname+"  "+salary+"  "+working);
		wb.close();
		fi.close();
	}

}
