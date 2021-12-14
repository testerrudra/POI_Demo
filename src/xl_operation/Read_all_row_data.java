package xl_operation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_all_row_data {

	public static void main(String[] args) throws IOException {
		// Script to Read all rows of data present in a Xlsheet
		
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb= new XSSFWorkbook(fi);
		Sheet ws=wb.getSheet("EmpData");
		int rowcount=ws.getLastRowNum();
		
		for(int i=1; i<=rowcount;i++)
		{
			Row r= ws.getRow(i);
			
			//short columncount=r.getLastCellNum();
			//for(int j=0;j<=columncount;j++)
			{
				//Cell c=r.getCell(j);
				/*String empno=c.getStringCellValue();
				String empname=c.getStringCellValue();
				double salary=c.getNumericCellValue();
				boolean working=c.getBooleanCellValue();*/
				
				//System.out.println(empno+ "   " +empname+"  "+salary+"  "+working);
				//System.out.println(c);
				//System.out.print(c);
				//System.out.println();
				//wb.close();
				//fi.close();
			}
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

}
