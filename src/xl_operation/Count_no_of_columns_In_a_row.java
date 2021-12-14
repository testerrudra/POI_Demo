package xl_operation;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Count_no_of_columns_In_a_row {

	public static void main(String[] args) throws IOException {
		// Script to Count No. of Columns of a XlSheet Row
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb= new XSSFWorkbook(fi);
		Sheet ws1= wb.getSheet("LoginData");
		Row r=ws1.getRow(0);
		int col_count=r.getLastCellNum();
		System.out.println("column count: "+ col_count);
		
		wb.close();
		fi.close();
	}

}
