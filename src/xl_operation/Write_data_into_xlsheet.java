package xl_operation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_data_into_xlsheet {

	public static void main(String[] args) throws IOException {
		// Script to write data into XlSheet Cells
		
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb=new XSSFWorkbook(fi);
		Sheet ws=wb.getSheet("LoginData");
		Row r=ws.getRow(1);
		Cell c=r.createCell(2);
		c.setCellValue("pass");
		FileOutputStream fo=new FileOutputStream("TestData.xlsx");
		wb.write(fo);
		wb.close();
		fi.close();
	
	}

}
