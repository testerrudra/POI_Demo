package xl_operation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Fill_color_in_passcell {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream fi= new FileInputStream("TestData.xlsx");
		Workbook wb=new XSSFWorkbook(fi);
		Sheet ws=wb.getSheet("LoginData");
		Row r=ws.getRow(1);
		Cell c=r.createCell(2);
		c.setCellValue("pass");
		
		CellStyle passstyle= wb.createCellStyle();
		passstyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		passstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		c.setCellStyle(passstyle);
		
		FileOutputStream fo=new FileOutputStream("TestData.xlsx");
		wb.write(fo);
		wb.close();
		fi.close();
	

	}

}
