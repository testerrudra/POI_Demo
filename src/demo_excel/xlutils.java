package demo_excel;

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
import org.apache.xmlbeans.impl.xb.xsdschema.Public;

public class xlutils 
{
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell cell;
	public static CellStyle passstyle;
	public static CellStyle failstyle;
	public static int getRowCount(String xlfile, String xlsheet) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		int rowcount= ws.getLastRowNum();
		wb.close();
		fi.close();
		return rowcount;
	}
	
	public static short getColumnCount(String xlfile, String xlsheet, int rownum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		
		short columncount =row.getLastCellNum();
		wb.close();
		fi.close();
		return  columncount;	
	}
	
	public static String getStringCellData(String xlfile, String xlsheet, int rownum, int columnnum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws= wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		
		String data;
		try 
		{
			cell=row.getCell(columnnum);
			data=cell.getStringCellValue();
			
		} catch (Exception e) 
		
		{
			data="";	
		}
		wb.close();
		fi.close();
		return data;
	}
       public static double getNumericCellData(String xlfile, String xlsheet, int rownum, int columnnum) throws IOException
       {
    	   fi=new FileInputStream(xlfile);
    	   wb=new XSSFWorkbook(fi);
    	   ws=wb.getSheet(xlsheet);
    	   row=ws.getRow(rownum);
    	   
    	   double data;
    	   try 
    	   {
    		   cell=row.getCell(columnnum);
    		   data=cell.getNumericCellValue();
    		   
		} catch (Exception e) 
    	   
    	{
			data= 0.0;
		}
    	   wb.close();
    	   fi.close();
    	   return data;   
       }
       
       public static boolean getBooleanCellData(String xlfile, String xlsheet, int rownum, int columnnum) throws IOException
       {
    	   fi=new FileInputStream(xlfile);
    	   wb=new XSSFWorkbook(fi);
    	   ws=wb.getSheet(xlsheet);
    	   row=ws.getRow(rownum);
    	   
    	   boolean data;
    	   try 
    	   {
    		   cell=row.getCell(columnnum);
    		   data=cell.getBooleanCellValue();
    		   
			
		} catch (Exception e) 
    	   {
			data= false;
		}
    	   wb.close();
    	   fi.close();
    	   return data;
       }
       
       public static void setCellData(String xlfile, String Xlsheet, int rownum, int columnnum, String data) throws IOException
       {
    	   fi= new FileInputStream(xlfile);
    	   wb=new XSSFWorkbook(fi);
    	   ws=wb.getSheet(Xlsheet);
    	   row=ws.getRow(rownum);
    	   cell=row.createCell(columnnum);
    	   cell.setCellValue(data);
    	   
    	   fo=new FileOutputStream(xlfile);
    	   wb.write(fo); 
    	   wb.close();
    	   fi.close();
    	   fo.close();   
       }
       
       public static void fillGreenColor(String xlfile, String Xlsheet, int rownum, int colmunnum) throws IOException
       {
    	   fi=new FileInputStream(xlfile);
    	   wb=new XSSFWorkbook(fi);
    	   ws=wb.getSheet(Xlsheet);
    	   row=ws.getRow(rownum);
    	   cell=row.getCell(colmunnum);
    	   passstyle=wb.createCellStyle();
    	   
    	   passstyle.setFillForegroundColor(IndexedColors.GREEN.index);
    	   passstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    	   cell.setCellStyle(passstyle);
    	   fo=new FileOutputStream(xlfile);
    	   wb.write(fo);
    	   wb.close();
    	   fi.close();
    	   fo.close();		   
       }
       
       public static void fillRedColor(String xlfile, String xlsheet, int rownum, int columnnum) throws IOException
       {
    	   fi=new FileInputStream(xlfile);
    	   wb=new XSSFWorkbook(fi);
    	   ws=wb.getSheet(xlsheet);
    	   row=ws.getRow(rownum);
    	   cell=row.getCell(columnnum);
    	   failstyle=wb.createCellStyle();
    	   
    	   failstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
    	   failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    	   cell.setCellStyle(failstyle);
    	   fo=new FileOutputStream(xlfile);
    	   wb.write(fo);
    	   wb.close();
    	   fi.close();
    	   fo.close();
    	 
       }
       
}



