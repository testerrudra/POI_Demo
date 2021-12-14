package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Get_Boolean_Cell_dataTest {

	public static void main(String[] args) throws IOException 
	{
	  boolean x=  xlutils.getBooleanCellData("TestData.xlsx", "EmpData", 7, 3);
	  System.out.println(x);

	}

}
