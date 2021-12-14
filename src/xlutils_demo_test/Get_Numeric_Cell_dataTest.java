package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Get_Numeric_Cell_dataTest {

	public static void main(String[] args) throws IOException 
	{
		double x=xlutils.getNumericCellData("TestData.xlsx", "EmpData", 2, 2);
		System.out.println(x);
	}

}
