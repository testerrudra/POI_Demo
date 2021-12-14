package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Get_String_Cell_dataTest {

	public static void main(String[] args) throws IOException 
	{
		String x= xlutils.getStringCellData("TestData.xlsx", "LoginData", 1, 1);
		System.out.println(x);

	}

}
