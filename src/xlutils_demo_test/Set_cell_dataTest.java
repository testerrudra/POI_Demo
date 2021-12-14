package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Set_cell_dataTest {

	public static void main(String[] args) throws IOException 
	{
	   xlutils.setCellData("TestData.xlsx", "LoginData", 2, 2, "Pass");
	}

}
