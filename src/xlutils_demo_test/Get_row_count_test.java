package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Get_row_count_test {

	public static void main(String[] args) throws IOException 
	{
		int rowcount= xlutils.getRowCount("TestData.xlsx", "LoginData");
		System.out.println(rowcount);	
	}

}
