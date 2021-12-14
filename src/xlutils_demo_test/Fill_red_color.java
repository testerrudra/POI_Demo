package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Fill_red_color {

	public static void main(String[] args) throws IOException 
	{
		xlutils.fillRedColor("TestData.xlsx", "LoginData", 1, 2);

	}

}
