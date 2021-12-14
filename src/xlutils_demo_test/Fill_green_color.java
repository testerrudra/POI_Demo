package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Fill_green_color {

	public static void main(String[] args) throws IOException {
		
		xlutils.fillGreenColor("TestData.xlsx", "EmpData", 1, 4);
	}

}
