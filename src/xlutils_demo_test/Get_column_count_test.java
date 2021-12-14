package xlutils_demo_test;

import java.io.IOException;

import demo_excel.xlutils;

public class Get_column_count_test {

	public static void main(String[] args) throws IOException {
		
		 int x= xlutils.getColumnCount("TestData.xlsx", "LoginData", 1);
		 System.out.println(x);
	}

}
