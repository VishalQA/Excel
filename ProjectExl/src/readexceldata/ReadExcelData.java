package readexceldata;

import lib.ExcelDataConfig;

public class ReadExcelData {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		
		ExcelDataConfig excel = new ExcelDataConfig("D:\\TestData.xlsx");
		System.out.println(excel.getData(0, 1, 1));
		
	}

}
