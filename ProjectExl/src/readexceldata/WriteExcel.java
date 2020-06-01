package readexceldata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		File src = new File ("D:\\TestData.xlsx");
		
        FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		XSSFSheet sheet1 = wb.getSheetAt(0);
		
		sheet1.getRow(0).createCell(2).setCellValue("Pass1");
		sheet1.getRow(1).createCell(2).setCellValue("Fail1");
		
		FileOutputStream fout = new FileOutputStream(src);
		
		
		wb.write(fout);
		wb.close();
		
		System.out.println("The test data is updated in the excelsheet successefully");
		
	}

}
