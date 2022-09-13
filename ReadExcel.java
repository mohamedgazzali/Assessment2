package testcase;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {

	public static String[][] ReadExcelData(String filename) throws IOException {
		// TODO Auto-generated method stub
		
		
		XSSFWorkbook book = new XSSFWorkbook("./ExcelData/"+filename+".xlsx");
		
		XSSFSheet sheet = book.getSheetAt(0);
		
		XSSFRow row = sheet.getRow(1);
//		XSSFCell col = row.getCell(2);
//		
//		String stringCellValue = col.getStringCellValue();
//		System.out.println(stringCellValue);
				
		int rowCount=sheet.getLastRowNum();
		
		System.out.println(rowCount);
		
		int colCount = row.getLastCellNum();
		
		System.out.println(colCount);
		String[][] data=new String[rowCount][colCount];
		for(int i=1; i<=rowCount; i++) {
			for(int j=0; j<colCount; j++) {
				String datas = sheet.getRow(i).getCell(j).getStringCellValue();
				System.out.println(datas);
				data[i-1][j]=datas;
				
			}
		}
		book.close();
		return data;

	}

}
