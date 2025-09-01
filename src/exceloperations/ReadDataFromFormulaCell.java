package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromFormulaCell {
	//Workbook
	//Sheet
	//Row
	//Cell

	public static void main(String[] args) throws IOException {

		
		FileInputStream file = new FileInputStream(".\\datafiles\\readformula.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1"); //"Sheet1" is the name of the sheet in the file
		
		int rows =sheet.getLastRowNum();
		int cols =sheet.getRow(0).getLastCellNum();
		
		for(int r = 0; r<=rows ; r++) {
			
			XSSFRow row = sheet.getRow(r);
			
			for(int c =0; c<cols ; c++) {
				
				XSSFCell cell =row.getCell(c);
				
				switch(cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue()); //Regular print instead of printLN
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case FORMULA: //without this case... we wont get the Total column values
					System.out.print(cell.getNumericCellValue()); //same as numeric, only difference is the case
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		file.close();
		

	}

}
