package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
	//Create file input stream to read our file w/ file path
		String excelFilePath = ".\\datafiles\\countries.xlsx";
		FileInputStream inputstream = new FileInputStream(excelFilePath);
		
	//We need a WorkBook
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		
	//We need to read sheet from that workbook
		XSSFSheet sheet = workbook.getSheet("Sheet1");
	//Or you can use index like this, to read the first sheet..
//		XSSFSheet sheet = workbook.getSheetAt(0);
		
		
//	//*******************************USING FOR LOOP
//		int rows = sheet.getLastRowNum(); //basically counts all rows for us, all the way to last row
//		int columns = sheet.getRow(1).getLastCellNum(); //gets All cells in the row we want, in this case.. it is row 1
//		
//		for(int r = 0; r<=rows; r++) {
//			
//			XSSFRow row = sheet.getRow(r); //this first for loop will go through every row for us when it loops
//			
//			for(int c = 0 ; c<columns ; c++) {
//				
//				XSSFCell cell =row.getCell(c); //this will return the cell from inner loop
//				
//				switch(cell.getCellType()) {//we're using this Switch case for depending on what the cell data type is
//				
//				case STRING: System.out.print(cell.getStringCellValue()); 
//				break;
//				
//				case NUMERIC: System.out.print(cell.getNumericCellValue());
//				break;
//				
//				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
//				break;
//				}
//				System.out.print(" | "); //use this after switch case to seperate the values.. if not each row will be connected to eachother. make sure it's print and not printLN
//			}
//			//After finishing inner for loop
//			System.out.println();//doing this will print next row on next line since this is a print ln
//		}
		
		
	//************************** USING ITERATOR
		
		Iterator iterator = sheet.iterator();
		
		while(iterator.hasNext()) {
			
			XSSFRow row =(XSSFRow) iterator.next();
			
			Iterator cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext()) {
				
				XSSFCell cell = (XSSFCell) cellIterator.next();
				switch(cell.getCellType()) {
				case STRING: System.out.print(cell.getStringCellValue());
				break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());
				break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
				break;
				}
				System.out.print("  |  "); // do this to seperate each value or else each row value will be connected to eachother
			}
			System.out.println();//do this here so each row will go to next line
		}
		

	}

}
