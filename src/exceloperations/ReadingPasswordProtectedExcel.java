package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingPasswordProtectedExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		String path = ".\\datafiles\\passwordProtected.xlsx";
		FileInputStream fis = new FileInputStream(path);
		String password = "opensaysme";
		
		//Normally we could just use line number 22, but because file is password protected we have to use starting line 24
//		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFWorkbook workbook = (XSSFWorkbook)WorkbookFactory.create(fis, password);
		XSSFSheet sheet = workbook.getSheetAt(0);

		
		
/*	//Read data from sheet using FOR LOOP
//		//getting number of rows
//		int rows = sheet.getLastRowNum();
//		System.out.println(rows);//we see the doc has 5 rows (started from 0)
//		
//		//getting number of columns
//		int cols = sheet.getRow(0).getLastCellNum();
//		System.out.println(cols);//we see the doc has 3 columns (started from 1)
//		
//		for(int r = 0; r<=rows; r++) {
//			
//			XSSFRow row= sheet.getRow(r);
//				for(int c = 0; c<cols; c++) {
//				
//				XSSFCell cell = row.getCell(c);
//				
//				switch(cell.getCellType()) {		
//				case NUMERIC: System.out.print(cell.getNumericCellValue());
//				break;
//				case STRING: System.out.print(cell.getStringCellValue());
//				break;
//				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
//				break;
//				case FORMULA: System.out.print(cell.getNumericCellValue());
//				break;
//					}
//				System.out.print(" | ");
//				}
//			System.out.println();
		} */

		
	//Read data from sheet using ITERATOR
		Iterator<Row> iterator = sheet.iterator();
		
		while(iterator.hasNext()) {
			
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator =nextRow.cellIterator();
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					
					switch(cell.getCellType()) {
					case NUMERIC: System.out.print(cell.getNumericCellValue());
					break;
					case STRING: System.out.print(cell.getStringCellValue());
					break;
					case BOOLEAN: System.out.print(cell.getStringCellValue());
					break;
					case FORMULA: System.out.print(cell.getNumericCellValue());
					break;
					}
					System.out.print(" | ");
				}
				System.out.println();
		}
		
		workbook.close();
		fis.close();

	}

}
