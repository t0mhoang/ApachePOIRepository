package exceloperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaForBooksFile {
	
	//Here total price is missing for "books.xlsx" file, we will add the formula into the file in this class
	
	//Formula will need to be on cell C8 (7th row, 2nd cell):
		//SUM(C2:C6)    <-- this will get us the sum from cells C2 to C6
	

	public static void main(String[] args) throws IOException {
		
	//Setup path of file we need to setup fis to grab the file and import the workbook
		String path = ".\\datafiles\\book.xlsx";
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	//Setup the sheet we need... getSheetAt will get us the sheet index. in this case ' 0 ' is the very first sheet whatever it's name is doesnt matter here.
		XSSFSheet sheet = workbook.getSheetAt(0);
		
	//Setup the formula	
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");
	
	//We've only been LOOKING at the workbook up to this point w the code, so close FIS and get ready to actually open it using FOS to write the formula
		fis.close();
		
	//Setup fos to write our code into Cell C8
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		
	//Now we can close workbook and FOS after writing.
		workbook.close();
		fos.close();
		
		System.out.println("Added and DONE! refresh folder, open the book.xlsx file w/ system editor double-click on C8 cell then press ENTER");
	}

}
