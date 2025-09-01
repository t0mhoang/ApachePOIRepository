package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell {

	public static void main(String[] args) throws IOException {
	//Set up workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet =workbook.createSheet("Numbers");
	//Row variable
		XSSFRow row =sheet.createRow(0);
	//Create n Add values for cells in first row	
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
	//in same Row we're	writing formula for last cell in the row
		row.createCell(3).setCellFormula("A1 * B1 * C1");
	//Set up where the file will be saved, excute the code w/ write class, then close.	
		FileOutputStream fos = new FileOutputStream(".\\datafiles\\writtenFormula.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("Formula Excel file w calculation created!..");
		
	}

}
