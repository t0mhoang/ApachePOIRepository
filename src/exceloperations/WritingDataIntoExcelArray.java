package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataIntoExcelArray {
	//Workbook
	//Sheet
	//Row
	//Cell

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Empty Info");
		
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
			empdata.add(new Object[] {"EmpId", "Name", "Job"});
			empdata.add(new Object[] {101, "David", "Engineer"});
			empdata.add(new Object[] {102, "Scott", "Manager"});
			empdata.add(new Object[] {103, "John", "Analyst"});
			
		
		//*******************USING FOR EACH LOOP
		int rowNumber = 0;
		
		for(Object[] data:empdata) {
			
			XSSFRow row = sheet.createRow(rowNumber++);
			
			int cellNumber = 0;
			for(Object value:data) {
				
				XSSFCell cell = row.createCell(cellNumber++);
				
				if(value instanceof String) //Checking if the specific data we have is String
					cell.setCellValue((String)value);//If it's a string then it'll set the specific cell it's in
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
			
		String filePath = "./datafiles/writtenEmployeeArray.xlsx";
		FileOutputStream outputStream = new FileOutputStream(filePath);
		workbook.write(outputStream);

		outputStream.close();
		System.out.println("employee excell sheet using array, successful!");
		
	}

}
