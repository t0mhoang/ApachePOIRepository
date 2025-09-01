package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataIntoExcel {
	//Workbook
	//Sheet
	//Row
	//Cell

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Empty Info"); //we're creating a new excel file sheet, naming the sheet "Empty Info"
		
		Object empdata[][] = {		{"EmpID", "Name", "Job"},
												{101, "David", "Engineer"},
												{102, "John", "Manager"},
												{103, "Scott", "Analyst"}
										};
		
//		//************************USING FOR LOOP
//		int rows = empdata.length;
//		int cols = empdata[0].length;
//		
//		System.out.println(rows); // 4 rows
//		System.out.println(cols); // 3 columns
//		
//		for(int r = 0; r<rows; r++) {
//			
//			XSSFRow row = sheet.createRow(r);
//			
//			for(int c = 0; c<cols; c++) {
//				XSSFCell cell = row.createCell(c);
//				Object value =empdata[r][c]; //Since we created original empdata as an object, the variable we use is also an Object
//				
//				if(value instanceof String)
//					cell.setCellValue((String)value);
//				if(value instanceof Integer)
//					cell.setCellValue((Integer)value);
//				if(value instanceof Boolean)
//					cell.setCellValue((Boolean)value);
//			}
//		}
		
		
		//******************USING FOR EACH LOOP
		int rowCount = 0;
		
		for(Object emp[] : empdata) {
			
			XSSFRow row = sheet.createRow(rowCount++);
			int columnCount = 0;
			
				for(Object value:emp){
				
					XSSFCell cell = row.createCell(columnCount++);
					
					if(value instanceof String)
						cell.setCellValue((String)value);
					if(value instanceof Integer)
						cell.setCellValue((Integer)value);
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value);
			}
		}
		
		
		String filePath = ".\\datafiles\\writtenEmployee.xlsx"; //This is where it's going to get created... we named the actual file "employee.xlsx"
		FileOutputStream outstream = new FileOutputStream(filePath);
		workbook.write(outstream);
		
		
		outstream.close();
		System.out.println("Employee.xlsx file written successfully...");
		
	}

}
