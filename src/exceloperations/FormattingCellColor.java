package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormattingCellColor {

	//Workbook
	//Sheet
	//Row
	//Cell
	
	public static void main(String[] args) throws IOException {
		//in this project we'll create everything ourselves (we'll create two cells, one example for BG the other for Foreground)

	//Create the workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
	//Create sheet, and name it "Sheet1"
		XSSFSheet sheet =workbook.createSheet("Sheet1");
	//Create row	, we'll just add 1 row
		XSSFRow row =sheet.createRow(1);
		
			
		//Before the cell we have to set properties to cell.. so lets set start w/ BACKGROUND color for our first
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.BLUE_GREY.getIndex()); //key word is "IndexedColors" once we type that and put ' . ' it will show us list of options
		style.setFillPattern(FillPatternType.DIAMONDS); //just like above, key word is "FillPatternType" once we type and put a ' . ' we'll get a list of options (eg. DIAMONDS)
		
	//Create cell, we'll also just add 1 cell saying "Welcome"
		XSSFCell cell =row.createCell(1);
		cell.setCellValue("Welcome");
		//Setting cell style to our created cell from line 34
		cell.setCellStyle(style);
		
		//now we'll do our FORGROUND cell
		style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.MAROON.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell = row.createCell(2);
		cell.setCellValue("Automation");
		cell.setCellStyle(style);
		
		
	//Now write everything into our workbook, closing workbook AND fos afterwards to finish out our job
		FileOutputStream fos = new FileOutputStream(".\\datafiles\\stylez.xslx");
		workbook.write(fos);
		workbook.close();
		fos.close();
		
		System.out.println("Creation Complete!");

	}

}
