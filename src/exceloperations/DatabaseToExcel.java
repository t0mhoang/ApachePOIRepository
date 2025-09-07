package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DatabaseToExcel {

	//IN THIS CLASS... we're going to create an excel sheet that will contain all information from specific table in our database
	
	public static void main(String[] args) throws SQLException, IOException {

		//Added MySQL connector jar to build path libraries
		//File: city.xlsx
		
	//Connect to Database
		Connection connect =DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","MysqlPassword9!"); //connect SQL localhost/Database, User, Password
		
	//statment/query
		Statement statement =connect.createStatement();
		ResultSet rs =statement.executeQuery("select * from city"); //this query will get All data from entire city table in the "world" database (Every column and row)
		
	//Excel created for the database
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet =workbook.createSheet("City data"); //created sheet named "City data" because of the query we're going to look at the "city" table in the DB

	//create columns in excel for columns we have in our city table we'll name them exactly the same as columns from DB
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("ID");
		row.createCell(1).setCellValue("Name");
		row.createCell(2).setCellValue("CountryCode");
		row.createCell(3).setCellValue("District");
		row.createCell(4).setCellValue("Population");
		
		int r = 1;
		while(rs.next()) { 
		//go to workbench and check "Type" of data in city table by running "describe city;" query... check the Type column... example:ID uses int... so our first rs will be int
			double cityId =rs.getDouble("ID");
			String name = rs.getString("Name");
			String countryCode = rs.getString("CountryCode");
			String district = rs.getString("District");
			double population = rs.getDouble("Population");
			
			row = sheet.createRow(r++);
			
			row.createCell(0).setCellValue(cityId);
			row.createCell(1).setCellValue(name);
			row.createCell(2).setCellValue(countryCode);
			row.createCell(3).setCellValue(district);
			row.createCell(4).setCellValue(population);
			
		}
		
	//Now write it into excel and close everything including connection
		FileOutputStream fos = new FileOutputStream(".\\datafiles\\city.xlsx");
		workbook.write(fos);
		
		workbook.close();
		fos.close();
		connect.close();
		
		
		System.out.println("Done. Refresh datafiles and open city.xlsx to see ALL data from world database.city table!");
	}

}
