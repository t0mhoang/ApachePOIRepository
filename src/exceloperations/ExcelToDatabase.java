package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDatabase {
	
	//This class will import Excel table into Database

	public static void main(String[] args) throws SQLException, IOException {
		
		//Added MySQL connector jar to build path libraries (if you havnt already)

	//Database connection
		Connection connect =DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","MysqlPassword9!"); //localhost/Database,User,Password
		Statement statement =connect.createStatement();
		
	//Create a new table in the database called 'places'....
		//Typically when we create a new table in mysql we use this query below... but we will just put it into a variable to use here:
		//create table places (LOCATION_ID decimal(4,0),
		//STREET_ADDRESS varchar(40),
		//POSTAL_CODE varchar(12),
		//CITY varchar(30),
		//STATE_PROVINCE varchar(25),
		//COUNTRY_ID varchar(2))
		
		String sql = "create table places (LOCATION_ID decimal(4,0), STREET_ADDRESS varchar(40), POSTAL_CODE varchar(12), CITY varchar(30), STATE_PROVINCE varchar(25), COUNTRY_ID varchar(2))";
		statement.execute(sql);
		
	//We already have "locations.xlsx" in our datafiles completed...line above only created the table, now we need the data from our "locations.xlsx" imported into DB	
		FileInputStream fis = new FileInputStream(".\\datafiles\\locations.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet =workbook.getSheet("Locations Data"); //use the name of the actual sheet to get it
		
		int rows = sheet.getLastRowNum(); //NOTE:this gets our columns... if we open locations.xlsx file we see from LOCATION_ID to COUNTRY_ID there are 6 columns.
		
		for(int r = 1; r<=rows; r++) {
		//here we'll get all the data from the rows	
			XSSFRow row =sheet.getRow(r);
			double locId =row.getCell(0).getNumericCellValue();//first cell is decimal... so we have to get numeric cell value and create variable for it
			String streetAdd =row.getCell(1).getStringCellValue();//second is street address it uses varchar so we need to getStringCellValue.. and so on for next lines
			String postalCode =row.getCell(2).getStringCellValue();//inside the file, we had to change entire postal_code column cells to "Text" (use right-click)
			String city =row.getCell(3).getStringCellValue();
			String stateProvince =row.getCell(4).getStringCellValue();//put a SPACE in the null places for file
			String countryId =row.getCell(5).getStringCellValue();
			
		//after we have gotten the data from the sheet like above, now we need to insert it into the DB
			//This is going to be weird... but we need to do it exactly like the sql variable below for it to work... ' symbol then " symbol then +variableHERE+ "symbol 'symbol 
			sql = "insert into places values('"+locId+"','"+streetAdd+"','"+postalCode+"','"+city+"','"+stateProvince+"','"+countryId+"')";
			statement.execute(sql);
			statement.execute("commit");
		}
	//After forloop has grabbed data and inserted it using statement execute and commit... we'll close everything
		workbook.close();
		fis.close();
		connect.close();
		
		System.out.println("Done. We've added our Excel data into our DB! Just refresh world db in myqsl then run select * from places; statement");
	}

}
