package ExcelRead;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadAndSave {

	public static void main(String[] args) throws Exception {
		
		Class.forName("com.mysql.jdbc.Driver");
		
		Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/excel reading", "root", "root");
		
		PreparedStatement prepareStatement=connection.prepareStatement("insert into appa(Id,Name,City) values(?,?,?)");
		
		Workbook workbook = new XSSFWorkbook(new FileInputStream("SAMPLEE.xlsx"));
		
		Sheet sheet = workbook.getSheetAt(0);
		
		for(int i=1; i<=sheet.getLastRowNum(); i++) {
			
			Row row = sheet.getRow(i);
			
			for(int j=0; j<row.getLastCellNum();j++) {
				prepareStatement.setString(j+1,row.getCell(j).toString());
			}
			prepareStatement.executeUpdate();
		}
		System.out.println("Data Inserted Successfully");
		
	}

}
