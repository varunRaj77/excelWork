package utills;

import java.io.*;
import java.sql.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Excel {

	public static void main(String[] args) {
		String jdbcURL = "jdbc:mysql://localhost:3306/excelutility";
		String username = "root";
		String password = "";

		String excelFilePath = "./Data/TestData.xlsx";

		int batchSize = 20;

		Connection connection = null;

		try {
			long start = System.currentTimeMillis();

			FileInputStream inputStream = new FileInputStream(excelFilePath);

			Workbook workbook = new XSSFWorkbook(inputStream);

			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = firstSheet.iterator();

			connection = DriverManager.getConnection(jdbcURL, username, password);
			connection.setAutoCommit(false);

			String sql = "INSERT INTO excel (StuID , StuName, FatherName, MotherName, Country, State, City, Pincode, MobileNo, BloodGrp ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
			PreparedStatement statement = connection.prepareStatement(sql);    

			int count = 0;

			rowIterator.next(); // skip the header row

			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					int columnIndex = nextCell.getColumnIndex();

					switch (columnIndex) {
					case 0:
						int StuID = (int) nextCell.getNumericCellValue();
						statement.setInt(1, StuID);
						break;
					case 1:
						String StuName = nextCell.getStringCellValue();
						statement.setString(2, StuName);
						break;
					case 2:	
						String FatherName = nextCell.getStringCellValue();
						statement.setString(3, FatherName);
						break;
					case 3:
						String MotherName = nextCell.getStringCellValue();
						statement.setString(4, MotherName);
						break;
					case 4:
						String Country = nextCell.getStringCellValue();
						statement.setString(5, Country);
						break;	
					case 5:
						String State = nextCell.getStringCellValue();
						statement.setString(6, State);
						break;	
					case 6:
						String City = nextCell.getStringCellValue();
						statement.setString(7, City);
						break;	
					case 7:
						int Pincode = (int) nextCell.getNumericCellValue();
						statement.setInt(8, Pincode);
						break;
					case 8:
						int MobileNo = (int) nextCell.getNumericCellValue();
						statement.setInt(9, MobileNo);
						break;
					case 9:
						String BloodGrp = nextCell.getStringCellValue();
						statement.setString(10, BloodGrp);
						break;	
					}

				}

				statement.addBatch();

				if (count % batchSize == 0) {
					statement.executeBatch();
				}              

			}

			workbook.close();

			// execute the remaining queries
			statement.executeBatch();

			connection.commit();
			connection.close();

			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms\n", (end - start));

		} catch (IOException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (SQLException ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}
}