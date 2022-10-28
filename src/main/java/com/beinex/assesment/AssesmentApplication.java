package com.beinex.assesment;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Date;
import java.util.Iterator;

@SpringBootApplication
public class AssesmentApplication {

	public static void main(String[] args) {
		SpringApplication.run(AssesmentApplication.class, args);
		String jdbcURL = "jdbc:mysql://localhost:3306/beinex_task";
		String username = "root";
		String password = "root";

		String excelFilePath = "Sample - Superstore_Orders.xlsx";

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

			String sql = "INSERT INTO beinex_task (category,city,country,customer,manufacturer,orderDate,orderId,postalCode,productName,region,segment,shipDate,shipMode,province,subCategory,discount,profit,quantity,sales) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)" ;
			PreparedStatement statement = connection.prepareStatement(sql);

			int count = 0;

			rowIterator.next();

			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					int columnIndex = nextCell.getColumnIndex();

					switch (columnIndex) {
						case 0:
							String category = nextCell.getStringCellValue();
							statement.setString(1, category);
							break;
						case 1:
							String city = nextCell.getStringCellValue();
							statement.setString(2, city);
						case 2:
							String country = nextCell.getStringCellValue();
							statement.setString(3, country);
						case 3:
							String customer = nextCell.getStringCellValue();
							statement.setString(4, customer);
						case 4:
							String manufacturer = nextCell.getStringCellValue();
							statement.setString(5, manufacturer);
						case 5:
							Date enrollDate = nextCell.getDateCellValue();
							statement.setTimestamp(6, new Timestamp(enrollDate.getTime()));
						case 6:
							String name = nextCell.getStringCellValue();
							statement.setString(7, name);

						case 7:
							Date orderDate = nextCell.getDateCellValue();
							statement.setTimestamp(8, new Timestamp(orderDate.getTime()));
						case 8:
							String orderId = nextCell.getStringCellValue();
							statement.setString(9, orderId);
						case 9:
							int postalCode = (int) nextCell.getNumericCellValue();
							statement.setInt(10, postalCode);

						case 10:
							String productName = nextCell.getStringCellValue();
							statement.setString(11, productName);
						case 11:
							String region = nextCell.getStringCellValue();
							statement.setString(12, region);
						case 12:
							String segment = nextCell.getStringCellValue();
							statement.setString(13, segment);

						case 13:
							Date shipDate = nextCell.getDateCellValue();
							statement.setTimestamp(14, new Timestamp(shipDate.getTime()));
						case 14:
							String shipMode = nextCell.getStringCellValue();
							statement.setString(15, shipMode);

						case 15:
							String province = nextCell.getStringCellValue();
							statement.setString(16, province);
						case 16:
							String subCategory = nextCell.getStringCellValue();
							statement.setString(17, subCategory);
						case 17:
							float discount = (float) nextCell.getNumericCellValue();
							statement.setFloat(18, discount);
						case 18:
							float profit = (float) nextCell.getNumericCellValue();
							statement.setFloat(19, profit);
						case 19:
							float quantity = (float) nextCell.getNumericCellValue();
							statement.setFloat(20, quantity);
						case 20:
							float sales = (float) nextCell.getNumericCellValue();
							statement.setFloat(21, sales);



					}

				}

				statement.addBatch();

				if (count % batchSize == 0) {
					statement.executeBatch();
				}

			}

			workbook.close();

			statement.executeBatch();

			connection.commit();
			connection.close();

			long end = System.currentTimeMillis();

		} catch (IOException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (SQLException ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}


}
