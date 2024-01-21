package com.ranjit.excelsheetreader;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalTime;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class ExcelSheetReader {
	public static void main(String[] args) {

		// Try block to check for exceptions
		try {

			// Reading file from local directory
			FileInputStream file = new FileInputStream(new File("C:/Users/rsuta/Downloads/Documents/friend.xlsx"));

			// Create Workbook instance holding reference to
			// .xlsx file

			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int count = 0;

			// Till there is an element condition holds true
			while (rowIterator.hasNext()) {

				
				
				// for iterating each row one by one
				Row row = rowIterator.next();
				
				// skip the first and Second row in table and if any empty value is present the this condition skip
	
				 if(row.getRowNum()==0 || row.getRowNum()==1 || row.getCell(4) == null){
					   continue;
					  }
				 String temcell="";
				
				// to check the consucative day of an Employee  save the employee id in 
				temcell = row.getCell(0).getStringCellValue();
				
				// for number of hour worked by Employee
				String time= row.getCell(4).getStringCellValue();
				String[] homi = time.split(":");
				
				//convert string  into integer
				int hour=Integer.parseInt(homi[0]);
				// For each row, iterate through all the
				// columns
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {

					// iterating the each cell in a row
					Cell cell = cellIterator.next();
					String t = cell + "";
					// counting the number of consecutive days.
					if (temcell.equals(t)) {
						count++;
					}
				}

				int counter=1;
				//condition for consugative
				if (count == 7 && counter==1 ) {
				
					//number of hour done by Employee
					if(hour>=1 && hour <=10 || hour>=14) {
					
				// printing the Employee name and
					System.out.println(row.getCell(0));
					System.out.println(row.getCell(7));
					counter++;
					count=0;
					}

				}
				System.out.println("");
			}

			// Closing file output streams
			file.close();
		}

		catch (Exception e) {
			e.printStackTrace();
		}

	}

}
