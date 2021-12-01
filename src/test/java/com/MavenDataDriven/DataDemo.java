package com.MavenDataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDemo {

	public static void main(String[] args) throws IOException {

		// path of the Excel sheet
		File f = new File("C:\\Users\\Prathap J\\eclipse-workspace\\MavenDataDriven\\DataDriven.xlsx");

		// to read excel sheet

		FileInputStream fis = new FileInputStream(f);

		// work book method

		Workbook wb = new XSSFWorkbook(fis); // upcasting

		// read the data rows (index based)

		Sheet sheetAt = wb.getSheetAt(0);   //1st sheet

		Row row = sheetAt.getRow(2); // get row 

		Cell cell = row.getCell(1);

		// if cell value has string get value or others ---

		CellType cellType = cell.getCellType();

		if (cellType.equals(cellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);

		} else if (cellType.equals(cellType.NUMERIC)) {

			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;
			System.out.println(value);

		}

		//write data ===> create sheet
		
		Sheet createSheet = wb.createSheet("Data");  //if u want update sheet wb.getSheet("Data");
		Row createRow = createSheet.createRow(0);
		Cell createCell = createRow.createCell(0);
		createCell.setCellValue("**Name***");
		
		wb.getSheet("Data").getRow(0).createCell(1).setCellValue("***Password***");
		wb.getSheet("Data").createRow(1).createCell(0).setCellValue("test");
		wb.getSheet("Data").getRow(1).createCell(1).setCellValue("12345");
		
		//File output stream
		FileOutputStream fos= new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		
		System.out.println("Your data created");
		
		
	}

}
