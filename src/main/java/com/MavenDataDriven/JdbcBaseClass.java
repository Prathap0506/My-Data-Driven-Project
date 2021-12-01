package com.MavenDataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.sql.rowset.WebRowSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JdbcBaseClass {

	private void particularValues(String path, int row ,int cell) throws IOException {
			
		File f= new File(path);
		FileInputStream fis = new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(fis);
		Sheet sheet =wb.getSheetAt(0);
		Row row =sheet.getRow(row);
		Cell cell= row.getCell(cell);
		String str=cell.getCellType().toString();
		
				
	}
	
}
