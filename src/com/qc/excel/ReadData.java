package com.qc.excel;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class ReadData {
	public static void main(String[] args) throws BiffException, IOException {
		File file = new File("ReadData.xls");
		Workbook book = Workbook.getWorkbook(file);
		Sheet sheet = book.getSheet("Sheet1");
		int rows=sheet.getRows();
		System.out.println("Number of rows: "+rows);
		int columns =sheet.getColumns();
		System.out.println("Number of columns: "+columns);
		
		for(int i=0;i<rows;i++) {
			for(int j=0;j<columns;j++) {
				Cell cell = sheet.getCell(j, i);
				System.out.println(cell.getContents());
			}
			System.out.println("-----------");
		}
		
	}

}
