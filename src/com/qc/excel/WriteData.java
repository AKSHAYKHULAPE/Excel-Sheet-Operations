package com.qc.excel;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteData {

	public static void main(String[] args) throws IOException, RowsExceededException, WriteException {
		File fis =new File("writeData.xls");
		WritableWorkbook book =Workbook.createWorkbook(fis);
		WritableSheet sheet = book.createSheet("QueueCodes", 0);
		
		Label col = new Label(0, 0, "Test Number");
		sheet.addCell(col);
		
		Label col1 = new Label(1, 0, "Test Name");
		sheet.addCell(col1);
		
		Label col2 = new Label(2, 0, "Test Result");
		sheet.addCell(col2);
		
		//Label col3 = new Label(1, 0, "");
		//sheet.addCell(col3);
		//case 1:
		Number num = new Number(0, 1, 1);
		sheet.addCell(num);
		
		Label test = new Label(1, 1, "Login Test");
		sheet.addCell(test);
		
		Label testResult = new Label(2, 1, "passed");
		sheet.addCell(testResult);
		
		//Case : 2
		Number num2 = new Number(0, 2, 2);
		sheet.addCell(num2);
		
		Label test2 = new Label(1, 2, "Registration Test");
		sheet.addCell(test2);
		
		Label testResult2 = new Label(2, 2, "failed");
		sheet.addCell(testResult2);
		
		book.write();
		System.out.println("Data is Written......!!!");
		book.close();


	}

}
