package com.adobe.support.felix.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteExcel {

	public static void main(String[] args) throws IOException {
		String excelFilePath = "MavenExcelSheet.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		Workbook workbook = new XSSFWorkbook(inputStream);
		Sheet sheet = workbook.getSheetAt(0);
		
		
		for (int i=0; i <= sheet.getLastRowNum(); ++i ) {
			String formula=String.format("SUM(A%d:B%d)", i+1, i+1);
			sheet.getRow(i).getCell(2).setCellFormula(formula);	
		}
		
		calcCells(1);
		
		inputStream.close();
		
		FileOutputStream outputStream = new FileOutputStream(excelFilePath);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();		
		
		 
			
		}
		public static double calcCells(int i) throws IOException {
			String excelFilePath = "MavenExcelSheet.xlsx";
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
			Workbook workbook = new XSSFWorkbook(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
		
		
			double valA= sheet.getRow(i).getCell(0).getNumericCellValue();
			double valB= sheet.getRow(i).getCell(1).getNumericCellValue();
			double total = valA + valB;
		
		
			System.out.println("Cell A value is: " + valA + "Cell B Value is:" + valB + " Together the sum is: " + total );
		
			return total; 
	}
	

}
