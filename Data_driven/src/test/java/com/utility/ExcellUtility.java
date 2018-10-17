package com.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcellUtility {
	
	static XSSFWorkbook wb;
	static XSSFSheet sh;
	
	
	public  ExcellUtility (String excelpath)
	{
		
		try
		{
			
			File source = new File(excelpath);
			FileInputStream inputFile = new FileInputStream(source);
			
			wb= new XSSFWorkbook(inputFile);
			
		}
		
		catch(Exception e) {
			
			System.out.println(e.getMessage());
			
			
		}
	}	
	
	
		public String getCellData( String sheetName, int rowNumber, int colNumber)
		{
			
			sh = wb.getSheet(sheetName);
	    String 	data =	sh.getRow(rowNumber).getCell(colNumber).getStringCellValue();
		return data;
			
		}
		
		public void setCellData(String sheetName, int rowNumber, int colNumber, String status)
		{
			
			
			sh = wb.getSheet(sheetName);
			Row r1 = sh.getRow(rowNumber);
			Cell c1 = r1.getCell(colNumber);
			
			if(c1==null)
			{
				c1 = r1.createCell(colNumber);
				c1.setCellValue(status);
			}
			
			else
			{
				
				c1.setCellValue(status);
			}
			
		}
			
	
		public int getNumberOfRows(String sheetName)
		{
			sh = wb.getSheet(sheetName);
			
			int totalRows = sh.getLastRowNum()+1;
			System.out.println(totalRows);
			return totalRows;
		}
		
		public int getNumberOfCos(String sheetName)
		{
			sh = wb.getSheet(sheetName);
			int totalcols = sh.getRow(0).getLastCellNum();
			System.out.println("Total coumns are:" + totalcols);
			return totalcols;
		}
		
		public static void writeAndSaveExcel(String excelpath)
		{
			
			try
			{
				File destination = new File(excelpath);
				FileOutputStream outputFile = new FileOutputStream(destination);
				
				wb.write(outputFile);
				outputFile.flush();
				outputFile.close();
			}
			
			catch(Exception e)
			{
				
				System.out.println(e.getMessage());
			}
				
			
			
		}
	}




