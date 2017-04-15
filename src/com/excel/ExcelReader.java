package com.excel;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.exceptions.XLSFatalException;

public abstract class ExcelReader {

	protected String fileName;
	protected Workbook excelWB;
	protected int noOfSheets = 0;
	
	protected Sheet excelWS;
	
	
	protected int[] sheetsToRead = {3, 4};
	
	public ExcelReader() {
		
	}
	
	
	protected void setFileName(String fileName) {
		this.fileName = fileName;
	}
	
	public String getFileName() {
		return this.fileName;
	}
	
	public void setWorkbook(Workbook excelWB) {
		this.excelWB = excelWB;
	}
		
	public void readWorkbooks() {
		for(int intCount = 0; intCount < sheetsToRead.length; intCount++) {
			this.excelWS = this.excelWB.getSheetAt(sheetsToRead[intCount]); 
			
			System.out.println("Sheet names " + this.excelWS.getSheetName());
			
		}
	}
	
	
	public InputStream getFileInputStream(String fileName) throws XLSFatalException{
		InputStream fis = null;
		
		try {
			fis = new FileInputStream(fileName);
		}
		catch(Exception e) {
			
			throw new XLSFatalException(e);
		}
		return fis;
	}
	
	public abstract Workbook getWorkbookObj(InputStream fis) throws XLSFatalException;
}
