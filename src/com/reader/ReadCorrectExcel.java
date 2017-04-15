package com.reader;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.poi.ss.usermodel.Workbook;

import com.excel.ExcelReader;
import com.excel.HSSFExcelReader;
import com.excel.XSSFExcelReader;
import com.exceptions.XLSFatalException;

public class ReadCorrectExcel {

	public ExcelReader excelReader;
	
	public ReadCorrectExcel() {
		
	}
	
	public ReadCorrectExcel(ExcelReader excelReader) {
		this.excelReader = excelReader;
	}
	
	public void read(String fileName) throws XLSFatalException {
		
		ExcelReader excelReader = null;
		InputStream fis = null;
		Workbook excelWB = null;
		boolean isError = false;
		
		try {
			excelReader = new XSSFExcelReader();
			fis = excelReader.getFileInputStream(fileName);
			excelWB = excelReader.getWorkbookObj(fis);
		}catch(Exception e) {
			e.printStackTrace();
			isError = true;
		}finally {
			try {
				fis.close();
			}catch(IOException e) {
				e.printStackTrace();
			}
		}
		if(isError) {
			try {
				
				excelReader = new HSSFExcelReader();
				fis = excelReader.getFileInputStream(fileName);
				excelWB = excelReader.getWorkbookObj(fis);
			}catch(Exception e) {
				throw new XLSFatalException("Cannot Read this type of File");
			}
			finally {
				try {
					fis.close();
				}catch(IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		excelReader.setWorkbook(excelWB);
		excelReader.readWorkbooks();
		
		
		
		
	}
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ReadCorrectExcel readExcel = new ReadCorrectExcel();
		String fileName = null; 
		try { 
			fileName = "4.20 DB ROBERTS COMPANY LEVEL CONFLICT MINERAL 4.20 TEMPLATE 1-17-17.xls"; // hssf
			//fileName = "American Electro CFSI CMRT 4-20.xlsx"; // xssf
			//fileName = "CID002515  Carlisle_CMRT_4-20_CIT_01-30-2017.xlsx"; // cannot be read by both
			readExcel.read(fileName);
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		

	}

}
