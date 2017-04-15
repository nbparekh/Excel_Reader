package com.reader;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

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
		
		ExcelReader hssfReader = null, xssfReader = null;
		InputStream fis = null;
		boolean isError = false;
		
		try {
			xssfReader = new XSSFExcelReader();
			fis = xssfReader.getFileInputStream(fileName);
			xssfReader.getWorkbookObj(fis);
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
		System.out.println("Coming here " + isError);
		if(isError) {
			try {
				
				hssfReader = new HSSFExcelReader();
				fis = hssfReader.getFileInputStream(fileName);
				hssfReader.getWorkbookObj(fis);
				System.out.println("read");
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
		
		System.out.println("Done");
		
		
	}
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ReadCorrectExcel readExcel = new ReadCorrectExcel();
		String fileName = null; 
		try { 
			//fileName = "4.20 DB ROBERTS COMPANY LEVEL CONFLICT MINERAL 4.20 TEMPLATE 1-17-17.xls"; // hssf
			//fileName = "American Electro CFSI CMRT 4-20.xlsx"; // xssf
			fileName = "CID002515  Carlisle_CMRT_4-20_CIT_01-30-2017.xlsx"; // cannot be read by both
			readExcel.read(fileName);
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		

	}

}
