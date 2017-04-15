package com.excel;

import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.exceptions.XLSFatalException;

public class HSSFExcelReader extends ExcelReader {
	
	public  HSSFExcelReader() {
		
	}
	
	public void setWorksheets() {
		
	}
	
	
	
	
	public Workbook getWorkbookObj(InputStream fis) throws XLSFatalException {
		
		Workbook hssfWorkbook  = null;
		
		try {
			hssfWorkbook = new HSSFWorkbook(fis);
		}catch(Exception e) {
			System.out.println("Exception while reading HSSF Format  : " + e);
			// to close the fis object here
			throw new XLSFatalException("Exception while reading HSSF Format : ", e);
		}
		
		
		return hssfWorkbook;
	}
	
}
