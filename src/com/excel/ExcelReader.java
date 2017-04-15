package com.excel;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;

import com.exceptions.XLSFatalException;

public abstract class ExcelReader {

	protected String fileName;
	
	protected void setFileName(String fileName) {
		this.fileName = fileName;
	}
	
	public String getFileName() {
		return this.fileName;
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
