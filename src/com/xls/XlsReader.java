package com.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsReader {
	private static HSSFWorkbook readFile(String filename) throws IOException {
	    FileInputStream fis = new FileInputStream(filename);
	    try {
	        return new HSSFWorkbook(fis);		// NOSONAR - should not be closed here
	    } finally {
	        fis.close();
	    }
	}
	public static void main(String[] args) 
    {
        try
        {
        	HSSFWorkbook wb = XlsReader.readFile("CID002515  Carlisle_CMRT_4-20_CIT_01-30-2017.xlsx");
 
        	HSSFSheet sheet = wb.getSheetAt(3);
        	int rows = sheet.getPhysicalNumberOfRows();
			System.out.println("Sheet " + 3 + " \"" + wb.getSheetName(3) + "\" has " + rows
					+ " row(s).");
			for (int r = 0; r < rows; r++) {
				HSSFRow row = sheet.getRow(r);
				if (row == null) {
					continue;
				}
				
				
				System.out.println("\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
				for (int c = 0; c < row.getLastCellNum(); c++) {
					HSSFCell cell = row.getCell(c);
					String value;

					if(cell != null) {
						switch (cell.getCellTypeEnum()) {

							case FORMULA:
								value = "FORMULA value=" + cell.getCellFormula();
								break;

							case NUMERIC:
								value = "NUMERIC value=" + cell.getNumericCellValue();
								break;

							case STRING:
								value = "STRING value=" + cell.getStringCellValue();
								break;

							case BLANK:
								value = "<BLANK>";
								break;

							case BOOLEAN:
								value = "BOOLEAN value-" + cell.getBooleanCellValue();
								break;

							case ERROR:
								value = "ERROR value=" + cell.getErrorCellValue();
								break;

							default:
								value = "UNKNOWN value of type " + cell.getCellTypeEnum();
						}
						System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE="
								+ value);
					}
			}
			}
        }
        catch(IOException e) {}
        catch(Exception e) {
        	e.printStackTrace();
        
        	System.out.println();
        }
    }
	

}
