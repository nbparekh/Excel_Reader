package com.excel;

import com.exceptions.XLSFatalException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class HSSFExcelReader extends ExcelReader {

    public HSSFExcelReader() {

    }

    public void setWorksheets() {

    }

    @Override
    public Workbook getWorkbookObj(InputStream fis) throws XLSFatalException {

        Workbook hssfWorkbook = null;

        try {
            hssfWorkbook = new HSSFWorkbook(fis);
        } catch (IOException e) {
            System.out.println("Exception while reading HSSF Format  : " + e);
            // to close the fis object here
            throw new XLSFatalException("Exception while reading HSSF Format : ", e);
        }

        return hssfWorkbook;
    }

}
