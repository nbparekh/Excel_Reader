package com.excel;

import com.exceptions.XLSFatalException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFExcelReader extends ExcelReader {

    public XSSFExcelReader() {

    }

    @Override
    public Workbook getWorkbookObj(InputStream fis) throws XLSFatalException {
        Workbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            // to close the fis object here
            throw new XLSFatalException("Exception while reading XSSF Format : ", e);
        }
        return xssfWorkbook;
    }
}
