package com.excel;

import com.exceptions.XLSFatalException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class ExcelReader {

    protected String fileName;

    protected void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFileName() {
        return this.fileName;
    }

    public InputStream getFileInputStream(String fileName) throws XLSFatalException {
        InputStream fis = null;

        try {
            fis = new FileInputStream(fileName);
        } catch (FileNotFoundException e) {

            throw new XLSFatalException(e);
        }
        return fis;
    }

    public abstract Workbook getWorkbookObj(InputStream fis) throws XLSFatalException;
}
