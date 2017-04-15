package com.reader;

import com.excel.ExcelReader;
import com.excel.HSSFExcelReader;
import com.excel.XSSFExcelReader;
import com.exceptions.XLSFatalException;
import java.io.File;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadCorrectExcel {

    public ExcelReader excelReader;

    public ReadCorrectExcel() {

    }

    public ReadCorrectExcel(ExcelReader excelReader) {
        this.excelReader = excelReader;
    }

    public void read(String fileName) throws XLSFatalException {

        ExcelReader reader = null;
        InputStream fis = null;
        Workbook excelWB = null;
        if (fileName.endsWith("xls")) {
            reader = new HSSFExcelReader();
            fis = reader.getFileInputStream(fileName);
            excelWB = reader.getWorkbookObj(fis);
            reader.setWorkbook(excelWB);
            reader.readWorkbooks();
        } else {
            if (fileName.endsWith("xlsx")) {

                reader = new XSSFExcelReader();
                fis = reader.getFileInputStream(fileName);
                excelWB = reader.getWorkbookObj(fis);
                reader.setWorkbook(excelWB);
                reader.readWorkbooks();
            }
        }

    }

    public static void main(String[] args) {
        ReadCorrectExcel readExcel = new ReadCorrectExcel();
        File resource = new File("");
        File resourceDirectory = new File(resource.toURI());
        File[] excelFiles = resourceDirectory.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.toLowerCase().contains(".xls");
            }
        });
        for (File excelFile : excelFiles) {
            try {
                readExcel.read(excelFile.getName());
            } catch (XLSFatalException ex) {
                Logger.getLogger(ReadCorrectExcel.class.getName()).log(Level.SEVERE, "Cannot Read this type of File", ex);
            }
        }
    }

}
