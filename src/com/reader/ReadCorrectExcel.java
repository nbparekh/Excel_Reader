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
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadCorrectExcel {

    public ExcelReader excelReader;

    public ReadCorrectExcel() {

    }

    public ReadCorrectExcel(ExcelReader excelReader) {
        this.excelReader = excelReader;
    }

    public void read(String fileName) {

        ExcelReader reader = null;
        InputStream fis = null;
        Workbook excelWB = null;
        try {
            reader = new XSSFExcelReader();
            fis = reader.getFileInputStream(fileName);
            excelWB = reader.getWorkbookObj(fis);
            reader.setWorkbook(excelWB);
            reader.readWorkbooks();
        } catch (OLE2NotOfficeXmlFileException ex) {
            try {
                reader = new HSSFExcelReader();
                fis = reader.getFileInputStream(fileName);
                excelWB = reader.getWorkbookObj(fis);
                reader.setWorkbook(excelWB);
                reader.readWorkbooks();
            } catch (XLSFatalException ex1) {
                Logger.getLogger(ReadCorrectExcel.class.getName()).log(Level.SEVERE, new StringBuilder().append(fileName).append(" ").append(ex1.getMessage()).toString());
            }
        } catch (XLSFatalException ex) {
            Logger.getLogger(ReadCorrectExcel.class.getName()).log(Level.SEVERE, ex.getMessage());
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
            readExcel.read(excelFile.getName());
        }
    }

}
