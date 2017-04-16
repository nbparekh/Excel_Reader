package com.excel;

import com.exceptions.XLSFatalException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class ExcelReader {

    protected String fileName;
    protected Workbook excelWB;
    protected int noOfSheets = 0;

    protected Sheet excelWS;

    protected int[] sheetsToRead = {3, 4};

    public ExcelReader() {
    }

    protected void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFileName() {
        return this.fileName;
    }

    public void setWorkbook(Workbook excelWB) {
        this.excelWB = excelWB;
    }

    public void readWorkbooks() throws XLSFatalException {
        FormulaEvaluator evaluator = this.excelWB.getCreationHelper().createFormulaEvaluator();
        for (int intCount = 0; intCount < sheetsToRead.length; intCount++) {
            this.excelWS = this.excelWB.getSheetAt(sheetsToRead[intCount]);
            // Logger.getLogger(ExcelReader.class.getName()).log(Level.INFO, "Sheet names {0}", this.excelWS.getSheetName());
            if (sheetsToRead[intCount] == 3) {
                int dRows = this.excelWS.getPhysicalNumberOfRows();

//        System.out.println("Sheet " + declIndex + " \"" + wb.getSheetName(declIndex) + "\" has " + dRows
//                + " row(s).");
                for (int r = 0; r < dRows; r++) {
                    Row row = this.excelWS.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    if ((r == 8) || (r == 9) || (r == 22) || (r == 26) || (r == 27) || (r == 28) || (r == 29)) {
                        //column no. 3 corresponds to D
                        int c = 3;
                        if (c < row.getLastCellNum()) {
                            Cell cell = row.getCell(c);
                            String value;
                            CellValue cellValue = evaluator.evaluate(cell);
                            if (cellValue != null) {
                                switch (cellValue.getCellTypeEnum()) {
                                    case NUMERIC:
                                        value = String.valueOf(cellValue.getNumberValue());
                                        break;
                                    case STRING:
                                        value = cellValue.getStringValue();
                                        break;
                                    case BLANK:
                                        value = "";
                                        break;
                                    case BOOLEAN:
                                        value = String.valueOf(cellValue.getBooleanValue());
                                        break;
                                    case ERROR:
                                        value = String.valueOf(cellValue.getErrorValue());
                                        break;
                                    default:
                                        value = String.valueOf(cell.getCellTypeEnum());
                                }
                                System.out.print(value + "\t");
                            }
                        }
                        System.out.println("");
                    }
                }
            } else {
                int rows = this.excelWS.getPhysicalNumberOfRows();

//        System.out.println("Sheet " + smelterIndex + " \"" + wb.getSheetName(smelterIndex) + "\" has " + rows
//                + " row(s).");
                //start from row 5
                for (int r = 4; r < rows; r++) {

                    Row row = this.excelWS.getRow(r);

                    if (row == null) {
                        continue;
                    }
                    try {
                        for (int c = 0; c < row.getLastCellNum(); c++) {
                            Cell cell = row.getCell(c);
                            String value;
                            CellValue cellValue = evaluator.evaluate(cell);
                            if (cellValue != null) {
                                switch (cellValue.getCellTypeEnum()) {
                                    case NUMERIC:
                                        value = String.valueOf(cellValue.getNumberValue());
                                        break;
                                    case STRING:
                                        value = cellValue.getStringValue();
                                        break;
                                    case BLANK:
                                        value = "";
                                        break;
                                    case BOOLEAN:
                                        value = String.valueOf(cellValue.getBooleanValue());
                                        break;
                                    case ERROR:
                                        value = String.valueOf(cellValue.getErrorValue());
                                        break;
                                    default:
                                        value = String.valueOf(cell.getCellTypeEnum());
                                }
                                System.out.print(value + "\t");
                            }
                        }
                    } catch (Exception ex) {
                        throw new XLSFatalException("Can't access external reference");
                    }
                    System.out.println("");
                }
            }
        }

    }

    public FileInputStream getFileInputStream(String fileName) throws XLSFatalException {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(fileName);
        } catch (FileNotFoundException e) {
            throw new XLSFatalException(e);
        }
        return fis;
    }

    public abstract Workbook getWorkbookObj(InputStream fis) throws XLSFatalException;
}
