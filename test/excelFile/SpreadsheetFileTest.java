package excelFile;

import excelFile.SpreadsheetFile;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.*;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Test;
import org.junit.Assert;

public class SpreadsheetFileTest {

    @BeforeClass
    void setUpSpreadSheetFile() {
        SpreadsheetFile spreadsheet = new SpreadsheetFile(2, 2);
        spreadsheet.createExcel();
    }

    @Before
    void setUpWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        Row row = sheet.createRow(1);
    }

    @Test
    void testMakeCell() {

    }


}
