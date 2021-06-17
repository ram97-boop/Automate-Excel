package excelFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadsheetFile {

    private final int rows;
    private final int columns;

    public SpreadsheetFile(int rows, int columns) {
        this.rows = rows+2;
        this.columns = columns+2;
    }

    public XSSFWorkbook createExcel() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        for (int rowNum=1; rowNum<rows; rowNum++) {
            Row row = sheet.createRow(rowNum);

            if (rowNum==1) {
                for (int colNum=1; colNum<=columns; colNum++) {
                    if (colNum==1) {
                        makeCell(workbook, row, colNum, "grey", "Name");
                    }
                    else if (colNum==columns) {
                        makeCell(workbook, row, colNum, "grey", "Total");
                    }
                    else {
                        makeCell(workbook, row, colNum, "grey", "Input " + (colNum-1));
                    }
                }
            }
            else {
                String firstColumnAddress = "";
                String lastColumnAddress = "";
                for (int colNum=2; colNum<=columns; colNum++) {
                    Cell cell = row.createCell(colNum);

                    if (colNum==2) {
                        firstColumnAddress = cell.getAddress().formatAsString();
                    }
                    if (colNum+1 == columns) {
                        lastColumnAddress = cell.getAddress().formatAsString();
                    }
                    if (colNum==columns) {
                        row.getCell(colNum).setCellFormula("SUM(" + firstColumnAddress + ":" + lastColumnAddress + ")");
                    }
                }
            }
        }
        Row row = sheet.getRow(2);
        Cell cell1 = row.getCell(2);
        row = sheet.getRow(rows-1);
        Cell cell2 = row.getCell(columns-1);
        row = sheet.createRow(rows);
        makeCell(workbook, row, columns, "yellow");
        row.createCell(1).setCellValue("TOTAL");
        row.getCell(columns).setCellFormula("SUM(" + cell1.getAddress().formatAsString() + ":" + cell2.getAddress().formatAsString() + ")");
        return workbook;
    }


    private void makeCell(Workbook workbook, Row row, int colNum, String color, String label) {
        Cell cell = row.createCell(colNum);
        cell.setCellValue(label);
        cell.setCellStyle(backgroundStyle(workbook, color));
    }


    private void makeCell(Workbook workbook, Row row, int colNum, String color) {
        makeCell(workbook, row, colNum, color, "");
    }


    private CellStyle backgroundStyle(Workbook workbook, String color) {
        CellStyle bgStyle = workbook.createCellStyle();
        bgStyle.setFillPattern(FillPatternType.DIAMONDS);

        if (color.equals("pink")) {
            bgStyle.setFillBackgroundColor(IndexedColors.PINK1.getIndex());
        }
        else if (color.equals("yellow")) {
            bgStyle.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        }
        else if (color.equals("green")) {
            bgStyle.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        }
        else if (color.equals("blue")) {
            bgStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        }
        else if (color.equals("grey")) {
            bgStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        }
        else {
            bgStyle.setFillPattern(FillPatternType.NO_FILL);
        }

        return bgStyle;
    }


    private CellStyle backgroundStyle(Workbook workbook) {
        return backgroundStyle(workbook, "");
    }
}
