package excelFile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.util.Scanner;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateFile {

    public static void main(String[] args) {
        Scanner userInput = new Scanner(System.in);
        String fileName;
        int rowsInt = 0;
        int columnsInt = 0;

        while (true) {
            try {
                System.out.println("Enter name of file to create, e.g. 'employeeSalaries' (enter 'quit' to abort): ");
                fileName = userInput.nextLine();
                if (fileName.equals("quit")) {
                    return;
                }
                else {
                    fileName += ".xlsx";
                    break;
                }
            }
            catch (Exception e) {
                e.printStackTrace();
            }
        }

        while (true) {
            try {
                System.out.println("How many rows (enter 'quit' to abort)? ");
                String rows = userInput.nextLine();
                if (rows.equals("quit")) {
                    return;
                }
                else {
                    rowsInt = Integer.parseInt(rows);
                    break;
                }
            }
            catch (NumberFormatException e) {
                System.out.println("You must enter a number.");
            }
        }
        while (true) {
            try {
                System.out.println("How many columns (enter 'quit' to abort)? ");
                String columns = userInput.nextLine();
                if (columns.equals("quit")) {
                    return;
                }
                else {
                    columnsInt = Integer.parseInt(columns);
                    break;
                }
            }
            catch (NumberFormatException e) {
                System.out.println("You must enter a number.");
            }
        }

        try {
            System.out.println("Creating excel.");

            FileOutputStream outputStream = new FileOutputStream(fileName);
            SpreadsheetFile excel = new SpreadsheetFile(rowsInt, columnsInt);
            XSSFWorkbook workbook = excel.createExcel();
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println("Done.");
        }
        catch (FileNotFoundException e) {
            e.printStackTrace();

            try {
                Files.delete(Paths.get(fileName));
            }
            catch (IOException e2) {
                e2.printStackTrace();
            }
        }
        catch (IOException e) {
            e.printStackTrace();

            try {
                Files.delete(Paths.get(fileName));
            }
            catch (IOException e2) {
                e2.printStackTrace();
            }
        }
        catch (Exception e) {
            e.printStackTrace();
            try {
                Files.delete(Paths.get(fileName));
            }
            catch (IOException e2) {
                e2.printStackTrace();
            }
        }
        try {
            Files.delete(Paths.get("quit"));
        }
        catch (NoSuchFileException e) {
            return;
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
}