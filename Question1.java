package task13;
	
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;

public class Question1 {

    public static void main(String[] args) {
        // Create a new Excel workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a new sheet named "Sheet1"
        Sheet sheet = workbook.createSheet("Sheet1");

        // Define the data to write
        String[] headers = {"Name", "Age", "Email"};
        String[][] data = {
            {"John Doe", "30", "john@test.com"},
            {"Jane Doe", "28", "jane@test.com"},
            {"Bob Smith", "35", "bob@example.com"},
            {"Sivapnil", "37", "siva@example.com"}
        };

        // Create a header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Write data to the sheet
        for (int rowNum = 0; rowNum < data.length; rowNum++) {
            Row row = sheet.createRow(rowNum + 1);
            for (int colNum = 0; colNum < data[rowNum].length; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(data[rowNum][colNum]);
            }
        }

        // Write the workbook to an Excel file
        try (FileOutputStream outputStream = new FileOutputStream("example.xlsx")) {
            workbook.write(outputStream);
            System.out.println("Data has been written to Excel file.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Now, let's read and print the data from the Excel file
        try (FileInputStream inputStream = new FileInputStream("example.xlsx");
             Workbook readWorkbook = WorkbookFactory.create(inputStream)) {

            Sheet readSheet = readWorkbook.getSheet("Sheet1");

            for (Row row : readSheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
