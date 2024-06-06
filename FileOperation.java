package com.mavencheck;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FileOperation {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        String[] headers = {"Name", "Age", "Email"};
        String[][] data = {
                {"John Doe", "30", "john@test.com"},
                {"Jane Doe", "28", "jane@test.com"},
                {"Bob Smith", "35", "bob@example.com"},
                {"Swapnil", "37", "swapnil@example.com"}
        };

        // Create header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Write data rows
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("Employees.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Excel file written successfully.");
    }
}