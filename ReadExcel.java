package com.mavencheck;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream("Employees.xlsx");
             Workbook workbook = new XSSFWorkbook(file)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            break;
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}