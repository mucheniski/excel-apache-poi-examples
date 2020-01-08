package com.example.demo;

// https://www.callicoder.com/java-read-excel-file-apache-poi/

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {
	
	public static final String SAMPLE_XML_FILE_PATH = "./sample-xlsx-file.xlsx";
	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		
		// Creating a Workbook from a excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XML_FILE_PATH));
		
		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets" );
		System.out.println();
		
		// 1 You can obtain a sheetIterator and iterate over it
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		System.out.println("Retrieving Sheets using iterator");
		while (sheetIterator.hasNext()) {
			Sheet sheet = sheetIterator.next();
			System.out.println("Sheet Name: " + sheet.getSheetName());
		}
		System.out.println();
		
		// 2 You can use a for-each loop
		System.out.println("Retrieving  Sheets using for-each loop");
		for (Sheet sheet: workbook) {
			System.out.println("Sheet Name: " + sheet.getSheetName());
		}
		System.out.println();
		
		// 3 You can use a Java 8 forEach with lambda
		System.out.println("Retrieving Sheets using a Java 8 forEach with lambda");
		workbook.forEach(sheet -> {
			System.out.println("Sheet Name: " + sheet.getSheetName());
		});
		System.out.println();
		
		// Getting the sheet at index zero
		Sheet sheet = workbook.getSheetAt(0);
		System.out.println("SheetZero: " + sheet.getSheetName());
		
		// Create a DataFormatter to format and get each cell's value as String
		DataFormatter dataFormatter = new DataFormatter();
		
		// 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 3. Or you can use Java 8 forEach loop with lambda
        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });
		
		
		// Closing the workbook
		workbook.close();
		
	}

}
