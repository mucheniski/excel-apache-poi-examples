package com.example.demo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.domains.Employee;

public class ExcelWriter {

	private static String[] columns = { "Name", "Email", "Date Of Birth", "Salary" };
	private static List<Employee> employees = new ArrayList<>();

	// Initializing employees data to insert into the excel file
	static {
		Calendar dateOfBirth = Calendar.getInstance();

		dateOfBirth.set(1992, 7, 21);
		employees.add(new Employee("Diego Mucheniski", "diego@example.com", dateOfBirth.getTime(), 1200000.0));

		dateOfBirth.set(1965, 10, 15);
		employees.add(new Employee("Bruna Ferreira Duarte", "bruna@example.com", dateOfBirth.getTime(), 1500000.0));
	}

	public static void main(String[] args) throws IOException, InvalidFormatException {

		// Create a Workbook
		Workbook workbook = new XSSFWorkbook(); // to generate .xlsx

		/*
		 * Creating a Helper helps us create instances of various things like
		 * DataFormat, Hyperlink, RichTextString, etc... In a format (HSSF, XSSF)
		 * independent way
		 */
		CreationHelper creationHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Employee");

		// Create a font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the fonts
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Crate a Row
		Row headerRow = sheet.createRow(0);

		// Create Cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create cellStyle for formating date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other Rows and Cells with Employee Data
		int rowNum = 1;
		for (Employee employee : employees) {

			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(employee.getName());

			row.createCell(1).setCellValue(employee.getEmail());

			Cell dateOfBirthCell = row.createCell(2);
			dateOfBirthCell.setCellValue(employee.getDateOfBirth());
			dateOfBirthCell.setCellStyle(dateCellStyle);

			row.createCell(3).setCellValue(employee.getSalary());
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();

	}

}
