package com.excelreadwrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderWriter {
	FileInputStream inputStream = null;
	Workbook workbook = null;
	Row row = null;
	Cell cell = null;
	Sheet sheet = null;
	String excelFilePath = "./EmployeeData.xlsx";

	public ExcelReaderWriter() {
		try {
			// Give excel file path
			inputStream = new FileInputStream(excelFilePath);
			workbook = new XSSFWorkbook(inputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public Map<String, String> readExcelData(String sheetName, String empID) {
		Map<String, String> testData = new HashMap<String, String>();
		sheet = workbook.getSheet(sheetName);
		Iterator<Row> rowIterator = sheet.iterator();
		row = sheet.getRow(0);
		int colNum = row.getLastCellNum();
		Map<String, Integer> colMapByName = new LinkedHashMap<String, Integer>();
		if (rowIterator.hasNext()) {
			Row nextRow = rowIterator.next();
			for (int j = 0; j < colNum; j++) {
				colMapByName.put(cellToString(nextRow.getCell(j)), j);
			}
		}

		while (rowIterator.hasNext()) {
			Row nextRow = rowIterator.next();
			if (nextRow.getCell(0).getStringCellValue().equalsIgnoreCase(empID)) {
				for (Entry<String, Integer> colData : colMapByName.entrySet()) {
					cell = nextRow.getCell(colData.getValue());
					testData.put(colData.getKey(), cellToString(cell));
				}
				break;
			}
		}
		try {
			inputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return testData;
	}

	/**
	 * Format cell value to a string
	 */
	private String cellToString(Cell cell) {
		final DataFormatter dataFormatter = new DataFormatter();
		return dataFormatter.formatCellValue(cell);
	}

	public void writeToExcel(String sheetName, String empID, String columnName, String data) {
		sheet = workbook.getSheet(sheetName);
		Iterator<Row> iterator = sheet.iterator();
		row = sheet.getRow(0);
		int columnNumber = 0;
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			cell = cellIterator.next();
			String cellValue = cell.getStringCellValue();
			if (columnName.equals(cellValue)) {
				columnNumber = cell.getColumnIndex();
				break;
			}
		}
		System.out.println(columnNumber);
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			if (nextRow.getCell(0).getStringCellValue().equalsIgnoreCase(empID)) {
				nextRow.getCell(columnNumber).setCellValue(data);
				break;
			}
		}
		try {
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(new File(excelFilePath));
			workbook.write(outputStream);
			outputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
