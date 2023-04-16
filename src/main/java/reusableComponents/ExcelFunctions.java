package reusableComponents;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFunctions {
	
	static Row row;
	public static Row headerRow;
	static XSSFWorkbook workbook = null;
	static XSSFSheet sheet = null;
	static BufferedReader csvReader;
	static File csvPrimaryFile;
	
	public void openExcel(String filePath) {
		try {
			File f = new File(filePath);
			if(f.exists()) {
				FileInputStream file = new FileInputStream(f);
				workbook = new XSSFWorkbook(file);
				sheet = workbook.getSheetAt(0);
				headerRow = sheet.getRow(0);
			}else {
				workbook = new XSSFWorkbook();
				sheet = workbook.createSheet();
				headerRow = sheet.createRow(0);
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Write a single value in Excel
	 * 
	 * @param value
	 * 
	 */
	
	public static void addInExcel(String value) {
		int rowNum;
		try {
			rowNum = sheet.getLastRowNum();
			row = sheet.createRow(rowNum+1);
			row.createCell(0).setCellValue(value);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	
	/**
	 * Write Excel
	 * 
	 * @param contractNumber - contract number value
	 * @param status - of the field
	 * 
	 */
	
	public static void addInExcel(String contractNumber, String status) {
		int rowNum;
		try {
			rowNum = sheet.getLastRowNum();
			row = sheet.createRow(rowNum + 1);
			row.createCell(0).setCellValue(contractNumber);
			row.createCell(1).setCellValue(status);
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Write series of value in Excel
	 * 
	 * @param fieldVal the field value
	 * 
	 */
	
	public static void addInExcel(Map<String,String> fieldVal) {
		int fieldIndex;
		int rowNum;
		try {
			rowNum = sheet.getLastRowNum();
			row = sheet.createRow(rowNum+1);
			String val = "";
			Cell cell;
			Cell headerCell;
			for(String key: fieldVal.keySet()) {
				val = fieldVal.get(key);
				fieldIndex = getColumnIndex(sheet,key);
				if(fieldIndex > 0) {
					cell = row.createCell(fieldIndex);
					cell.setCellValue(val);
				} else {
					int size = (int) StreamSupport.stream(headerRow.spliterator(), false).count();
					headerCell = headerRow.createCell(size);
					headerCell.setCellValue(key);
					cell = row.createCell(size);
					cell.setCellValue(val);
				}
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Flush the output stream to excel
	 * 
	 * @param filePath - 
	 * 
	 */

	public static void writeInExcel(String filePath) {
		try(FileOutputStream outputStream = new FileOutputStream(filePath)){
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	/**
	 * Gets the index of the column in the excel
	 * 
	 * @param sheet - sheet for which column index is to be found
	 * @param columnName - Name of the column for which index is to be found
	 * @return columnIndex - Index of column
	 * 
	 */
	
	private static int getColumnIndex(XSSFSheet sheet, String columnName) {
		Row headerRow = sheet.getRow(0);
		Iterator<Cell> cellIterator = headerRow.cellIterator();
		while(cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if(cell.getStringCellValue().trim().equalsIgnoreCase(columnName.trim())) {
				return cell.getColumnIndex();
			}
		}
		return -1;
	}
	
	
	/**
	 * Gets the index of the row in the excel
	 * 
	 * @param sheet - sheet for which row index is to be found
	 * @param columnName - Name of the row for which index is to be found
	 * 
	 */
	
	private static int getRowIndex(XSSFSheet sheet, String value) {
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.cellIterator();
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				if(cell.getStringCellValue().trim().equalsIgnoreCase(value.trim()) && cell.getColumnIndex() == 1) {
					return currentRow.getRowNum();
				}
			}
		}
		return -1;
	}
	
	
	/**
	 * Gets the index of the row in the excel
	 * 
	 * @param sheet - sheet for which row index is to be found
	 * @param value - cell value
	 * @return rowNum
	 * 
	 */
	
	private static int getRowIndex(XSSFSheet sheet, String value, int columnIndex) {
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.cellIterator();
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				if(cell.getStringCellValue().trim().equalsIgnoreCase(value.trim()) && cell.getColumnIndex() == columnIndex) {
					return currentRow.getRowNum();
				}
			}
		}
		return -1;
	}
	
	/**
	 * Gets the column Indexes of all the columns
	 * 
	 * @param sheet - Sheet for which row index is to be found
	 * @return columnIndexes - Indexes of all the columns
	 * 
	 */
	
	public Map<String, Integer> getColumnIndexes(XSSFSheet sheet) {
		Map<String,Integer> headerIndexes = new HashMap<String, Integer>();
		Row headerRow= sheet.getRow(0);
		Iterator<Cell> cellIterator = headerRow.cellIterator();
		while(cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			headerIndexes.put(cell.getStringCellValue(), cell.getColumnIndex());
		}
		return headerIndexes;
	}

	/**
	 * Moves files
	 * 
	 * @param srcFileName - Name of the source file
	 * @param destinationFileName - Name of the destination file
	 * 
	 */

	public static boolean copyFile(String srcFileName, String destinationFileName) {
		try {
			File srcFile = new File(srcFileName);
			File destinationFile = new File(destinationFileName);
			destinationFile.mkdirs();
			FileUtils.copyFile(srcFile, destinationFile);
			return Files.exists(Paths.get(destinationFileName));
		}catch(Exception e) {
			e.printStackTrace();
		}
		return false;
	}


	/**
	 * Read mapping values from CSV file
	 * 
	 * @param e - FieldMap
	 * 
 	*/

	public static List<String> readFieldMappingValuesList(FieldMap e) throws IOException{
		File csvPrimaryFile = new File(CommonBean.execution.getProperty("fieldMappingPath") + 
			"FieldMappingSheet - " + CommonBean.productType.toUpperCase() + ".csv");
		BufferedReader csvReader = new BufferedReader(new FileReader(csvPrimaryFile.getAbsolutePath()));
	
		List<String> fieldMapping = new ArrayList<>();
		csvReader.lines().findFirst().get().split(",");
	
		for(String line : csvReader.lines().collect(Collectors.toList())) {
			String rows[] = line.split(",");
			if(fieldMapping.contains(rows[e.ordinal()])) {
				fieldMapping.add(rows[e.ordinal()]);
			}
		}
		csvReader.close();
		return fieldMapping;
	}

	
	/**
	 * Read mapping values from CSV file
	 * 
	 * @param e - FieldMap
	 * 
	 */

	public static LinkedHashMap<String,String> readFieldMappingValuesList(FieldMap e) throws IOException{
		if(csvPrimaryFile == null || !csvPrimaryFile.getName().contains(CommonBean.productType)) {
			File csvPrimaryFile = new File(CommonBean.execution.getProperty("fieldMappingPath") + 
					"FieldMappingSheet - " + CommonBean.productType.toUpperCase() + ".csv");
			csvReader = new BufferedReader(new FileReader(csvPrimaryFile.getAbsolutePath()));
		}
		LinkedHashMap<String,String> fieldMapping = new LinkedHashMap<>();
		csvReader.lines().findFirst().get().split(",");
	
		for(String line : csvReader.lines().collect(Collectors.toList())) {
			String rows[] = line.split(",");
			try {
				fieldMapping.put(rows[FieldMap.PDF.ordinal()], rows[e.ordinal()]);
			}catch(ArrayIndexOutOfBoundsException ae) {
				fieldMapping.put(rows[FieldMap.PDF.ordinal()], "");
			}
			csvReader.close();
			return fieldMapping;
		}
		
	public static void writeXMLValues(List<Map<String,String>> resultList) throws IOException {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("ddMMyyyy");
		for(Map<String,String> map : resultList) {
			String fileName = CommonBean.execution.getProperty("excelOutputFile") + "\\" +
								map.get("CONTRACT_NUMBER").replaceAll("[A\\-]","") + "-" +
								LocalDateTime.now().format(dtf) + ".xlsx";
			ExcelFunctions.openExcel(fileName);
			CellStyle style = workbook.createCellStyle();
			style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
			Font font = workbook.createFont();
			font.setFontHeightInPoints((short)15);
			font.setBold(true);
			style.setFont(font);
			
			headerRow.createCell(0).setCellValue("Contract#");
			headerRow.getCell(0).setCellStyle(style);
			headerRow.createCell(1).setCellValue("Field Names");
			headerRow.getCell(1).setCellStyle(style);
			headerRow.createCell(2).setCellValue("PDF Values");
			headerRow.getCell(2).setCellStyle(style);
			headerRow.createCell(3).setCellValue("XML Values");
			headerRow.getCell(3).setCellStyle(style);
			headerRow.createCell(4).setCellValue("Vantage Values");
			headerRow.getCell(4).setCellStyle(style);
			headerRow.createCell(5).setCellValue("P/F");
			headerRow.getCell(5).setCellStyle(style);
			
			System.out.println("The sheet name is :" + sheet.getSheetName());
			System.out.println("Generating the compare report");
			
			int i = 1;
			
			Map<String, String> fieldMapping = ExcelFunctions.readFieldMappingValuesList(FieldMap.XML);
			
			for(Map.Entry<String, String> k : fieldMapping.entrySet()) {
				Row row = sheet.createRow(i);
				row.createCell(1).setCellValue(k.getKey());
				for(String value : k.getValue().split(";")) {
					try {
						value = row.getCell(3).getStringCellValue() + "\n" + map.get(value);
						row.getCell(3).setCellValue(value);
					}catch(NullPointerException e) {
						if(value.isEmpty()) {
							continue;
						} else {
							row.createCell(3);
							value = map.get(value);
							row.getCell(3).setCellValue(value);
						}
					}
				}
				i++;
			}
			sheet.getRow(1).createCell(0).setCellValue(map.get("CONTRACT_NUM"));
			ExcelFunctions.writeInExcel(fileName);
		}
		
		
		public static void writeVantageValues(Map<String, String> map) throws IOException {
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("ddMMyyyy");
			String fileName = CommonBean.execution.getProperty("excelOutputFile") + "\\" +
								CommonBean.contractNumber + "-" + LocalDateTime.now().format(dtf) +
									".xlsx";
			ExcelFunctions.openExcel(fileName);
			
			System.out.println("The sheet name is :" +sheet.getSheetName());
			System.out.println("Generating the compare report");
			
			Map<String, String> fieldMapping = ExcelFunctions.readFieldMappingValues(FieldMap.VTG);
			
			for(Map.Entry<String, String> k : fieldMapping.entrySet()) {
				int rowIndex = getRowIndex(sheet, k.getKey());
				row = sheet.getRow(rowIndex);
				for(String value : k.getValue().split(";")) {
					if(value.isEmpty()) {
						continue;
					} else {
						row.createCell(4);
						value = map.get(value);
						row.getCell(4).setCellValue(value);
					}
				}
			}
			
			ExcelFunctions.writeInExcel(fileName);
		}
		
		
		public static void writeVantageValuesWithXMLValues(Map<String, String> map) throws IOException{
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("ddMMyyyy");
			String fileName = CommonBean.execution.getProperty("excelOutputFile") + "\\" + 
								CommonBean.contractNumber + "-" + LocalDateTime.now().format(dtf)
									+ ".xlsx";
								
			ExcelFunctions.openExcel(fileName);
			CellStyle style = workbook.createCellStyle();
			style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
			Font font = workbook.createFont();
			font.setFontHeightInPoints((short)15);
			font.setBold(true);
			style.setFont(font);
			
			headerRow.createCell(0).setCellValue("Field Names");
			headerRow.getCell(0).setCellStyle(style);
			headerRow.createCell(1).setCellValue("PDF Values");
			headerRow.getCell(1).setCellStyle(style);
			headerRow.createCell(2).setCellValue("XML Values");
			headerRow.getCell(2).setCellStyle(style);
			headerRow.createCell(3).setCellValue("Vantage Values");
			headerRow.getCell(3).setCellStyle(style);
			headerRow.createCell(4).setCellValue("P/F");
			headerRow.getCell(4).setCellStyle(style);
			
			System.out.println("The sheet name is :" + sheet.getSheetName());
			System.out.println("Generating the compare report");
			
			int i = 1;
			
			Map<String, String> fieldMapping = ExcelFunctions.readFieldMappingValues(FieldMap.VTG);
			
			for(Map.Entry<String, String> k : fieldMapping.entrySet()) {
				Row row = sheet.createRow(i);
				row.createCell(0).setCellValue(k.getKey());
				String value;
				if(k.getValue().isEmpty()) {
					continue;
				} else {
					value = (String) CommonBean.json.get(k.getKey());
					row.createCell(2).setCellValue(value);
					value = map.get(k.getValue()) == null ? : map.get(k.getValue());
					row.createCell(3).setCellValue(value);
				}
				i++;
			}
			ExcelFunctions.writeInExcel(fileName);
		}
		
	
	
}
	 
