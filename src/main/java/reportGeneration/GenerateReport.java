package reportGeneration;

import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.json.simple.parser.JSONParser;

public class GenerateReport {

	JSONObject valuesDetails = null;
	static LinkedHashMap<String,String> pojoMap;
	static LinkedHashMap<String,String> transformationFieldSequence;
	boolean fileCheck = true;
	StringBuffer fileName;
	XSSFWorkbook workbook;
	File file;
	//PropertyManager prop = new PropertyManager("execution.properties");
	
	public void generateReport(Map<String, PojoOutput> generatedMap, String contractNum)
						throws FileNotFoundException, IOException, ParseException {
		int rowCount = 0;
		readInputSheet();
		pojoMap = new LinkedHashMap<String,String>();
		
		if(fileCheck) {
			workbook = new XSSFWorkbook();
			
			Date date = new Date();
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
			fileName = new StringBuffer(prop.getProperty("ComparisonReport.path"));
			
			fileName.append("Execution report_");
			fileName.append(dateFormat.format(date));
			fileName.append("_.xlsx");
			file = new File(fileName.toString());
			fileCheck = false;
		}
		
		XSSFSheet sheet = workbook.createSheet(contractNum);
		
		sheet.setColumnWidth(0, 5 * 256);
		sheet.setColumnWidth(1, 50 * 256);
		sheet.setColumnWidth(2, 28 * 256);
		sheet.setColumnWidth(3, 28 * 256);
		sheet.setColumnWidth(4, 28 * 256);
		sheet.setColumnWidth(5, 28 * 256);

		Row rowHeading = sheet.createRow(0);
		
		rowHeading.createCell(0).setCellValue("Sr No");
		rowHeading.createCell(0).setCellValue("Field Name");
		rowHeading.createCell(0).setCellValue("Source Value");
		rowHeading.createCell(0).setCellValue("Transformed Value");
		rowHeading.createCell(0).setCellValue("Targeted Value");
		rowHeading.createCell(0).setCellValue("Result");
		
		/*
		 * CellStyle backgroundStyle = workbook.createCellStyle();
		 * style1.setFillBackgroundColor((short)75);
		 * backgroundStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		 * backgroundStyle.setAlignment(arg0);
		 */

		generatedMap.forEach(k, value) -> {
			PojoOutput obj = (PojoOutput) value;
			
			/*Row row = sheet.createRow(++rowCount);
			int columnCount = 0;
			
			PojoOutput obj = (PojoOutput) value;
			
			Cell cell = row.createCell(0);
			cell.setCellValue(rowCount);
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(obj.getFieldName());
			Cell cell2 = row.createCell(2);
			cell1.setCellValue(obj.getTransformedValue());
			Cell cell3 = row.createCell(3);
			cell1.setCellValue(obj.getTargetValue());
			Cell cell4 = row.createCell(4);
			cell1.setCellValue(obj.getResult());
			
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
			
			if(result.equalsIgnoreCase("Passed")) {
				font.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
				style.setFont(font);
			} else {
				font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
			}
			cell3.setCellStyle(style);*/
			
			pojoMap.put(obj.getFieldName(),
					obj.getTransformedValue() + "`~" + obj.getTargetValue() + "`~" + obj.getResult());
		});
		
		for(String key : TransformationFieldSequence.keySet()) {
			if(pojoMap.get(key)!=null) {
				String arr[] = pojoMap.get(key).split("`~");
				Row row = sheet.createRow(++rowCount);
				int columnCount = 0;
				
				PojoOutput obj = (PojoOutput) value;
				
				Cell cell = row.createCell(0);
				cell.setCellValue(rowCount);
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(key);
				Cell cell2 = row.createCell(3);
				cell2.setCellValue(arr[0]);
				Cell cell3 = row.createCell(5);
				cell3.setCellValue(arr[1]);
				Cell cell4 = row.createCell(0);
				String result = arr[2];
				cell4.setCellValue(result);
				
				CellStyle style = workbook.createCellStyle();
				Font font = workbook.createFont();
				
				if(result.equalsIgnoreCase("Passed")) {
					font.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
					style.setFont(font);
				} else {
					font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
					style.setFont(font);
				}
				cell4.setCellStyle(style);
			}
		}
		
		if(!file.exists()) {
			file.getParentFile().mkdir();
		}
		
		FileOutputStream output_file = new FileOutputStream(file);
		workbook.write(output_file);
		output_file.close();
		workbook.close();
		JSONParser parser = new JSONParser();
		String fileName1 = prop.getProperty("VantageJSON.path") + contractNum + "_VTG.json";
		valuesDetails = (JSONObject) parser.parse(new FileReader(fileName1));
		int  fieldCount = ExcelUtils.getRowCount(fileName.toString(), contractNum);
		
		for(int counter=1; counter < fieldCount; counter++) {
			System.out.println(counter);
			String fieldName = ExcelUtils.getCellData(fileName.toString(), contractNum, counter, 1);
			String valueFromJSON = (String) valuesDetails.get(fieldName);
			ExcelUtils.setCellData(fileName.toString(), contractNum, counter, 2, valueFromJSON);
		}
		System.out.println("Report generated successfully");
	}

	public void readInputSheet() throws FileNotFoundException, IOException, ParseException {
		try {
			TransformationFieldSequence = new LinkedHashMap<String, String>();
			FileInputStream ip = new FileInputStream(prop.getProperty("TransformationLayer.path"));
			Workbook wb = WorkbookFactory.create(ip);
			Sheet sheet = wb.getSheet("Transformation");
			
			int i,j;
			int rowCount = sheet.getPhysicalNumberOfRows(), cellCount = 1;
			Row row;
			Cell cell;
			for(i=1; i<rowCount; i++) {
				for(j=0; j<cellCount; j++)
				{
					row = sheet.getRow(i);
					cell = row.getCell(j);
					
					String fid = (row.getCell(0).toString().trim());
					String fname = (row.getCell(1).toString().trim());
					TransformationFieldSequence.put(fname, fid);
				}
			}
			wb.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
