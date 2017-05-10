package org.ak.spreadsheet.converter;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.ak.spreadsheet.exceptions.InvalidSpreadSheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class Converter {

	private static final String NEW_LINE = "\r\n";
	public static final String SPACE = " ";
	
	private ConverterConfig config = null;
	
	public Converter(ConverterConfig config) {
		this.config = config;
	}
	
	public Converter() {
		this.config = new ConverterConfig();
	}
	
	public String[] toCSV(String filename) throws InvalidSpreadSheetException {
		
		try {
			Workbook wb = WorkbookFactory.create(new File(filename));
			String[] csvResults = convertWorkbookToCSV(wb);
			return csvResults;
		} catch (Exception e) {
			throw new InvalidSpreadSheetException(e.getMessage());
		} 

	}

	public String[] toCSV(InputStream is) throws InvalidSpreadSheetException {
		try {
			Workbook wb = WorkbookFactory.create(is);
			String[] csvResults = convertWorkbookToCSV(wb);
			return csvResults;
		} catch (Exception e) {
			throw new InvalidSpreadSheetException(e.getMessage());
		} 
	}
	
	public String toJSON(String filename) throws InvalidSpreadSheetException {
		try {
			Workbook wb = WorkbookFactory.create(new File(filename));
			String json = convertWorkbookToJSON(wb);
			return json;
		} catch (Exception e) {
			throw new InvalidSpreadSheetException(e.getMessage());
		} 
	}
	
	public String toJSON(InputStream is) throws InvalidSpreadSheetException {
		try {
			Workbook wb = WorkbookFactory.create(is);
			String json = convertWorkbookToJSON(wb);
			return json;
		} catch (Exception e) {
			throw new InvalidSpreadSheetException(e.getMessage());
		} 
	}
	
	private String toXML(String filename) {
		
		return "";
	}
	
	private String toXML(InputStream is) {
		
		return "";
	}
	
	private String[] convertWorkbookToCSV(Workbook wb) {
		int sheetCount = wb.getNumberOfSheets();
		String[] csvResults = new String[sheetCount];
		FormulaEvaluator evaluator = null;
		if(config.isExecuteFormula()) {
			evaluator = wb.getCreationHelper().createFormulaEvaluator();
		}
		
		for(int i = 0; i < sheetCount; i++) {
			StringBuffer csvContent  = new StringBuffer();
			Sheet sheet = wb.getSheetAt(i);
			
			Row row = null;
		    Cell cell = null;
		    
		    Iterator<Row> rowIterator = sheet.iterator();
		    String delim = config.getDelimiter();
		    
		    while (rowIterator.hasNext()) {
		        row = rowIterator.next();
		        Iterator<Cell> cellIterator = row.cellIterator();
		        while (cellIterator.hasNext()) {
		            cell = cellIterator.next();
		            Object cellData = ConversionHelper.getCellValue(config, cell, evaluator);
		            if(cellData != null) {
		            	csvContent.append(cellData).append(delim);
		            } else{
		            	csvContent.append(SPACE).append(delim);
		            }
		            
		        }
		        
		        csvContent.append(NEW_LINE);
		    }
		    
		    csvResults[i] = csvContent.toString();
		}
		return csvResults;
	}
	
	private String convertWorkbookToJSON(Workbook wb) throws InvalidSpreadSheetException {
		int sheetCount = wb.getNumberOfSheets();
		FormulaEvaluator evaluator = null;
		if(config.isExecuteFormula()) {
			evaluator = wb.getCreationHelper().createFormulaEvaluator();
		}
		
		JsonObject wbObject = new JsonObject();
		JsonArray sheetsArray = new JsonArray();
		for(int i = 0; i < sheetCount; i++) {
			
			// populate sheets
			Sheet sheet = wb.getSheetAt(i);
			Row row = null;
		    Cell cell = null;
		    Iterator<Row> rowIterator = sheet.iterator();
		    
		    JsonObject sheetObject = new JsonObject();
		    ArrayList<String> columns = new ArrayList<String>();
		    int rowIndex=0;
		    
		    JsonArray rowsArray = new JsonArray();
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                if(rowIndex == 0) {
                	// retrieve column names
                	columns = ConversionHelper.getColumnHeadings(config, evaluator, row);
                } else if(rowIndex > 0 && columns.size() > 0){
                	JsonObject rowObject = new JsonObject();
                	Iterator<Cell> cellIterator = row.cellIterator();
                	int cellIndex = 0;
    		        while (cellIterator.hasNext()) {
    		            cell = cellIterator.next();
    		            Object cellData = ConversionHelper.getCellValue(config, cell, evaluator);
    		            if(cellData instanceof Number) { 
    		            	rowObject.addProperty(columns.get(cellIndex), (Number)cellData);
    		            } else if(cellData instanceof Boolean) {
    		            	rowObject.addProperty(columns.get(cellIndex), (Boolean)cellData);
    		            } else{
    		            	rowObject.addProperty(columns.get(cellIndex), (String)cellData);
    		            }
    		            
    		            cellIndex ++;
    		        }
                	
                	
                	rowsArray.add(rowObject);
                }
                
                rowIndex ++;
            }
		    sheetObject.add("rows", rowsArray);
			sheetsArray.add(sheetObject);
		}
		wbObject.add("sheets", sheetsArray);
		
		return wbObject.getAsString();
	}

}
