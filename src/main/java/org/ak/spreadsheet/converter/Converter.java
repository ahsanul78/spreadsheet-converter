package org.ak.spreadsheet.converter;

import java.io.File;
import java.io.InputStream;
import java.util.Iterator;

import org.ak.spreadsheet.exceptions.InvalidSpreadSheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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
	
	public String toJSON(InputStream ios) {
		
		return "";
	}
	
	private String toXML(InputStream ios) {
		
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
}
