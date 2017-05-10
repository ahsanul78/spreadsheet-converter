package org.ak.spreadsheet.converter;

import java.util.ArrayList;
import java.util.Iterator;

import org.ak.spreadsheet.exceptions.InvalidSpreadSheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class ConversionHelper {
	
	public static Object getCellValue(ConverterConfig config, Cell cell, FormulaEvaluator evaluator) {
		if(cell.getCellTypeEnum().equals(CellType.BOOLEAN)) {
        	return cell.getBooleanCellValue();
        } else if(cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
        	Double value= null;
            value = cell.getNumericCellValue();
            return value;
            
        } else if(cell.getCellTypeEnum().equals(CellType.STRING)) {
        	String text = null;
        	if(cell.getStringCellValue() != null){
        		text = cell.getStringCellValue();
        		if(text.contains(config.getDelimiter())){
        			text = "\""+text+"\"";
        		}
        	}
        	return text;
        	
        } else if(cell.getCellTypeEnum().equals(CellType.FORMULA)) {
        	// based on config take decision to execute
        	Object cellData = null;
        	if(config.isExecuteFormula() && evaluator != null) {
        		CellValue cellValue = evaluator.evaluate(cell);
        		if(cellValue.getCellTypeEnum().equals(CellType.BOOLEAN)) {
        			cellData = cellValue.getBooleanValue();
        		} else if(cellValue.getCellTypeEnum().equals(CellType.NUMERIC)) {
        			Double value= null;
                    value = cellValue.getNumberValue();
                    cellData = value;
        		} else if(cellValue.getCellTypeEnum().equals(CellType.STRING)) {
        			String text = cellValue.getStringValue();
            		if(text.contains(config.getDelimiter())){
            			text = "\""+text+"\"";
            		}
            		cellData = text;
        		} 
        		
        	} else{
        		cellData = cell.getCellFormula();
        	}
        	return cellData;
        } else {
        	return null;
        }
        	
	}
	
	public static ArrayList<String> getColumnHeadings(ConverterConfig config, FormulaEvaluator evaluator, Row row) 
			throws InvalidSpreadSheetException {
		if(row != null) {
			ArrayList<String> columns = new ArrayList<String>();
			Cell cell;
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
			    cell = cellIterator.next();
			    Object cellData = ConversionHelper.getCellValue(config, cell, evaluator);
			    if(cellData != null) {
			    	columns.add(cellData.toString());
			    } else{
			    	throw new InvalidSpreadSheetException("Column heading is not present");
			    }
			}
			return columns;
		} else{
			throw new InvalidSpreadSheetException("Column heading is not present");
		}
	}
}
