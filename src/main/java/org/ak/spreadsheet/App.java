package org.ak.spreadsheet;

import org.ak.spreadsheet.converter.Converter;
import org.ak.spreadsheet.converter.ConverterConfig;
import org.ak.spreadsheet.exceptions.InvalidSpreadSheetException;

public class App 
{
    public static void main( String[] args )
    {
        ConverterConfig config = new ConverterConfig();
        Converter converter = new Converter(config);
        
        try {
			String xml = converter.toXML("C://programs.xlsx");
			System.out.println(xml);
		} catch (InvalidSpreadSheetException e) {
			e.printStackTrace();
		}
    }
}
