package org.ak.spreadsheet;

import org.ak.spreadsheet.converter.Converter;
import org.ak.spreadsheet.converter.ConverterConfig;

public class App 
{
    public static void main( String[] args )
    {
        ConverterConfig config = new ConverterConfig();
        Converter converter = new Converter(config);
        
        try {
        	String[] csvs = converter.toCSV("C://programs.xlsx");
			System.out.println(csvs[0]);
			
			String xml = converter.toXML("C://programs.xlsx");
			System.out.println(xml);
			
			String json = converter.toJSON("C://programs.xlsx");
			System.out.println(json);
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}
