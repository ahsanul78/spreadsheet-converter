# spreadsheet-converter

This is a utility library for convering excel spreadsheets to CSV, JSON and XML. Created usinf Apache POI library. 

## Code Example

```javascript
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

```
