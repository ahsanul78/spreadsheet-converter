package org.ak.spreadsheet.converter;

public class ConverterConfig {
	public static final String DEFAULT_DELIMITER = ",";
	
	private String delimiter;
	private boolean executeFormula;
	
	public ConverterConfig() {
		this.delimiter = DEFAULT_DELIMITER;
		this.executeFormula = false;
	}

	public String getDelimiter() {
		return delimiter;
	}

	public void setDelimiter(String delimiter) {
		this.delimiter = delimiter;
	}

	public boolean isExecuteFormula() {
		return executeFormula;
	}

	public void setExecuteFormula(boolean executeFormula) {
		this.executeFormula = executeFormula;
	}

}
