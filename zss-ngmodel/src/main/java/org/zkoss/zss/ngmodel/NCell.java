package org.zkoss.zss.ngmodel;

import java.util.Date;

public interface NCell {

	public enum CellType {
		BLANK,
		STRING,
		FORMULA,
		NUMBER,
		DATE,
		ERROR
	}
	
	CellType getType();
	CellType getFormulaResultType();
	
	boolean isNull();
	
	int getRowIndex();
	
	int getColumnIndex();
	
	String asString(boolean enableSheetName);
	
	NCellStyle getCellStyle();
	
	NCellStyle getCellStyle(boolean local);

//	boolean isReadonly();
//	
	
	Object getValue();
	void setValue(Object value);

	//clear cell value , reset it to blank
	void clearValue();//
	void clearFormulaResultCache();
	
	void setStringValue(String value);
	String getStringValue();
	
	/**
	 * set formula as string with '=', ex: SUM(A1:B2)
	 * @param fromula
	 */
	void setFormulaValue(String formula);
	String getFormulaValue();
	
	void setNumberValue(Number number);
	Number getNumberValue();
	
	void setDateValue(Date date);
	Date getDateValue();
	
	ErrorValue getErrorValue();
	void setErrorValue(ErrorValue errorValue);
	
	void setCellStyle(NCellStyle cellStyle);
}
