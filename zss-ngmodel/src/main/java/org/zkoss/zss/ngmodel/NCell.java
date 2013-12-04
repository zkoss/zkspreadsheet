package org.zkoss.zss.ngmodel;

import java.util.Date;

import org.zkoss.zss.ngmodel.impl.AbstractSheet;

public interface NCell {

	public enum CellType {
		BLANK,
		STRING,
		FORMULA,
		NUMBER,
		DATE,
		BOOLEAN,
		ERROR
	}
	
	public NSheet getSheet();
	
	public CellType getType();
	public CellType getFormulaResultType();
	
	public boolean isNull();
	
	public int getRowIndex();
	
	public int getColumnIndex();
	
	public String getReferenceString();
	
	public NCellStyle getCellStyle();
	
	public NCellStyle getCellStyle(boolean local);

//	boolean isReadonly();
//	
	
	public Object getValue();
	public void setValue(Object value);

	//clear cell value , reset it to blank
	public void clearValue();//
	public void clearFormulaResultCache();
	
	public void setStringValue(String value);
	public String getStringValue();
	
	/**
	 * set formula as string with '=', ex: SUM(A1:B2)
	 * @param fromula
	 */
	public void setFormulaValue(String formula);
	public String getFormulaValue();
	
	public void setNumberValue(Number number);
	public Number getNumberValue();
	
	public void setDateValue(Date date);
	public Date getDateValue();
	
	public ErrorValue getErrorValue();
	public void setErrorValue(ErrorValue errorValue);
	
	public void setCellStyle(NCellStyle cellStyle);
}
