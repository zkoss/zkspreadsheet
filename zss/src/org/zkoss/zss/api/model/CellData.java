package org.zkoss.zss.api.model;

public interface CellData {
	
	public enum CellType{
		NUMERIC,
		STRING,
		FORMULA,
		BLANK,
		BOOLEAN,
		ERROR;
	}
	
	public int getRow();
	public int getColumn();
	
	public CellType getType();
	
	public CellType getResultType();
	
	public Object getValue();
	public String getFormatText();
	public String getEditText();
	
	public boolean isBlank();
	public boolean isFormula();
	
	public void setValue(Object value);
	
	public void setEditText(String editText);
	
	public boolean validateEditText(String editText);
}
