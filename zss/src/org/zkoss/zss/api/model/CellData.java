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
	
	public void setValue(Object value);
	
	public void setEditText(String editText);
	
	//public void setCellType(CellType type);
}
