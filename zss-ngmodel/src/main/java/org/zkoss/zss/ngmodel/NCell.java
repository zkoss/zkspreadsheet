package org.zkoss.zss.ngmodel;

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
	
	boolean isNull();
	
	int getRowIndex();
	
	int getColumnIndex();
	
	String asString(boolean enableSheetName);
	
//	NCellStyle getCellStyle();
//	
//	boolean isReadonly();
//	
	
	//editable
	void setValue(Object value);
	Object getValue(); 
}
