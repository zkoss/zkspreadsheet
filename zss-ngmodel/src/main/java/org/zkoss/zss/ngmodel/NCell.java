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
	
	NCellStyle getCellStyle();
	
	NCellStyle getCellStyle(boolean local);

//	boolean isReadonly();
//	
	
	//editable
	void setValue(Object value);
	Object getValue();
	void setCellStyle(NCellStyle cellStyle);
}
