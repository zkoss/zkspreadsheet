package org.zkoss.zss.ngmodel;

public interface NRow {

	int getIndex();
	String asString();
	boolean isNull();
	NCellStyle getCellStyle();
	NCellStyle getCellStyle(boolean local);
	
	int getStartCellIndex();
	int getEndCellIndex();
	
	//editable
	void setCellStyle(NCellStyle cellStyle);
}
