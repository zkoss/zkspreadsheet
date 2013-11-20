package org.zkoss.zss.ngmodel;

public interface NColumn {

	int getIndex();
	String asString();
	boolean isNull();
	NCellStyle getCellStyle();
	NCellStyle getCellStyle(boolean local);
	
	//editable
	void setCellStyle(NCellStyle cellStyle);
}
