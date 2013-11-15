package org.zkoss.zss.ngmodel;

public interface NRow {

	int getIndex();
	String asString();
	boolean isNull();
//	NCellStyle getCellStyle();
	
	int getStartCellIndex();
	int getEndCellIndex();
}
