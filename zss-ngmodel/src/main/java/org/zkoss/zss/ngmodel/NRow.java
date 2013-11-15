package org.zkoss.zss.ngmodel;

public interface NRow {

	int getIndex();
	
	boolean isNull();
//	NCellStyle getCellStyle();
	
	int getStartCellIndex();
	int getEndCellIndex();
}
