package org.zkoss.zss.ngmodel;

public interface NRow {

	public int getIndex();
	public String asString();
	public boolean isNull();
	public NCellStyle getCellStyle();
	public NCellStyle getCellStyle(boolean local);
	
	public int getStartCellIndex();
	public int getEndCellIndex();
	
	//editable
	public void setCellStyle(NCellStyle cellStyle);
}
