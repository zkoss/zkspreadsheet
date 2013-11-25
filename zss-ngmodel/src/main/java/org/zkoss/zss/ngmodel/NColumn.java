package org.zkoss.zss.ngmodel;

public interface NColumn {

	public int getIndex();
	public String asString();
	public boolean isNull();
	public NCellStyle getCellStyle();
	public NCellStyle getCellStyle(boolean local);
	
	//editable
	public void setCellStyle(NCellStyle cellStyle);
}
