package org.zkoss.zss.model;

public interface SColumnArray {

	public int getIndex();
	public int getLastIndex();
	
	public SSheet getSheet();
	
	
	public SCellStyle getCellStyle();
	
	//editable
	public void setCellStyle(SCellStyle cellStyle);
	
	public int getWidth();
	public boolean isHidden();
	public boolean isCustomWidth();
	
	public void setWidth(int width);
	public void setHidden(boolean hidden);
	public void setCustomWidth(boolean custom);
}
