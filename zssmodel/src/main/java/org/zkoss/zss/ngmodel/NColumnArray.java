package org.zkoss.zss.ngmodel;

public interface NColumnArray {

	public int getIndex();
	public int getLastIndex();
	
	public NSheet getSheet();
	
	
	public NCellStyle getCellStyle();
	
	//editable
	public void setCellStyle(NCellStyle cellStyle);
	
	public int getWidth();
	public boolean isHidden();
	public boolean isCustomWidth();
	
	public void setWidth(int width);
	public void setHidden(boolean hidden);
	public void setCustomWidth(boolean custom);
}
