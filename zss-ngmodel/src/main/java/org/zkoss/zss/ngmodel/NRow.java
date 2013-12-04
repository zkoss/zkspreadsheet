package org.zkoss.zss.ngmodel;

import org.zkoss.zss.ngmodel.impl.SheetAdv;

public interface NRow {

	public NSheet getSheet();
	
	public int getIndex();
	public String asString();
	public boolean isNull();
	public NCellStyle getCellStyle();
	public NCellStyle getCellStyle(boolean local);
	
	public int getStartCellIndex();
	public int getEndCellIndex();
	
	//editable
	public void setCellStyle(NCellStyle cellStyle);
	
	public int getHeight();
	public boolean isHidden();
	
	public void setHeight(int height);
	public void setHidden(boolean hidden);
}
