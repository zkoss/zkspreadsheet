package org.zkoss.zss.ngmodel;

import org.zkoss.zss.ngmodel.impl.SheetAdv;

public interface NColumn {

	public int getIndex();
	public NSheet getSheet();
	public String asString();
	public boolean isNull();
	public NCellStyle getCellStyle();
	public NCellStyle getCellStyle(boolean local);
	
	//editable
	public void setCellStyle(NCellStyle cellStyle);
}
