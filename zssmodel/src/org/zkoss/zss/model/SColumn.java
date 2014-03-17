/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model;

/**
 * A column of a sheet.
 * @author dennis
 * @since 3.5.0
 */
public interface SColumn {

	public int getIndex();
	public SSheet getSheet();
//	public String asString();
	public boolean isNull();
	public SCellStyle getCellStyle();
//	public NCellStyle getCellStyle(boolean local);
	
	//editable
	public void setCellStyle(SCellStyle cellStyle);
	
	public int getWidth();
	public boolean isHidden();
	public boolean isCustomWidth();
	
	public void setWidth(int width);
	public void setHidden(boolean hidden);
	public void setCustomWidth(boolean custom);
	
}
