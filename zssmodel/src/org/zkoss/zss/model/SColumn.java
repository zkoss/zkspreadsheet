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
	public boolean isNull();
	
	/**
	 * Get the cell style, it always looks forward the sheet's style if local style is null.
	 * @see #getCellStyle(boolean)
	 */
	public SCellStyle getCellStyle();
	
	/**
	 * Get the cell style locally or look forward the sheet's style.
	 * @param local true to get the local style only, 
	 */
	public SCellStyle getCellStyle(boolean local);
	
	/**
	 * Set the cell style, give the cell-style to set a local one or null to clean local one
	 * @param cellStyle the style to set, null to clean local style
	 */
	public void setCellStyle(SCellStyle cellStyle);
	
	public int getWidth();
	public boolean isHidden();
	public boolean isCustomWidth();
	
	public void setWidth(int width);
	public void setHidden(boolean hidden);
	public void setCustomWidth(boolean custom);
	
}
