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
 * A row of a sheet.
 * @author dennis
 * @since 3.5.0
 */
public interface SRow {

	public SSheet getSheet();
	
	public int getIndex();
//	public String asString();
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
	
	public int getHeight();
	public boolean isHidden();
	public boolean isCustomHeight();
	
	public void setHeight(int height);
	public void setHidden(boolean hidden);
	public void setCustomHeight(boolean custom);
}
