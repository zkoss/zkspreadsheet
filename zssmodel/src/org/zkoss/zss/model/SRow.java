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
	public SCellStyle getCellStyle();
	
	//editable
	public void setCellStyle(SCellStyle cellStyle);
	
	public int getHeight();
	public boolean isHidden();
	public boolean isCustomHeight();
	
	public void setHeight(int height);
	public void setHidden(boolean hidden);
	public void setCustomHeight(boolean custom);
}
