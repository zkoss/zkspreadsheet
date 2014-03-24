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
public interface SRow extends SCellStyleHolder{

	public SSheet getSheet();
	
	/**
	 * @return 0-based row index
	 */
	public int getIndex();
//	public String asString();
	public boolean isNull();
	
	/**
	 * @return height in pixels
	 */
	public int getHeight();
	public boolean isHidden();
	public boolean isCustomHeight();
	
	public void setHeight(int height);
	public void setHidden(boolean hidden);
	
	/**
	 * @param custom TRUE, if the height is set by users and is not calculated by system automatically
	 */
	public void setCustomHeight(boolean custom);
}
