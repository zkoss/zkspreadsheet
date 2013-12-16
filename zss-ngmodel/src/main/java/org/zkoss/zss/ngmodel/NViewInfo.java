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
package org.zkoss.zss.ngmodel;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NViewInfo {
	public int getNumOfRowFreeze();

	public int getNumOfColumnFreeze();

	public void setNumOfRowFreeze(int num);

	public void setNumOfColumnFreeze(int num);

	public boolean isDisplayGridline();

	public void setDisplayGridline(boolean enable);
}
