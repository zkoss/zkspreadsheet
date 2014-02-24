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
public interface NSheetViewInfo {
	public int getNumOfRowFreeze();

	public int getNumOfColumnFreeze();

	public void setNumOfRowFreeze(int num);

	public void setNumOfColumnFreeze(int num);

	public boolean isDisplayGridline();

	public void setDisplayGridline(boolean enable);
	
	public NHeader getHeader();
	
	public NFooter getFooter();
	

	public int[] getRowBreaks();
	public void setRowBreaks(int[] breaks);
	public void addRowBreak(int rowIdx);
	public int[] getColumnBreaks();
	public void setColumnBreaks(int[] breaks);
	public void addColumnBreak(int columnIdx);
}
