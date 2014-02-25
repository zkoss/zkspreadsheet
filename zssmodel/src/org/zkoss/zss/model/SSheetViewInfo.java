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
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface SSheetViewInfo {
	public int getNumOfRowFreeze();

	public int getNumOfColumnFreeze();

	public void setNumOfRowFreeze(int num);

	public void setNumOfColumnFreeze(int num);

	public boolean isDisplayGridline();

	public void setDisplayGridline(boolean enable);
	
	public SHeader getHeader();
	
	public SFooter getFooter();
	

	public int[] getRowBreaks();
	public void setRowBreaks(int[] breaks);
	public void addRowBreak(int rowIdx);
	public int[] getColumnBreaks();
	public void setColumnBreaks(int[] breaks);
	public void addColumnBreak(int columnIdx);
}
