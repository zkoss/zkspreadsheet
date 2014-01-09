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
public interface NPrintSetup {
	
	
	
	public boolean isPrintGridline();
	public void setPrintGridline(boolean enable);
	
	public int getHeaderMargin();
	public void setHeaderMargin(int px);
	public int getFooterMargin();
	public void setFooterMargin(int px);
	public int getLeftMargin();
	public void setLeftMargin(int px);
	public int getRightMargin();
	public void setRightMargin(int px);
	public int getTopMargin();
	public void setTopMargin(int px);
	public int getBottomMargin();
	public void setBottomMargin(int px);
}
