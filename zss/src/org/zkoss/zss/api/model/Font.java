/* Font.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public interface Font {
	
	public enum TypeOffset{
		NONE, 
		SUPER, 
		SUB
	}
	public enum Underline{
		NONE,
		SINGLE,
		DOUBLE,
		SINGLE_ACCOUNTING,
		DOUBLE_ACCOUNTING
	}
	
	public enum Boldweight{
		NORMAL,
		BOLD
	}
	
	
	public Color getColor();
	public String getFontName();
	public Boldweight getBoldweight();
	public int getFontHeight();
	public int getFontHeightInPoint();
	public boolean isItalic();
	public boolean isStrikeout();
	public TypeOffset getTypeOffset();
	public Underline getUnderline();
	public void setFontName(String fontName);
	public void setBoldweight(Boldweight boldweight);
	public void setItalic(boolean italic);
	public void setStrikeout(boolean strikeout);
	public void setUnderline(Underline underline);
	public void setFontHeight(int height);
	public void setFontHeightInPoint(int point);
	public void setColor(Color color);
	
}
