/* EditableFont.java

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
 * The editable font
 * @author dennis
 * @since 3.0.0
 */
public interface EditableFont extends Font{

	public void setFontName(String fontName);
	public void setBoldweight(Boldweight boldweight);
	public void setItalic(boolean italic);
	public void setStrikeout(boolean strikeout);
	public void setUnderline(Underline underline);
	public void setFontHeight(int height);
	public void setFontHeightInPoint(int point);
	public void setColor(Color color);
	
}
