/* CellStyle.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 10:28:37 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.FontUnderline;
import org.zkoss.zss.model.impl.BookHelper;


/**
 * @author Ian Tsai / Sam
 *
 */
public interface CellStyle {
	
	/**
	 * Align
	 */
	public static int ALIGN_LEFT = org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
	public static int ALIGN_CENTER = org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
	public static int ALIGN_RIGHT = org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_RIGHT;
	public static int ALIGN_TOP = org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_TOP;
	public static int ALIGN_MIDDLE = org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_CENTER;
	public static int ALIGN_BOTTOM = org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_BOTTOM;
	
	/**
	 * Border position
	 */
	public static int BORDER_EDGE_RIGHT = BookHelper.BORDER_EDGE_RIGHT;
	public static int BORDER_EDGE_TOP = BookHelper.BORDER_EDGE_TOP;
	public static int BORDER_EDGE_LEFT = BookHelper.BORDER_EDGE_LEFT;
	public static int BORDER_INSIDE_HORIZONTAL = BookHelper.BORDER_INSIDE_HORIZONTAL;
	public static int BORDER_INSIDE_VERTICAL = BookHelper.BORDER_INSIDE_VERTICAL;
	public static int BORDER_DIAGONAL_DOWN = BookHelper.BORDER_DIAGONAL_DOWN;
	public static int BORDER_DIAGONAL_UP = BookHelper.BORDER_DIAGONAL_UP;
	public static int BORDER_FULL = BookHelper.BORDER_FULL;
	public static int BORDER_OUTLINE = BookHelper.BORDER_OUTLINE;
	public static int BORDER_INSIDE = BookHelper.BORDER_INSIDE;
	public static int BORDER_DIAGONAL = BookHelper.BORDER_DIAGONAL;
	
	/**
	 * Underline style
	 */
	public static int UNDERLINE_SINGLE = FontUnderline.SINGLE.getValue();
	public static int UNDERLINE_DOUBLE = FontUnderline.DOUBLE.getValue();
	public static int UNDERLINE_SINGLE_ACCOUNTING = FontUnderline.SINGLE_ACCOUNTING.getValue();
	public static int UNDERLINE_DOUBLE_ACCOUNTING =  FontUnderline.DOUBLE_ACCOUNTING.getValue();
	public static int UNDERLINE_NONE = FontUnderline.NONE.getValue();
	
	/**
	 * 
	 * @param family
	 */
	public void setFontFamily(String family);
	/**
	 * 
	 * @return
	 */
	public String getFontFamily();
	
	/**
	 * 
	 * @param 
	 */
	public void setFontSize(int size);
	/**
	 * 
	 * @return
	 */
	public int getFontSize();
	
	public void setFontColor(String color);
	
	public String getFontColor();
	
	/**
	 * 
	 * @param bold
	 */
	public void setBold(boolean bold);
	/**
	 * 
	 * @return
	 */
	public boolean isBold();
	
	public void setItalic(boolean italic);
	
	public boolean isItalic();
	
	public void setUnderline(int underline);
	
	public int getUnderline();

	
	public void setBorder(int borderPosition, BorderStyle borderStyle, String color);
	
	//TODO getBottomBorder();
//	public int getBorder();
	
	public void setAlignment(int alignment);
	
	public int getAlignment();

	public void setVerticalAlignment(int alignment);
	
	public int getVerticalAlignment();
	
	public void setStrikethrough(boolean strikethrough);
	/**
	 * @return
	 */
	public boolean isStrikethrough();
	
	public void setCellColor(String color);
	/**
	 * @return
	 */
	public String getCellColor();
	/**
	 * @param isWrapText
	 */
	public void setWrapText(boolean isWrapText);
	
	public boolean isWrapText();
}
