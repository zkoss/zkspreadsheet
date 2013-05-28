/* CellStyle.java

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

import org.zkoss.zss.api.Range.CellStyleHelper;

/**
 * The cell style object.
 * @author dennis
 * @since 3.0.0
 */
public interface CellStyle {

	/**
	 * The background color fill pattern
	 */
	public enum FillPattern {
		NO_FILL, SOLID_FOREGROUND, FINE_DOTS, ALT_BARS, SPARSE_DOTS, THICK_HORZ_BANDS, THICK_VERT_BANDS, THICK_BACKWARD_DIAG, THICK_FORWARD_DIAG, BIG_SPOTS, BRICKS, THIN_HORZ_BANDS, THIN_VERT_BANDS, THIN_BACKWARD_DIAG, THIN_FORWARD_DIAG, SQUARES, DIAMONDS, LESS_DOTS, LEAST_DOTS
	}

	/**
	 * The horizontal alignment
	 *
	 */
	public enum Alignment {
		GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY, CENTER_SELECTION
	}

	/**
	 * The vertical alignment
	 */
	public enum VerticalAlignment {
		TOP, CENTER, BOTTOM, JUSTIFY
	}

	/**
	 * The border type
	 */
	public enum BorderType {
		NONE, THIN, MEDIUM, DASHED, HAIR, THICK, DOUBLE, DOTTED, MEDIUM_DASHED, DASH_DOT, MEDIUM_DASH_DOT, DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT;
	}

	/**
	 * @return the font
	 */
	public Font getFont();

	/**
	 * Sets font
	 * @param font the font
	 */
	public void setFont(Font font);

	/**
	 * @return background-color
	 */
	public Color getBackgroundColor();

	/**
	 * Sets background-color
	 * @param color background-color
	 */
	public void setBackgroundColor(Color color);

	/**
	 * Gets the fill/background pattern <br/>
	 * @return the fill pattern
	 */
	public FillPattern getFillPattern();

	/**
	 * Sets the fill/background pattern <br/>
	 * Note: Spreadsheet (UI display) supports only {@link FillPattern#NO_FILL} and {@link FillPattern#SOLID_FOREGROUND} 
	 * (Other pattern will be display as {@link FillPattern#SOLID_FOREGROUND), 
	 * However you can still set another pattern, the data will still be kept when export. 
	 * @param pattern
	 */
	public void setFillPattern(FillPattern pattern);

	/**
	 * Sets horizontal alignment <br/>
	 * Note: Spreadsheet(UI display) supports only {@link Alignment#LEFT}, {@link Alignment#CENTER}, {@link Alignment#RIGHT}
	 * ({@link Alignment#CENTER_SELECTION} will be display as {@link Alignment#CENTER}, Other alignment will be display as  {@link Alignment#LEFT}),
	 * However you can still set another alignment, the data will still be kept when export.
	 * @param alignment
	 */
	public void setAlignment(Alignment alignment);

	/**
	 * Gets the horizontal alignment <br/>
	 * @return the horizontal alignment
	 */
	public Alignment getAlignment();

	/**
	 * Sets vertical alignment <br/>
	 * Note: Spreadsheet(UI display) supports only {@link VerticalAlignment#TOP}, {@link VerticalAlignment#CENTER}, {@link VerticalAlignment#TOP},
	 * (Other alignment will be display as  {@link VerticalAlignment#BOTTOM}),
	 * @param alignment
	 */
	public void setVerticalAlignment(VerticalAlignment alignment);

	/**
	 * Gets vertical alignment <br/>
	 * @return
	 */
	public VerticalAlignment getVerticalAlignment();

	/**
	 * @return true if wrap-text
	 */
	public boolean isWrapText();

	/**
	 * Sets wrap-text
	 * @param wraptext wrap-text
	 */
	public void setWrapText(boolean wraptext);

	/**
	 * Sets left border <br/>
	 * Note: Spreadsheet(UI display) only supports {@link BorderType#NONE}, {@link BorderType#THIN}, {@link BorderType#DOTTED} and {@link BorderType#HAIR},
	 * ({@link BorderType#DASH_DOT} will be display as {@link Alignment#DOTTED}, Other alignment will be display as  {@link Alignment#THIN}),
	 * However you can still set another alignment, the data will still be kept when export.
	 * @param borderType
	 */
	public void setBorderLeft(BorderType borderType);

	/**
	 * @return left border
	 */
	public BorderType getBorderLeft();

	/**
	 * Sets top border <br/>
	 * Note: Spreadsheet(UI display) only supports {@link BorderType#NONE}, {@link BorderType#THIN}, {@link BorderType#DOTTED} and {@link BorderType#HAIR},
	 * ({@link BorderType#DASH_DOT} will be display as {@link Alignment#DOTTED}, Other alignment will be display as  {@link Alignment#THIN}),
	 * However you can still set another alignment, the data will still be kept when export. 
	 * @param borderType
	 */
	public void setBorderTop(BorderType borderType);

	/**
	 * @return top border
	 */
	public BorderType getBorderTop();

	/**
	 * Sets right border <br/>
	 * Note: Spreadsheet(UI display) only supports {@link BorderType#NONE}, {@link BorderType#THIN}, {@link BorderType#DOTTED} and {@link BorderType#HAIR},
	 * ({@link BorderType#DASH_DOT} will be display as {@link Alignment#DOTTED}, Other alignment will be display as  {@link Alignment#THIN}),
	 * However you can still set another alignment, the data will still be kept when export. 
	 * @param borderType
	 */
	public void setBorderRight(BorderType borderType);

	/**
	 * @return right border
	 */
	public BorderType getBorderRight();

	/**
	 * Sets bottom border <br/>
	 * Note: Spreadsheet(UI display) only supports {@link BorderType#NONE}, {@link BorderType#THIN}, {@link BorderType#DOTTED} and {@link BorderType#HAIR},
	 * ({@link BorderType#DASH_DOT} will be display as {@link Alignment#DOTTED}, Other alignment will be display as  {@link Alignment#THIN}),
	 * However you can still set another alignment, the data will still be kept when export. 
	 * @param borderType
	 */
	public void setBorderBottom(BorderType borderType);

	/**
	 * @return bottom border
	 */
	public BorderType getBorderBottom();

	/**
	 * Sets top border color. <br/>
	 * you could use {@link CellStyleHelper#createColorFromHtmlColor(String)} to create a {@link Color}
	 * @param color
	 */
	public void setBorderTopColor(Color color);

	/**
	 * @return top border color
	 */
	public Color getBorderTopColor();

	/**
	 * Sets left border color<br/>
	 * you could use {@link CellStyleHelper#createColorFromHtmlColor(String)} to create a {@link Color}
	 * @param color
	 */
	public void setBorderLeftColor(Color color);

	/** 
	 * @return left border color
	 */
	public Color getBorderLeftColor();

	/**
	 * Sets bottom border color <br/>
	 * you could use {@link CellStyleHelper#createColorFromHtmlColor(String)} to create a {@link Color}
	 * @param color
	 */
	public void setBorderBottomColor(Color color);

	/**
	 * 
	 * @return bottom border color
	 */
	public Color getBorderBottomColor();

	/**
	 * Sets right border color<br/>
	 * you could use {@link CellStyleHelper#createColorFromHtmlColor(String)} to create a {@link Color}
	 * @param color
	 */
	public void setBorderRightColor(Color color);

	/**
	 * @return right border color
	 */
	public Color getBorderRightColor();
	
	/**
	 * @return data format
	 */
	public String getDataFormat();
	
	/**
	 * Sets data format
	 * @param format
	 */
	public void setDataFormat(String format);
	
	/**
	 * 
	 * @return true if locked
	 */
	public boolean isLocked();
	
	/**
	 * Sets locked
	 * @param locked
	 */
	public void setLocked(boolean locked);
	
	/**
	 * 
	 * @return true if hidden
	 */
	public boolean isHidden();
	
	/**
	 * Sets hidden
	 * @param hidden
	 */
	public void setHidden(boolean hidden);
	
}
