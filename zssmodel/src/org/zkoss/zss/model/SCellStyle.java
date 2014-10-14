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
 * Represent style information e.g. alignment, border, format pattern, font, color, fill pattern, wrap text, and hidden status. It may associate with a book or cells.
 * @author dennis
 * @since 3.5.0
 */
public interface SCellStyle {

	public static final String FORMAT_GENERAL = "General";
	
	/**
	 * The background color fill pattern
	 * @since 3.5.0
	 */
	public enum FillPattern {
		NO_FILL, SOLID_FOREGROUND, FINE_DOTS, ALT_BARS, SPARSE_DOTS, THICK_HORZ_BANDS, THICK_VERT_BANDS, THICK_BACKWARD_DIAG, THICK_FORWARD_DIAG, BIG_SPOTS, BRICKS, THIN_HORZ_BANDS, THIN_VERT_BANDS, THIN_BACKWARD_DIAG, THIN_FORWARD_DIAG, SQUARES, DIAMONDS, LESS_DOTS, LEAST_DOTS
	}

	/**
	 * The horizontal alignment
	 * @since 3.5.0
	 */
	public enum Alignment {
		GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY, CENTER_SELECTION
	}

	/**
	 * The vertical alignment
	 * @since 3.5.0
	 */
	public enum VerticalAlignment {
		TOP, CENTER, BOTTOM, JUSTIFY
	}

	/**
	 * The border type
	 * @since 3.5.0
	 */
	public enum BorderType {
		NONE, THIN, MEDIUM, DASHED, HAIR, THICK, DOUBLE, DOTTED, MEDIUM_DASHED, DASH_DOT, MEDIUM_DASH_DOT, DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT;
	}
	
	
	/**
	 * @return fill foreground-color
	 */
	public SColor getFillColor();

	/**
	 * @return fill background-color
	 * @since 3.6.0
	 */
	public SColor getBackColor();

	/**
	 * Gets the fill/background pattern <br/>
	 * @return the fill pattern
	 */
	public FillPattern getFillPattern();

	/**
	 * Gets the horizontal alignment <br/>
	 * @return the horizontal alignment
	 */
	public Alignment getAlignment();

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
	 * @return left border
	 */
	public BorderType getBorderLeft();

	/**
	 * @return top border
	 */
	public BorderType getBorderTop();

	/**
	 * @return right border
	 */
	public BorderType getBorderRight();

	/**
	 * @return bottom border
	 */
	public BorderType getBorderBottom();

	/**
	 * @return top border color
	 */
	public SColor getBorderTopColor();

	/** 
	 * @return left border color
	 */
	public SColor getBorderLeftColor();

	/**
	 * 
	 * @return bottom border color
	 */
	public SColor getBorderBottomColor();

	/**
	 * @return right border color
	 */
	public SColor getBorderRightColor();
	
	/**
	 * @return data format
	 */
	public String getDataFormat();
	
	/**
	 * 
	 * @return true if the data format is direct data format, which mean it will not care Locale when formatting
	 */
	public boolean isDirectDataFormat();
	
	/**
	 * 
	 * @return true if locked
	 */
	public boolean isLocked();
	
	/**
	 * 
	 * @return true if hidden
	 */
	public boolean isHidden();
	

	public void setFillColor(SColor fillColor);
	
	public void setBackgroundColor(SColor backColor); //ZSS-780

	public void setFillPattern(FillPattern fillPattern);
	
	public void setAlignment(Alignment alignment);
	
	public void setVerticalAlignment(VerticalAlignment verticalAlignment);

	public void setWrapText(boolean wrapText);

	public void setBorderLeft(BorderType borderLeft);
	public void setBorderLeft(BorderType borderLeft,SColor color);

	public void setBorderTop(BorderType borderTop);
	public void setBorderTop(BorderType borderTop,SColor color);

	public void setBorderRight(BorderType borderRight);
	public void setBorderRight(BorderType borderRight,SColor color);

	public void setBorderBottom(BorderType borderBottom);
	public void setBorderBottom(BorderType borderBottom,SColor color);

	public void setBorderTopColor(SColor borderTopColor);

	public void setBorderLeftColor(SColor borderLeftColor);

	public void setBorderBottomColor(SColor borderBottomColor);

	public void setBorderRightColor(SColor borderRightColor);

	public void setDataFormat(String dataFormat);
	
	public void setDirectDataFormat(String dataFormat);
	
	public void setLocked(boolean locked);

	public void setHidden(boolean hidden);
	
	public SFont getFont();
	
	public void setFont(SFont font);

	public void copyFrom(SCellStyle src);
}
