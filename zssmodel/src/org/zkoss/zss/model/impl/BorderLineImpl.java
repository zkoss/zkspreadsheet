/* BorderLine.java

	Purpose:
		
	Description:
		
	History:
		Mar 31, 2015 7:13:20 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SBorderLine;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SBook;

/**
 * A border line.
 * @author henri
 * @since 3.8.0
 */
public class BorderLineImpl extends AbstractBorderLineAdv implements SBorderLine {
	private static final long serialVersionUID = -541335216584161785L;
	
	private BorderType type;
	private SColor color;
	private boolean showUp;
	private boolean showDown;
	
	//ZSS-977
	public BorderLineImpl(BorderType type, String htmlColor) {
		this(type, htmlColor == null ? ColorImpl.BLACK : new ColorImpl(htmlColor), false, false);
	}
	//ZSS-977
	public BorderLineImpl(BorderType type, SColor color) {
		this(type, color, false, false);
	}

	public BorderLineImpl(BorderType type, SColor color, boolean showUp, boolean showDown) {
		this.type = type;
		this.color = color;
		this.showUp = showUp;
		this.showDown = showDown;
	}

	public BorderType getBorderType() {
		return type;
	}
	
	public void setBorderType(BorderType type) {
		this.type = type;
	}
	
	public SColor getColor() {
		return color == null ? ColorImpl.BLACK : color; //ZSS-1185
	}
	
	public void setColor(SColor color) {
		this.color = color;
	}
	
	public boolean isShowDiagonalUpBorder() {
		return showUp;
	}
	
	public void setShowDiagonalUpBorder(boolean show) {
		showUp = show;
	}
	
	public boolean isShowDiagonalDownBorder() {
		return showDown;
	}
	
	public void setShowDiagonalDownBorder(boolean show) {
		showDown = show;
	}
	
	/*package*/ String getStyleKey() {
		return new StringBuilder()
			.append(type == null ? "" : type.ordinal())
			.append(".").append(color == null ? "" : color.getHtmlColor())
			.append(".").append(showUp ? "1" : "0")
			.append(".").append(showDown ? "1" : "0").toString();
	}
	
	//ZSS-1183
	//@since 3.9.0
	@Override
	/*package*/ SBorderLine cloneBorderLine(SBook book) {
		final AbstractColorAdv srcColor = (AbstractColorAdv)this.getColor(); 
		final SColor color = srcColor == null ? null : srcColor.cloneColor(book);  
		return new BorderLineImpl(this.type, color, this.showUp, this.showDown);		
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;

		BorderLineImpl that = (BorderLineImpl) o;

		if (showUp != that.showUp) return false;
		if (showDown != that.showDown) return false;
		if (type != that.type) return false;
		return color != null ? color.equals(that.color) : that.color == null;
	}

	@Override
	public int hashCode() {
		int result = type != null ? type.hashCode() : 0;
		result = 31 * result + (color != null ? color.hashCode() : 0);
		result = 31 * result + (showUp ? 1 : 0);
		result = 31 * result + (showDown ? 1 : 0);
		return result;
	}
}
