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
package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class CellStyleImpl extends AbstractCellStyleAdv {
	private static final long serialVersionUID = 1L;

	private AbstractFontAdv font;
	private NColor fillColor = ColorImpl.WHITE;
	private FillPattern fillPattern = FillPattern.NO_FILL;
	private Alignment alignment = Alignment.GENERAL;
	private VerticalAlignment verticalAlignment = VerticalAlignment.BOTTOM;
	private boolean wrapText = false;

	private BorderType borderLeft = BorderType.NONE;
	private BorderType borderTop = BorderType.NONE;
	private BorderType borderRight = BorderType.NONE;
	private BorderType borderBottom = BorderType.NONE;
	private NColor borderTopColor = ColorImpl.BLACK;
	private NColor borderLeftColor = ColorImpl.BLACK;
	private NColor borderBottomColor = ColorImpl.BLACK;
	private NColor borderRightColor = ColorImpl.BLACK;

	private String dataFormat = FORMAT_GENERAL;
	private boolean directFormat = false;
	private boolean locked = true;// default locked as excel.
	private boolean hidden = false;

	public CellStyleImpl(AbstractFontAdv font){
		this.font = font;
	}
	
	public NFont getFont(){
		return font;
	}
	
	public void setFont(NFont font){
		Validations.argInstance(font, AbstractFontAdv.class);
		this.font = (AbstractFontAdv)font;
	}

	@Override
	public NColor getFillColor() {
		return fillColor;
	}

	@Override
	public void setFillColor(NColor fillColor) {
		Validations.argNotNull(fillColor);
		this.fillColor = fillColor;
	}

	@Override
	public FillPattern getFillPattern() {
		return fillPattern;
	}

	@Override
	public void setFillPattern(FillPattern fillPattern) {
		Validations.argNotNull(fillPattern);
		this.fillPattern = fillPattern;
	}

	@Override
	public Alignment getAlignment() {
		return alignment;
	}

	@Override
	public void setAlignment(Alignment alignment) {
		Validations.argNotNull(alignment);
		this.alignment = alignment;
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return verticalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		Validations.argNotNull(verticalAlignment);
		this.verticalAlignment = verticalAlignment;
	}

	@Override
	public boolean isWrapText() {
		return wrapText;
	}

	@Override
	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	@Override
	public BorderType getBorderLeft() {
		return borderLeft;
	}

	@Override
	public void setBorderLeft(BorderType borderLeft) {
		Validations.argNotNull(borderLeft);
		this.borderLeft = borderLeft;
	}

	@Override
	public BorderType getBorderTop() {
		return borderTop;
	}

	@Override
	public void setBorderTop(BorderType borderTop) {
		Validations.argNotNull(borderTop);
		this.borderTop = borderTop;
	}

	@Override
	public BorderType getBorderRight() {
		return borderRight;
	}

	@Override
	public void setBorderRight(BorderType borderRight) {
		Validations.argNotNull(borderRight);
		this.borderRight = borderRight;
	}

	@Override
	public BorderType getBorderBottom() {
		return borderBottom;
	}

	@Override
	public void setBorderBottom(BorderType borderBottom){
		Validations.argNotNull(borderBottom);
		this.borderBottom = borderBottom;
	}

	@Override
	public NColor getBorderTopColor() {
		return borderTopColor;
	}

	@Override
	public void setBorderTopColor(NColor borderTopColor) {
		Validations.argNotNull(borderTopColor);
		this.borderTopColor = borderTopColor;
	}

	@Override
	public NColor getBorderLeftColor() {
		return borderLeftColor;
	}

	@Override
	public void setBorderLeftColor(NColor borderLeftColor) {
		Validations.argNotNull(borderLeftColor);
		this.borderLeftColor = borderLeftColor;
	}

	@Override
	public NColor getBorderBottomColor() {
		return borderBottomColor;
	}

	@Override
	public void setBorderBottomColor(NColor borderBottomColor) {
		Validations.argNotNull(borderBottomColor);
		this.borderBottomColor = borderBottomColor;
	}

	@Override
	public NColor getBorderRightColor() {
		return borderRightColor;
	}

	@Override
	public void setBorderRightColor(NColor borderRightColor) {
		Validations.argNotNull(borderRightColor);
		this.borderRightColor = borderRightColor;
	}

	@Override
	public String getDataFormat() {
		return dataFormat;
	}
	
	@Override
	public boolean isDirectDataFormat(){
		return directFormat;
	}

	@Override
	public void setDataFormat(String dataFormat) {
		//set to general if null to compatible with 3.0
		if(dataFormat==null || "".equals(dataFormat.trim())){
			dataFormat = FORMAT_GENERAL;
		}
		this.dataFormat = dataFormat;
		directFormat = false;
	}
	
	@Override
	public void setDirectDataFormat(String dataFormat){
		setDataFormat(dataFormat);
		directFormat = true;
	}

	@Override
	public boolean isLocked() {
		return locked;
	}

	@Override
	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	@Override
	public boolean isHidden() {
		return hidden;
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	@Override
	public void copyFrom(NCellStyle src) {
		if (src == this)
			return;
		Validations.argInstance(src, CellStyleImpl.class);
		setFont(src.getFont());//assign directly
		
		setFillColor(src.getFillColor());
		setFillPattern(src.getFillPattern());
		setAlignment(src.getAlignment());
		setVerticalAlignment(src.getVerticalAlignment());
		setWrapText(src.isWrapText());

		setBorderLeft(src.getBorderLeft());
		setBorderTop(src.getBorderTop());
		setBorderRight(src.getBorderRight());
		setBorderBottom(src.getBorderBottom());
		setBorderTopColor(src.getBorderTopColor());
		setBorderLeftColor(src.getBorderLeftColor());
		setBorderBottomColor(src.getBorderBottomColor());
		setBorderRightColor(src.getBorderRightColor());

		setDataFormat(src.getDataFormat());
		setLocked(src.isLocked());
		setHidden(src.isHidden());
	}
	
	@Override
	String getStyleKey() {
		StringBuilder sb = new StringBuilder();
		sb.append(font.getStyleKey())
		.append(".").append(fillPattern.ordinal())
		.append(".").append(fillColor.getHtmlColor())
		.append(".").append(alignment.ordinal())
		.append(".").append(verticalAlignment.ordinal())
		.append(".").append(wrapText?"T":"F")
		.append(".").append(borderLeft.ordinal())
		.append(".").append(borderLeftColor.getHtmlColor())
		.append(".").append(borderRight.ordinal())
		.append(".").append(borderRightColor.getHtmlColor())
		.append(".").append(borderTop.ordinal())
		.append(".").append(borderTopColor.getHtmlColor())
		.append(".").append(borderBottom.ordinal())
		.append(".").append(borderBottomColor.getHtmlColor())
		.append(".").append(dataFormat)
		.append(".").append(locked?"T":"F")
		.append(".").append(hidden?"T":"F");
		return sb.toString();
	}

	@Override
	public void setBorderLeft(BorderType borderLeft, NColor color) {
		setBorderLeft(borderLeft);
		setBorderLeftColor(color);
	}

	@Override
	public void setBorderTop(BorderType borderTop, NColor color) {
		setBorderTop(borderTop);
		setBorderTopColor(color);
	}

	@Override
	public void setBorderRight(BorderType borderRight, NColor color) {
		setBorderRight(borderRight);
		setBorderRightColor(color);
	}

	@Override
	public void setBorderBottom(BorderType borderBottom, NColor color) {
		setBorderBottom(borderBottom);
		setBorderBottomColor(color);
	}
}
