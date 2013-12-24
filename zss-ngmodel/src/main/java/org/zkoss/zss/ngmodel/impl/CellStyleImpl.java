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
	private NColor backgroundColor = ColorImpl.WHITE;
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
		return backgroundColor;
	}

	@Override
	public void setFillColor(NColor backgroundColor) {
		Validations.argNotNull(backgroundColor);
		this.backgroundColor = backgroundColor;
	}

	@Override
	public FillPattern getFillPattern() {
		return fillPattern;
	}

	@Override
	public void setFillPattern(FillPattern fillPattern) {
		Validations.argNotNull(backgroundColor);
		this.fillPattern = fillPattern;
	}

	@Override
	public Alignment getAlignment() {
		return alignment;
	}

	@Override
	public void setAlignment(Alignment alignment) {
		Validations.argNotNull(backgroundColor);
		this.alignment = alignment;
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return verticalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		Validations.argNotNull(backgroundColor);
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
		Validations.argNotNull(backgroundColor);
		this.borderLeft = borderLeft;
	}

	@Override
	public BorderType getBorderTop() {
		return borderTop;
	}

	@Override
	public void setBorderTop(BorderType borderTop) {
		Validations.argNotNull(backgroundColor);
		this.borderTop = borderTop;
	}

	@Override
	public BorderType getBorderRight() {
		return borderRight;
	}

	@Override
	public void setBorderRight(BorderType borderRight) {
		Validations.argNotNull(backgroundColor);
		this.borderRight = borderRight;
	}

	@Override
	public BorderType getBorderBottom() {
		return borderBottom;
	}

	@Override
	public void setBorderBottom(BorderType borderBottom){
		Validations.argNotNull(backgroundColor);
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
	public void setDataFormat(String dataFormat) {
		Validations.argNotNull(dataFormat);
		this.dataFormat = dataFormat;
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
}
