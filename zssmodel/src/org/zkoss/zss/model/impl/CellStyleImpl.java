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
package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class CellStyleImpl extends AbstractCellStyleAdv {
	private static final long serialVersionUID = 1L;

	private AbstractFontAdv _font;
	private SColor _fillColor = ColorImpl.WHITE;
	private SColor _backColor = ColorImpl.WHITE;
	private FillPattern _fillPattern = FillPattern.NO_FILL;
	private Alignment _alignment = Alignment.GENERAL;
	private VerticalAlignment _verticalAlignment = VerticalAlignment.BOTTOM;
	private boolean _wrapText = false;

	private BorderType _borderLeft = BorderType.NONE;
	private BorderType _borderTop = BorderType.NONE;
	private BorderType _borderRight = BorderType.NONE;
	private BorderType _borderBottom = BorderType.NONE;
	private SColor _borderTopColor = ColorImpl.BLACK;
	private SColor _borderLeftColor = ColorImpl.BLACK;
	private SColor _borderBottomColor = ColorImpl.BLACK;
	private SColor _borderRightColor = ColorImpl.BLACK;

	private String _dataFormat = FORMAT_GENERAL;
	private boolean _directFormat = false;
	private boolean _locked = true;// default locked as excel.
	private boolean _hidden = false;

	public CellStyleImpl(AbstractFontAdv font){
		this._font = font;
	}
	
	public SFont getFont(){
		return _font;
	}
	
	public void setFont(SFont font){
		Validations.argInstance(font, AbstractFontAdv.class);
		this._font = (AbstractFontAdv)font;
	}

	@Override
	public SColor getFillColor() {
		return _fillColor;
	}

	@Override
	public void setFillColor(SColor fillColor) {
		Validations.argNotNull(fillColor);
		this._fillColor = fillColor;
	}

	@Override
	public FillPattern getFillPattern() {
		return _fillPattern;
	}

	@Override
	public void setFillPattern(FillPattern fillPattern) {
		Validations.argNotNull(fillPattern);
		this._fillPattern = fillPattern;
	}

	@Override
	public Alignment getAlignment() {
		return _alignment;
	}

	@Override
	public void setAlignment(Alignment alignment) {
		Validations.argNotNull(alignment);
		this._alignment = alignment;
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return _verticalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		Validations.argNotNull(verticalAlignment);
		this._verticalAlignment = verticalAlignment;
	}

	@Override
	public boolean isWrapText() {
		return _wrapText;
	}

	@Override
	public void setWrapText(boolean wrapText) {
		this._wrapText = wrapText;
	}

	@Override
	public BorderType getBorderLeft() {
		return _borderLeft;
	}

	@Override
	public void setBorderLeft(BorderType borderLeft) {
		Validations.argNotNull(borderLeft);
		this._borderLeft = borderLeft;
	}

	@Override
	public BorderType getBorderTop() {
		return _borderTop;
	}

	@Override
	public void setBorderTop(BorderType borderTop) {
		Validations.argNotNull(borderTop);
		this._borderTop = borderTop;
	}

	@Override
	public BorderType getBorderRight() {
		return _borderRight;
	}

	@Override
	public void setBorderRight(BorderType borderRight) {
		Validations.argNotNull(borderRight);
		this._borderRight = borderRight;
	}

	@Override
	public BorderType getBorderBottom() {
		return _borderBottom;
	}

	@Override
	public void setBorderBottom(BorderType borderBottom){
		Validations.argNotNull(borderBottom);
		this._borderBottom = borderBottom;
	}

	@Override
	public SColor getBorderTopColor() {
		return _borderTopColor;
	}

	@Override
	public void setBorderTopColor(SColor borderTopColor) {
		Validations.argNotNull(borderTopColor);
		this._borderTopColor = borderTopColor;
	}

	@Override
	public SColor getBorderLeftColor() {
		return _borderLeftColor;
	}

	@Override
	public void setBorderLeftColor(SColor borderLeftColor) {
		Validations.argNotNull(borderLeftColor);
		this._borderLeftColor = borderLeftColor;
	}

	@Override
	public SColor getBorderBottomColor() {
		return _borderBottomColor;
	}

	@Override
	public void setBorderBottomColor(SColor borderBottomColor) {
		Validations.argNotNull(borderBottomColor);
		this._borderBottomColor = borderBottomColor;
	}

	@Override
	public SColor getBorderRightColor() {
		return _borderRightColor;
	}

	@Override
	public void setBorderRightColor(SColor borderRightColor) {
		Validations.argNotNull(borderRightColor);
		this._borderRightColor = borderRightColor;
	}

	@Override
	public String getDataFormat() {
		return _dataFormat;
	}
	
	@Override
	public boolean isDirectDataFormat(){
		return _directFormat;
	}

	@Override
	public void setDataFormat(String dataFormat) {
		//set to general if null to compatible with 3.0
		if(dataFormat==null || "".equals(dataFormat.trim())){
			dataFormat = FORMAT_GENERAL;
		}
		this._dataFormat = dataFormat;
		_directFormat = false;
	}
	
	@Override
	public void setDirectDataFormat(String dataFormat){
		setDataFormat(dataFormat);
		_directFormat = true;
	}

	@Override
	public boolean isLocked() {
		return _locked;
	}

	@Override
	public void setLocked(boolean locked) {
		this._locked = locked;
	}

	@Override
	public boolean isHidden() {
		return _hidden;
	}

	@Override
	public void setHidden(boolean hidden) {
		this._hidden = hidden;
	}

	@Override
	public void copyFrom(SCellStyle src) {
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
		sb.append(_font.getStyleKey())
		.append(".").append(_fillPattern.ordinal())
		.append(".").append(_fillColor.getHtmlColor())
		.append(".").append(_alignment.ordinal())
		.append(".").append(_verticalAlignment.ordinal())
		.append(".").append(_wrapText?"T":"F")
		.append(".").append(_borderLeft.ordinal())
		.append(".").append(_borderLeftColor.getHtmlColor())
		.append(".").append(_borderRight.ordinal())
		.append(".").append(_borderRightColor.getHtmlColor())
		.append(".").append(_borderTop.ordinal())
		.append(".").append(_borderTopColor.getHtmlColor())
		.append(".").append(_borderBottom.ordinal())
		.append(".").append(_borderBottomColor.getHtmlColor())
		.append(".").append(_dataFormat)
		.append(".").append(_locked?"T":"F")
		.append(".").append(_hidden?"T":"F");
		return sb.toString();
	}

	@Override
	public void setBorderLeft(BorderType borderLeft, SColor color) {
		setBorderLeft(borderLeft);
		setBorderLeftColor(color);
	}

	@Override
	public void setBorderTop(BorderType borderTop, SColor color) {
		setBorderTop(borderTop);
		setBorderTopColor(color);
	}

	@Override
	public void setBorderRight(BorderType borderRight, SColor color) {
		setBorderRight(borderRight);
		setBorderRightColor(color);
	}

	@Override
	public void setBorderBottom(BorderType borderBottom, SColor color) {
		setBorderBottom(borderBottom);
		setBorderBottomColor(color);
	}

	//ZSS-780
	@Override
	public SColor getBackgroundColor() {
		// TODO Auto-generated method stub
		return _backColor;
	}

	//ZSS-780
	@Override
	public void setBackgroundColor(SColor backColor) {
		Validations.argNotNull(backColor);
		this._backColor = backColor;
	}
}
