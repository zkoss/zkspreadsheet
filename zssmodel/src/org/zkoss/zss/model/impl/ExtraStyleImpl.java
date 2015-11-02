/* ExtraStyleImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 3:58:03 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SBorder;
import org.zkoss.zss.model.SBorderLine;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.util.Strings;

/**
 * @author henri
 * @since 3.8.2
 */
public class ExtraStyleImpl extends CellStyleImpl implements SExtraStyle {
	private static final long serialVersionUID = -3304797955338410853L;

	public ExtraStyleImpl(SFont font, SFill fill, SBorder border, String dataFormat) {
		super((AbstractFontAdv)font, (AbstractFillAdv) fill, (AbstractBorderAdv) border);
		if (dataFormat != null && !Strings.isBlank(dataFormat))
			setDataFormat(dataFormat);
	}

	//ZSS-1145
	@Override
	public BorderType getBorderLeft() {
		if (_border == null) return null;
		SBorderLine bline = _border.getLeftLine();
		return bline == null ? null : bline.getBorderType();
	}

	//ZSS-1145
	@Override
	public BorderType getBorderTop() {
		if (_border == null) return null;
		SBorderLine bline = _border.getTopLine();
		return bline == null ? null : bline.getBorderType();
	}

	//ZSS-1145
	@Override
	public BorderType getBorderRight() {
		if (_border == null) return null;
		SBorderLine bline = _border.getRightLine();
		return bline == null ? null : bline.getBorderType();
	}
	
	//ZSS-1145
	@Override
	public BorderType getBorderBottom() {
		if (_border == null) return null;
		SBorderLine bline = _border.getBottomLine();
		return bline == null ? null : bline.getBorderType();
	}
	
	//ZSS-1145
	@Override
	public BorderType getBorderDiagonal() {
		if (_border == null) return null;
		SBorderLine bline = _border.getDiagonalLine();
		return bline == null ? null : bline.getBorderType();
	}

	//ZSS-1145
	@Override
	public BorderType getBorderVertical() {
		if (_border == null) return null;
		SBorderLine bline = _border.getVerticalLine();
		return bline == null ? null : bline.getBorderType();
	}

	//ZSS-1145
	@Override
	public BorderType getBorderHorizontal() {
		if (_border == null) return null;
		SBorderLine bline = _border.getHorizontalLine();
		return bline == null ? null : bline.getBorderType();
	}

	//ZSS-1145
	@Override
	public String getDataFormat() {
		return _dataFormat;
	}
}
