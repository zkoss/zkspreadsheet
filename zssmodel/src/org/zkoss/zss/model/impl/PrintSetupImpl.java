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

import java.io.Serializable;

import org.zkoss.zss.model.SPrintSetup;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public class PrintSetupImpl implements SPrintSetup,Serializable {
	private static final long serialVersionUID = 1L;
	private boolean _printGridline = false; 
	private int _headerMargin;
	private int _footerMargin;
	private int _leftMargin;
	private int _rightMargin;
	private int _topMargin;
	private int _bottomMargin;
	
	private boolean _landscape = false;
	// private short scale = 100;
	private PaperSize _paperSize = PaperSize.A4;
	
	
	@Override
	public boolean isPrintGridlines() {
		return _printGridline;
	}

	@Override
	public void setPrintGridlines(boolean enable) {
		_printGridline = enable;
	}

	public int getHeaderMargin() {
		return _headerMargin;
	}

	public void setHeaderMargin(int headerMargin) {
		this._headerMargin = headerMargin;
	}

	public int getFooterMargin() {
		return _footerMargin;
	}

	public void setFooterMargin(int footerMargin) {
		this._footerMargin = footerMargin;
	}

	public int getLeftMargin() {
		return _leftMargin;
	}

	public void setLeftMargin(int leftMargin) {
		this._leftMargin = leftMargin;
	}

	public int getRightMargin() {
		return _rightMargin;
	}

	public void setRightMargin(int rightMargin) {
		this._rightMargin = rightMargin;
	}

	public int getTopMargin() {
		return _topMargin;
	}

	public void setTopMargin(int topMargin) {
		this._topMargin = topMargin;
	}

	public int getBottomMargin() {
		return _bottomMargin;
	}

	public void setBottomMargin(int bottomMargin) {
		this._bottomMargin = bottomMargin;
	}

	@Override
	public void setPaperSize(PaperSize size) {
		this._paperSize = size;
	}

	@Override
	public PaperSize getPaperSize() {
		return _paperSize;
	}

	@Override
	public void setLandscape(boolean landscape) {
		this._landscape = landscape;
	}

	@Override
	public boolean isLandscape() {
		return _landscape;
	}

//	@Override
//	public void setScale(short scale) {
//		this.scale = scale;
//	}
//
//	@Override
//	public short getScale() {
//		return scale;
//	}

}
