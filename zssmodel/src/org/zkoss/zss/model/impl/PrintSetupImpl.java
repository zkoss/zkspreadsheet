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

public class PrintSetupImpl implements SPrintSetup,Serializable {
	private static final long serialVersionUID = 1L;
	private boolean printGridline = false; 
	private int headerMargin;
	private int footerMargin;
	private int leftMargin;
	private int rightMargin;
	private int topMargin;
	private int bottomMargin;
	
	private boolean landscape = false;
	// private short scale = 100;
	private PaperSize paperSize = PaperSize.A4;
	
	
	@Override
	public boolean isPrintGridlines() {
		return printGridline;
	}

	@Override
	public void setPrintGridlines(boolean enable) {
		printGridline = enable;
	}

	public int getHeaderMargin() {
		return headerMargin;
	}

	public void setHeaderMargin(int headerMargin) {
		this.headerMargin = headerMargin;
	}

	public int getFooterMargin() {
		return footerMargin;
	}

	public void setFooterMargin(int footerMargin) {
		this.footerMargin = footerMargin;
	}

	public int getLeftMargin() {
		return leftMargin;
	}

	public void setLeftMargin(int leftMargin) {
		this.leftMargin = leftMargin;
	}

	public int getRightMargin() {
		return rightMargin;
	}

	public void setRightMargin(int rightMargin) {
		this.rightMargin = rightMargin;
	}

	public int getTopMargin() {
		return topMargin;
	}

	public void setTopMargin(int topMargin) {
		this.topMargin = topMargin;
	}

	public int getBottomMargin() {
		return bottomMargin;
	}

	public void setBottomMargin(int bottomMargin) {
		this.bottomMargin = bottomMargin;
	}

	@Override
	public void setPaperSize(PaperSize size) {
		this.paperSize = size;
	}

	@Override
	public PaperSize getPaperSize() {
		return paperSize;
	}

	@Override
	public void setLandscape(boolean landscape) {
		this.landscape = landscape;
	}

	@Override
	public boolean isLandscape() {
		return landscape;
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
