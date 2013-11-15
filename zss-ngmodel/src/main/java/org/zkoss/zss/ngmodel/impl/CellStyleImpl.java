package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;

public class CellStyleImpl implements NCellStyle {

	String fontName;
	String fontColor;
	String backgroundColor;
	
	public CellStyleImpl(){
		fontName = "Arial";
		fontColor = "#000000";
		backgroundColor = "#FFFFFF";
	}

	public String getFontName() {
		return fontName;
	}

	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	public String getFontColor() {
		return fontColor;
	}

	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	public String getBackgroundColor() {
		return backgroundColor;
	}

	public void setBackgroundColor(String backgroundColor) {
		this.backgroundColor = backgroundColor;
	}
}
