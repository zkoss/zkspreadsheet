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

import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class FontImpl extends FontAdv {
	private static final long serialVersionUID = 1L;


	public static final String FORMAT_GENERAL = "General";
	/**
     * By default, Microsoft Office Excel 2007 uses the Calibry font in font size 11
     */
	private String fontName = "Calibri";
	
	 /**
     * By default, Microsoft Office Excel 2007 uses the Calibry font in font size 11
     */
	private int fontHeightPoint = 11;
	
	private NColor fontColor = ColorImpl.BLACK;
	
	private Boldweight fontBoldweight = Boldweight.NORMAL;
	
	private boolean fontItalic = false;
	private boolean fontStrikeout = false;
	private TypeOffset fontTypeOffset = TypeOffset.NONE;
	private Underline fontUnderline = Underline.NONE;

	@Override
	public String getName() {
		return fontName;
	}

	@Override
	public void setName(String fontName) {
		this.fontName = fontName;
	}

	@Override
	public NColor getColor() {
		return fontColor;
	}

	@Override
	public void setColor(NColor fontColor) {
		Validations.argNotNull(fontColor);
		this.fontColor = fontColor;
	}

	@Override
	public Boldweight getBoldweight() {
		return fontBoldweight;
	}

	@Override
	public void setBoldweight(Boldweight fontBoldweight) {
		Validations.argNotNull(fontBoldweight);
		this.fontBoldweight = fontBoldweight;
	}

	@Override
	public int getHeightPoints() {
		return fontHeightPoint;
	}

	@Override
	public void setHeightPoints(int fontHeightPoint) {
		this.fontHeightPoint = fontHeightPoint;
	}

	@Override
	public boolean isItalic() {
		return fontItalic;
	}

	@Override
	public void setItalic(boolean fontItalic) {
		this.fontItalic = fontItalic;
	}

	@Override
	public boolean isStrikeout() {
		return fontStrikeout;
	}

	@Override
	public void setStrikeout(boolean fontStrikeout) {
		this.fontStrikeout = fontStrikeout;
	}

	@Override
	public TypeOffset getTypeOffset() {
		return fontTypeOffset;
	}

	@Override
	public void setTypeOffset(TypeOffset fontTypeOffset) {
		Validations.argNotNull(fontTypeOffset);
		this.fontTypeOffset = fontTypeOffset;
	}

	@Override
	public Underline getUnderline() {
		return fontUnderline;
	}

	@Override
	public void setUnderline(Underline fontUnderline) {
		Validations.argNotNull(fontUnderline);
		this.fontUnderline = fontUnderline;
	}

	

	@Override
	public void copyTo(FontAdv dest) {
		if (dest == this)
			return;
		Validations.argInstance(dest, FontImpl.class);
		FontImpl another = (FontImpl) dest;
		another.fontName = fontName;
		another.fontColor = fontColor;
		another.fontBoldweight = fontBoldweight;
		another.fontHeightPoint = fontHeightPoint;
		another.fontItalic = fontItalic;
		another.fontStrikeout = fontStrikeout;
		another.fontTypeOffset = fontTypeOffset;
		another.fontUnderline = fontUnderline;
	}
}
