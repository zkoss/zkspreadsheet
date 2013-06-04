/* FontImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class FontImpl implements Font{
	
	ModelRef<XBook> bookRef;
	ModelRef<org.zkoss.poi.ss.usermodel.Font> fontRef;
	
	public FontImpl(ModelRef<XBook> book, ModelRef<org.zkoss.poi.ss.usermodel.Font> font) {
		this.bookRef = book;
		this.fontRef = font;
	}
	public String getFontName() {
		return getNative().getFontName();
	}
	public org.zkoss.poi.ss.usermodel.Font getNative() {
		return fontRef.get();
	}
	public ModelRef<org.zkoss.poi.ss.usermodel.Font> getRef(){
		return fontRef;
	}
	
	public Color getColor(){
		org.zkoss.poi.ss.usermodel.Color c = BookHelper.getFontColor(bookRef.get(), fontRef.get());
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(c));
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((fontRef == null) ? 0 : fontRef.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		FontImpl other = (FontImpl) obj;
		if (fontRef == null) {
			if (other.fontRef != null)
				return false;
		} else if (!fontRef.equals(other.fontRef))
			return false;
		return true;
	}
	public Boldweight getBoldweight() {
		return EnumUtil.toFontBoldweight(getNative().getBoldweight());
	}
	public int getFontHeight() {
		return getNative().getFontHeight();
	}
	public boolean isItalic() {
		return getNative().getItalic();
	}
	public boolean isStrikeout() {
		return getNative().getStrikeout();
	}
	public TypeOffset getTypeOffset() {
		return EnumUtil.toFontTypeOffset(getNative().getTypeOffset());
	}
	public Underline getUnderline() {
		return EnumUtil.toFontUnderline(getNative().getUnderline());
	}

	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(getNative()).append("]");
		sb.append(getFontName());
		return sb.toString();
	}

	public int getFontHeightInPoint() {
		return UnitUtil.twipToPoint(getFontHeight());
	}
}
