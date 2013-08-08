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
	
	protected ModelRef<XBook> _bookRef;
	protected ModelRef<org.zkoss.poi.ss.usermodel.Font> _fontRef;
	
	public FontImpl(ModelRef<XBook> book, ModelRef<org.zkoss.poi.ss.usermodel.Font> font) {
		this._bookRef = book;
		this._fontRef = font;
	}
	public String getFontName() {
		return getNative().getFontName();
	}
	public org.zkoss.poi.ss.usermodel.Font getNative() {
		return _fontRef.get();
	}
	public ModelRef<org.zkoss.poi.ss.usermodel.Font> getRef(){
		return _fontRef;
	}
	
	public Color getColor(){
		org.zkoss.poi.ss.usermodel.Color c = BookHelper.getFontColor(_bookRef.get(), _fontRef.get());
		return new ColorImpl(_bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(c));
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_fontRef == null) ? 0 : _fontRef.hashCode());
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
		if (_fontRef == null) {
			if (other._fontRef != null)
				return false;
		} else if (!_fontRef.equals(other._fontRef))
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
