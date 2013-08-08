/* EditableFontImpl.java

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
import org.zkoss.zss.api.model.EditableFont;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class EditableFontImpl extends FontImpl implements EditableFont{
	

	public EditableFontImpl(ModelRef<XBook> book, ModelRef<org.zkoss.poi.ss.usermodel.Font> font) {
		super(book,font);
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
		EditableFontImpl other = (EditableFontImpl) obj;
		if (_fontRef == null) {
			if (other._fontRef != null)
				return false;
		} else if (!_fontRef.equals(other._fontRef))
			return false;
		return true;
	}
	
	public void copyAttributeFrom(Font src) {
		org.zkoss.poi.ss.usermodel.Font sfont = ((FontImpl)src).getNative();
		org.zkoss.poi.ss.usermodel.Font font = getNative();
		font.setBoldweight(sfont.getBoldweight());
		org.zkoss.poi.ss.usermodel.Color srcColor = BookHelper.getFontColor(_bookRef.get(), sfont);
		BookHelper.setFontColor(_bookRef.get(), font, srcColor);
		font.setFontHeight(sfont.getFontHeight());
		font.setFontName(sfont.getFontName());
		font.setItalic(sfont.getItalic());
		font.setStrikeout(sfont.getStrikeout());
		font.setTypeOffset(sfont.getTypeOffset());
		font.setUnderline(sfont.getUnderline());
	}
	public void setFontName(String fontName) {
		getNative().setFontName(fontName);
	}
	public void setBoldweight(Boldweight boldweight) {
		getNative().setBoldweight(EnumUtil.toFontBoldweight(boldweight));
	}
	public void setItalic(boolean italic) {
		getNative().setItalic(italic);		
	}
	
	public void setStrikeout(boolean strikeout) {
		getNative().setStrikeout(strikeout);	
	}
	public void setUnderline(Underline underline) {
		getNative().setUnderline(EnumUtil.toFontUnderline(underline));
	}
	public void setFontHeight(int height){
		getNative().setFontHeight((short)height);
	}
	public void setColor(Color color) {
		BookHelper.setFontColor(_bookRef.get(), getNative(), ((ColorImpl)color).getNative());
	}
	public void setFontHeightInPoint(int point) {
		setFontHeight(UnitUtil.pointToTwip(point));	
	}
}
