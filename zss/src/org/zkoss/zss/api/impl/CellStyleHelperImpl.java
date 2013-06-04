/* CellStyleHelperImpl.java

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
package org.zkoss.zss.api.impl;

import org.zkoss.zss.api.Range.CellStyleHelper;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.EditableFont;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.CellStyleImpl;
import org.zkoss.zss.api.model.impl.ColorImpl;
import org.zkoss.zss.api.model.impl.EditableCellStyleImpl;
import org.zkoss.zss.api.model.impl.EditableFontImpl;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.api.model.impl.FontImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
/*package*/ class CellStyleHelperImpl implements CellStyleHelper{

	RangeImpl range;
	
	public CellStyleHelperImpl(RangeImpl range) {
		this.range = range;
	}
	
	/**
	 * create a new cell style and clone from src if it is not null
	 * @param src the source to clone, could be null
	 * @return the new cell style
	 */
	public EditableCellStyle createCellStyle(CellStyle src){
		XBook book = range.getNative().getSheet().getBook();
		EditableCellStyle style = new EditableCellStyleImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.CellStyle>(book.createCellStyle()));
		if(src!=null){
			((EditableCellStyleImpl)style).copyAttributeFrom(src);
		}
		return style;
	}

	public EditableFont createFont(Font src) {
		XBook book = range.getNative().getSheet().getBook();
		org.zkoss.poi.ss.usermodel.Font font = book.createFont();

		EditableFont nf = new EditableFontImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.Font>(font));
		if(src!=null){
			((EditableFontImpl)nf).copyAttributeFrom(src);
		}
			
		return nf;
	}

	public Color createColorFromHtmlColor(String htmlColor) {
		Book book = range.getBook();
		org.zkoss.poi.ss.usermodel.Color color = BookHelper.HTMLToColor(
				((BookImpl) book).getNative(), htmlColor);// never null
		return new ColorImpl(((BookImpl) book).getRef(),
				new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(color));
	}

	public Font findFont(Boldweight boldweight, Color color,
			int fontHeight, String fontName, boolean italic,
			boolean strikeout, TypeOffset typeOffset, Underline underline) {
		Book book = range.getBook();

		org.zkoss.poi.ss.usermodel.Font font;

		font = ((BookImpl) book).getNative().findFont(
				EnumUtil.toFontBoldweight(boldweight),
				((ColorImpl) color).getNative(), (short)fontHeight, fontName, italic,
				strikeout, EnumUtil.toFontTypeOffset(typeOffset),
				EnumUtil.toFontUnderline(underline));
		return font == null ? null : new FontImpl(((BookImpl) book).getRef(),
				new SimpleRef<org.zkoss.poi.ss.usermodel.Font>(font));
	}
}
