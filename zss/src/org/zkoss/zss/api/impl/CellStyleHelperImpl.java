package org.zkoss.zss.api.impl;

import org.zkoss.zss.api.Range.CellStyleHelper;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.CellStyleImpl;
import org.zkoss.zss.api.model.impl.ColorImpl;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.api.model.impl.FontImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;

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
	public CellStyle createCellStyle(CellStyle src){
		XBook book = range.getNative().getSheet().getBook();
			CellStyle style = new CellStyleImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.CellStyle>(book.createCellStyle()));
			if(src!=null){
				((CellStyleImpl)style).copyAttributeFrom(src);
			}
			return style;
	}

	public Font createFont(Font src) {
		XBook book = range.getNative().getSheet().getBook();
		org.zkoss.poi.ss.usermodel.Font font = book.createFont();

			Font nf = new FontImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.Font>(font));
			if(src!=null){
				((FontImpl)nf).copyAttributeFrom(src);
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
