package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.HSSFBookImpl;
import org.zkoss.zss.model.impl.XSSFBookImpl;

public class NBook {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}

	Book book;
	BookType type;
	
	public NBook(Book book){
		this.book = book;
		if (book instanceof HSSFBookImpl) {
			type = BookType.EXCEL_2003;
		} else if (book instanceof XSSFBookImpl) {
			type = BookType.EXCEL_2007;
		} else {
			throw new IllegalArgumentException("unknow book type "+book);
		}
	}

	public Book getNative() {
		return book;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((book == null) ? 0 : book.hashCode());
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
		NBook other = (NBook) obj;
		if (book == null) {
			if (other.book != null)
				return false;
		} else if (!book.equals(other.book))
			return false;
		return true;
	}

	public String getBookName() {
		return book.getBookName();
	}
	
	public BookType getType(){
		return type; 
	}

	public int getSheetIndex(NSheet sheet) {
		if(sheet==null) return -1;
		return book.getSheetIndex(sheet.getNative());
	}
	

	public NFont findFont(Boldweight boldweight, NColor color, short fontHeight,
			String fontName, boolean italic, boolean strikeout,
			TypeOffset typeOffset, Underline underline) {
		Font font;
		
		font = book.findFont(EnumUtil.toFontBoldweight(boldweight), color.getNative(), fontHeight, fontName,
				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
		return font==null?null:new NFont(book,font);
	}

	public NColor getColorFromHtmlColor(String htmlColor) {
		Color color = BookHelper.HTMLToColor(book, htmlColor);//never null
		return new NColor(book,color);
	}
	
}
