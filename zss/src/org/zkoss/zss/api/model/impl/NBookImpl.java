package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NColor;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.Book;
import org.zkoss.zss.model.sys.Worksheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;

public class NBookImpl implements NBook{

	ModelRef<Book> bookRef;
	BookType type;
	
	public NBookImpl(ModelRef<Book> ref){
		this.bookRef = ref;
		Book book = ref.get();
		if (book instanceof HSSFBookImpl) {
			type = BookType.EXCEL_2003;
		} else if (book instanceof XSSFBookImpl) {
			type = BookType.EXCEL_2007;
		} else {
			throw new IllegalArgumentException("unknow book type "+book);
		}
	}

	public Book getNative() {
		return bookRef.get();
	}
	
	public ModelRef<Book> getRef(){
		return bookRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((bookRef == null) ? 0 : bookRef.hashCode());
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
		NBookImpl other = (NBookImpl) obj;
		if (bookRef == null) {
			if (other.bookRef != null)
				return false;
		} else if (!bookRef.equals(other.bookRef))
			return false;
		return true;
	}

	public String getBookName() {
		return getNative().getBookName();
	}
	
	public BookType getType(){
		return type; 
	}

	public int getSheetIndex(NSheet sheet) {
		if(sheet==null) return -1;
		return getNative().getSheetIndex(((NSheetImpl)sheet).getNative());
	}
	

	public NFontImpl findFont(Boldweight boldweight, NColor color, short fontHeight,
			String fontName, boolean italic, boolean strikeout,
			TypeOffset typeOffset, Underline underline) {
		Font font;
		
		font = getNative().findFont(EnumUtil.toFontBoldweight(boldweight), ((NColorImpl)color).getNative(), fontHeight, fontName,
				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
		return font==null?null:new NFontImpl(bookRef,new SimpleRef<Font>(font));
	}

	public NColorImpl getColorFromHtmlColor(String htmlColor) {
		Color color = BookHelper.HTMLToColor(getNative(), htmlColor);//never null
		return new NColorImpl(bookRef,new SimpleRef<Color>(color));
	}

	public int getNumberOfSheets() {
		return getNative().getNumberOfSheets();
	}
	
	public NSheetImpl getSheetAt(int index){
		Worksheet sheet = getNative().getWorksheetAt(index);
		return new NSheetImpl(new SimpleRef<Worksheet>(sheet));
	}
	
	public NSheetImpl getSheet(String name){
		Worksheet sheet = getNative().getWorksheet(name);
		
		return sheet==null?null:new NSheetImpl(new SimpleRef<Worksheet>(sheet));
	}
	
}
