package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;

public class BookImpl implements Book{

	ModelRef<XBook> bookRef;
	BookType type;
	
	public BookImpl(ModelRef<XBook> ref){
		this.bookRef = ref;
		XBook book = ref.get();
		if (book instanceof HSSFBookImpl) {
			type = BookType.EXCEL_2003;
		} else if (book instanceof XSSFBookImpl) {
			type = BookType.EXCEL_2007;
		} else {
			throw new IllegalArgumentException("unknow book type "+book);
		}
	}

	public XBook getNative() {
		return bookRef.get();
	}
	
	public ModelRef<XBook> getRef(){
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
		BookImpl other = (BookImpl) obj;
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

	public int getSheetIndex(Sheet sheet) {
		if(sheet==null) return -1;
		return getNative().getSheetIndex(((SheetImpl)sheet).getNative());
	}
	

	public FontImpl findFont(Boldweight boldweight, Color color, short fontHeight,
			String fontName, boolean italic, boolean strikeout,
			TypeOffset typeOffset, Underline underline) {
		Font font;
		
		font = getNative().findFont(EnumUtil.toFontBoldweight(boldweight), ((ColorImpl)color).getNative(), fontHeight, fontName,
				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
		return font==null?null:new FontImpl(bookRef,new SimpleRef<Font>(font));
	}

	public Color getColorFromHtmlColor(String htmlColor) {
		org.zkoss.poi.ss.usermodel.Color color = BookHelper.HTMLToColor(getNative(), htmlColor);//never null
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(color));
	}

	public int getNumberOfSheets() {
		return getNative().getNumberOfSheets();
	}
	
	public SheetImpl getSheetAt(int index){
		XSheet sheet = getNative().getWorksheetAt(index);
		return new SheetImpl(new SimpleRef<XSheet>(sheet));
	}
	
	public SheetImpl getSheet(String name){
		XSheet sheet = getNative().getWorksheet(name);
		
		return sheet==null?null:new SheetImpl(new SimpleRef<XSheet>(sheet));
	}
	
}
