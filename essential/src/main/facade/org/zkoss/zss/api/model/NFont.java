package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;

public class NFont {
	
	public enum TypeOffset{
		NONE, 
		SUPER, 
		SUB
	}
	public enum Underline{
		NONE,
		SINGLE,
		DOUBLE,
		SINGLE_ACCOUNTING,
		DOUBLE_ACCOUNTING
	}
	
	public enum Boldweight{
		NORMAL,
		BOLD
	}
	
	ModelRef<Book> bookRef;
	ModelRef<Font> fontRef;
	
	public NFont(ModelRef<Book> book, ModelRef<Font> font) {
		this.bookRef = book;
		this.fontRef = font;
	}
	public String getFontName() {
		return getNative().getFontName();
	}
	public Font getNative() {
		return fontRef.get();
	}
	public ModelRef<Font> getRef(){
		return fontRef;
	}
	
	public NColor getColor(){
		Color c = BookHelper.getFontColor(bookRef.get(), fontRef.get());
		return new NColor(bookRef,new SimpleRef<Color>(c));
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
		NFont other = (NFont) obj;
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
	public short getFontHeight() {
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
	public void cloneAttribute(NFont src) {
		Font sfont = src.getNative();
		Font font = getNative();
		font.setBoldweight(sfont.getBoldweight());
		Color srcColor = BookHelper.getFontColor(bookRef.get(), sfont);
		BookHelper.setFontColor(bookRef.get(), font, srcColor);
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
	public void setFontHeight(short height){
		getNative().setFontHeight(height);
	}
	public void setColor(NColor color) {
		BookHelper.setFontColor(bookRef.get(), getNative(), color.getNative());
	}
}
