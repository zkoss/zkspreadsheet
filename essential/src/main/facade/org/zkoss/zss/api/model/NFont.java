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
	
	Book book;
	Font font;
	
	public NFont(Book book, Font font) {
		this.book = book;
		this.font = font;
	}
	public String getFontName() {
		return font.getFontName();
	}
	public Font getNative() {
		return font;
	}
	
	public NColor getColor(){
		Color c = BookHelper.getFontColor(book, font);
		return new NColor(book,c);
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((font == null) ? 0 : font.hashCode());
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
		if (font == null) {
			if (other.font != null)
				return false;
		} else if (!font.equals(other.font))
			return false;
		return true;
	}
	public Boldweight getBoldweight() {
		return EnumUtil.toFontBoldweight(font.getBoldweight());
	}
	public short getFontHeight() {
		return font.getFontHeight();
	}
	public boolean isItalic() {
		return font.getItalic();
	}
	public boolean isStrikeout() {
		return font.getStrikeout();
	}
	public TypeOffset getTypeOffset() {
		return EnumUtil.toFontTypeOffset(font.getTypeOffset());
	}
	public Underline getUnderline() {
		return EnumUtil.toFontUnderline(font.getUnderline());
	}

	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(font).append("]");
		sb.append(getFontName());
		return sb.toString();
	}
	public void cloneAttribute(NFont src) {
		Font sfont = src.getNative();
		
		font.setBoldweight(sfont.getBoldweight());
		Color srcColor = BookHelper.getFontColor(book, sfont);
		BookHelper.setFontColor(book, font, srcColor);
		font.setFontHeight(sfont.getFontHeight());
		font.setFontName(sfont.getFontName());
		font.setItalic(sfont.getItalic());
		font.setStrikeout(sfont.getStrikeout());
		font.setTypeOffset(sfont.getTypeOffset());
		font.setUnderline(sfont.getUnderline());
	}
	public void setFontName(String fontName) {
		font.setFontName(fontName);
	}
	public void setBoldweight(Boldweight boldweight) {
		font.setBoldweight(EnumUtil.toFontBoldweight(boldweight));
	}
	public void setItalic(boolean italic) {
		font.setItalic(italic);		
	}
	
	public void setStrikeout(boolean strikeout) {
		font.setStrikeout(strikeout);	
	}
	public void setUnderline(Underline underline) {
		font.setUnderline(EnumUtil.toFontUnderline(underline));
	}
	public void setFontHeight(short height){
		font.setFontHeight(height);
	}
	public void setColor(NColor color) {
		BookHelper.setFontColor(book, font, color.getNative());
	}
}
