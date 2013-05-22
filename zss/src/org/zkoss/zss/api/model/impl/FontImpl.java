package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;

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
	public void copyAttributeFrom(Font src) {
		org.zkoss.poi.ss.usermodel.Font sfont = ((FontImpl)src).getNative();
		org.zkoss.poi.ss.usermodel.Font font = getNative();
		font.setBoldweight(sfont.getBoldweight());
		org.zkoss.poi.ss.usermodel.Color srcColor = BookHelper.getFontColor(bookRef.get(), sfont);
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
	public void setFontHeight(int height){
		getNative().setFontHeight((short)height);
	}
	public void setColor(Color color) {
		BookHelper.setFontColor(bookRef.get(), getNative(), ((ColorImpl)color).getNative());
	}
	@Override
	public int getFontHeightInPoint() {
		return UnitUtil.twipToPoint(getFontHeight());
	}
	@Override
	public void setFontHeightInPoint(int point) {
		setFontHeight(UnitUtil.pointToTwip(point));	
	}
}
