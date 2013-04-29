package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NColor;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;

public class NCellStyleImpl implements NCellStyle{
	
	ModelRef<Book> bookRef;
	ModelRef<CellStyle> styleRef;
	
	NFontImpl nfont;
	
	public NCellStyleImpl(ModelRef<Book> book,ModelRef<CellStyle> style) {
		this.bookRef = book;
		this.styleRef = style;
	}
	
	public CellStyle getNative(){
		return styleRef.get();
	}
	public ModelRef<CellStyle> getRef(){
		return styleRef;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((styleRef == null) ? 0 : styleRef.hashCode());
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
		NCellStyleImpl other = (NCellStyleImpl) obj;
		if (styleRef == null) {
			if (other.styleRef != null)
				return false;
		} else if (!styleRef.equals(other.styleRef))
			return false;
		return true;
	}

	public NFontImpl getFont() {
		if(nfont!=null){
			return nfont;
		}
		Book book = bookRef.get();
		Font font = book.getFontAt(getNative().getFontIndex());
		return nfont = new NFontImpl(bookRef,new SimpleRef<Font>(font));
	}

	public void setFont(NFont nfont) {
		this.nfont = (NFontImpl)nfont; 
		getNative().setFont(nfont==null?null:this.nfont.getNative());
	}
	
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(getNative()).append("]");
		sb.append("font:[").append(getFont()).append("]");
		return sb.toString();
	}

	public void cloneAttribute(NCellStyle src) {
		getNative().cloneStyleFrom(((NCellStyleImpl)src).getNative());
	}

	public NColorImpl getBackgroundColor() {
		Color srcColor = getNative().getFillForegroundColorColor();
		return new NColorImpl(bookRef,new SimpleRef<Color>(srcColor));
	}
	
	public void setBackgroundColor(NColor color) {
		BookHelper.setFillForegroundColor(getNative(), ((NColorImpl)color).getNative());
	}
	
	public FillPattern getFillPattern(){
		return EnumUtil.toStyleFillPattern(getNative().getFillPattern());
	}

	public void setFillPattern(FillPattern pattern) {
		getNative().setFillPattern(EnumUtil.toStyleFillPattern(pattern));	
	}
	
//	public NColor getForegroundColor(){
//		return getFont().getColor();
//	}
//	
//	public void setForegroundColor(NColor color){
//		getFont().setFontColor(color);
//	}

//	public void setFontColor(NColor color) {
//		//set color form here will not go through BookHelper, cause set color issue of a theme color in XSSFont.
//		//use font set color
//		style.setFontColorColor(color.getNative());
//	}

	public void setAlignment(Alignment alignment){
		getNative().setAlignment(EnumUtil.toStyleAlignemnt(alignment));
	}
	public Alignment getAlignment(){
		return EnumUtil.toStyleAlignemnt(getNative().getAlignment());
	}
	public void setVerticalAlignment(VerticalAlignment alignment){
		getNative().setVerticalAlignment(EnumUtil.toStyleVerticalAlignemnt(alignment));
	}
	public VerticalAlignment getVerticalAlignment(){
		return EnumUtil.toStyleVerticalAlignemnt(getNative().getVerticalAlignment());
	}
	

	public boolean isWrapText() {
		return getNative().getWrapText();
	}

	public void setWrapText(boolean wraptext) {
		getNative().setWrapText(wraptext);
	}
	
	public void setBorderLeft(BorderType borderType){
		getNative().setBorderLeft(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderLeft(){
		return EnumUtil.toStyleBorderType(getNative().getBorderLeft());
	}

	public void setBorderTop(BorderType borderType){
		getNative().setBorderTop(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderTop(){
		return EnumUtil.toStyleBorderType(getNative().getBorderTop());
	}

	public void setBorderRight(BorderType borderType){
		getNative().setBorderRight(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderRight(){
		return EnumUtil.toStyleBorderType(getNative().getBorderRight());
	}

	public void setBorderBottom(BorderType borderType){
		getNative().setBorderBottom(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderBottom(){
		return EnumUtil.toStyleBorderType(getNative().getBorderBottom());
	}
	
	public void setBorderTopColor(NColor color){
		BookHelper.setTopBorderColor(getNative(),  ((NColorImpl)color).getNative());
	}
	public NColor getBorderTopColor(){
		return new NColorImpl(bookRef,new SimpleRef<Color>(getNative().getTopBorderColorColor()));
	}
	
	public void setBorderLeftColor(NColor color){
		BookHelper.setLeftBorderColor(getNative(), ((NColorImpl)color).getNative());
	}
	public NColor getBorderLeftColor(){
		return new NColorImpl(bookRef,new SimpleRef<Color>(getNative().getLeftBorderColorColor()));
	}
	
	public void setBorderBottomColor(NColor color){
		BookHelper.setBottomBorderColor(getNative(), ((NColorImpl)color).getNative());
	}
	public NColor getBorderBottomColor(){
		return new NColorImpl(bookRef,new SimpleRef<Color>(getNative().getBottomBorderColorColor()));
	}
	
	public void setBorderRightColor(NColor color){
		BookHelper.setRightBorderColor(getNative(), ((NColorImpl)color).getNative());
	}
	public NColor getBorderRightColor(){
		return new NColorImpl(bookRef,new SimpleRef<Color>(getNative().getRightBorderColorColor()));
	}
}
