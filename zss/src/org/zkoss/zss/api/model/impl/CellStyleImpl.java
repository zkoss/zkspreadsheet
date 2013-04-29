package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;

public class CellStyleImpl implements CellStyle{
	
	ModelRef<XBook> bookRef;
	ModelRef<org.zkoss.poi.ss.usermodel.CellStyle> styleRef;
	
	FontImpl nfont;
	
	public CellStyleImpl(ModelRef<XBook> book,ModelRef<org.zkoss.poi.ss.usermodel.CellStyle> style) {
		this.bookRef = book;
		this.styleRef = style;
	}
	
	public org.zkoss.poi.ss.usermodel.CellStyle getNative(){
		return styleRef.get();
	}
	public ModelRef<org.zkoss.poi.ss.usermodel.CellStyle> getRef(){
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
		CellStyleImpl other = (CellStyleImpl) obj;
		if (styleRef == null) {
			if (other.styleRef != null)
				return false;
		} else if (!styleRef.equals(other.styleRef))
			return false;
		return true;
	}

	public FontImpl getFont() {
		if(nfont!=null){
			return nfont;
		}
		XBook book = bookRef.get();
		org.zkoss.poi.ss.usermodel.Font font = book.getFontAt(getNative().getFontIndex());
		return nfont = new FontImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Font>(font));
	}

	public void setFont(Font nfont) {
		this.nfont = (FontImpl)nfont; 
		getNative().setFont(nfont==null?null:this.nfont.getNative());
	}
	
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(getNative()).append("]");
		sb.append("font:[").append(getFont()).append("]");
		return sb.toString();
	}

	public void cloneAttribute(CellStyle src) {
		getNative().cloneStyleFrom(((CellStyleImpl)src).getNative());
	}

	public ColorImpl getBackgroundColor() {
		org.zkoss.poi.ss.usermodel.Color srcColor = getNative().getFillForegroundColorColor();
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(srcColor));
	}
	
	public void setBackgroundColor(Color color) {
		BookHelper.setFillForegroundColor(getNative(), ((ColorImpl)color).getNative());
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
	
	public void setBorderTopColor(Color color){
		BookHelper.setTopBorderColor(getNative(),  ((ColorImpl)color).getNative());
	}
	public Color getBorderTopColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getTopBorderColorColor()));
	}
	
	public void setBorderLeftColor(Color color){
		BookHelper.setLeftBorderColor(getNative(), ((ColorImpl)color).getNative());
	}
	public Color getBorderLeftColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getLeftBorderColorColor()));
	}
	
	public void setBorderBottomColor(Color color){
		BookHelper.setBottomBorderColor(getNative(), ((ColorImpl)color).getNative());
	}
	public Color getBorderBottomColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getBottomBorderColorColor()));
	}
	
	public void setBorderRightColor(Color color){
		BookHelper.setRightBorderColor(getNative(), ((ColorImpl)color).getNative());
	}
	public Color getBorderRightColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getRightBorderColorColor()));
	}
}
