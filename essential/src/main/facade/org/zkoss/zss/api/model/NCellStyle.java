package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;

public class NCellStyle {

	public enum FillPattern{
		NO_FILL,
		SOLID_FOREGROUND,
		FINE_DOTS,
		ALT_BARS,
		SPARSE_DOTS,
		THICK_HORZ_BANDS,
		THICK_VERT_BANDS,
		THICK_BACKWARD_DIAG,
		THICK_FORWARD_DIAG,
		BIG_SPOTS,
		BRICKS,
		THIN_HORZ_BANDS,
		THIN_VERT_BANDS,
		THIN_BACKWARD_DIAG,
		THIN_FORWARD_DIAG,
		SQUARES,
		DIAMONDS,
		LESS_DOTS,
		LEAST_DOTS
	}
	
	public enum Alignment{
		GENERAL,
		LEFT,
		CENTER,
		RIGHT,
		FILL,
		JUSTIFY,
		CENTER_SELECTION
	}
	public enum VerticalAlignment{
		TOP,
		CENTER,
		BOTTOM,
		JUSTIFY
	}
	
	public enum BorderType{
	    NONE,
	    THIN,
	    MEDIUM,
	    DASHED,
	    HAIR,
	    THICK,
	    DOUBLE,
	    DOTTED,
	    MEDIUM_DASHED,
	    DASH_DOT,
	    MEDIUM_DASH_DOT,
	    DASH_DOT_DOT,
	    MEDIUM_DASH_DOT_DOT,
	    SLANTED_DASH_DOT;
	}
	
	
	Book book;
	CellStyle style;
	
	NFont nfont;
	
	public NCellStyle(Book book,CellStyle style) {
		this.book = book;
		this.style = style;
	}
	
	public CellStyle getNative(){
		return style;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((style == null) ? 0 : style.hashCode());
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
		NCellStyle other = (NCellStyle) obj;
		if (style == null) {
			if (other.style != null)
				return false;
		} else if (!style.equals(other.style))
			return false;
		return true;
	}

	public NFont getFont() {
		if(nfont!=null){
			return nfont;
		}
		return nfont = new NFont(book,book.getFontAt(style.getFontIndex()));
	}

	public void setFont(NFont nfont) {
		this.nfont = nfont; 
		style.setFont(nfont==null?null:nfont.getNative());
	}
	
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(style).append("]");
		sb.append("font:[").append(getFont()).append("]");
		return sb.toString();
	}

	public void cloneAttribute(NCellStyle src) {
		style.cloneStyleFrom(src.getNative());
	}

	public NColor getBackgroundColor() {
		Color srcColor = style.getFillForegroundColorColor();
		return new NColor(book,srcColor);
	}
	
	public void setBackgroundColor(NColor color) {
		BookHelper.setFillForegroundColor(style, color.getNative());
	}
	
	public FillPattern getFillPattern(){
		return EnumUtil.toStyleFillPattern(style.getFillPattern());
	}

	public void setFillPattern(FillPattern pattern) {
		style.setFillPattern(EnumUtil.toStyleFillPattern(pattern));	
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
		style.setAlignment(EnumUtil.toStyleAlignemnt(alignment));
	}
	public Alignment getAlignment(){
		return EnumUtil.toStyleAlignemnt(style.getAlignment());
	}
	public void setVerticalAlignment(VerticalAlignment alignment){
		style.setVerticalAlignment(EnumUtil.toStyleVerticalAlignemnt(alignment));
	}
	public VerticalAlignment getVerticalAlignment(){
		return EnumUtil.toStyleVerticalAlignemnt(style.getVerticalAlignment());
	}
	

	public boolean isWrapText() {
		return style.getWrapText();
	}

	public void setWrapText(boolean wraptext) {
		style.setWrapText(wraptext);
	}
	
	public void setBorderLeft(BorderType borderType){
		style.setBorderLeft(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderLeft(){
		return EnumUtil.toStyleBorderType(style.getBorderLeft());
	}

	public void setBorderTop(BorderType borderType){
		style.setBorderTop(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderTop(){
		return EnumUtil.toStyleBorderType(style.getBorderTop());
	}

	public void setBorderRight(BorderType borderType){
		style.setBorderRight(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderRight(){
		return EnumUtil.toStyleBorderType(style.getBorderRight());
	}

	public void setBorderBottom(BorderType borderType){
		style.setBorderBottom(EnumUtil.toStyleBorderType(borderType));
	}
	public BorderType getBorderBottom(){
		return EnumUtil.toStyleBorderType(style.getBorderBottom());
	}
	
	public void setBorderTopColor(Color color){
		BookHelper.setTopBorderColor(style, color);
	}
	public NColor getBorderTopColor(){
		return new NColor(book,style.getTopBorderColorColor());
	}
	
	public void setBorderLeftColor(Color color){
		BookHelper.setLeftBorderColor(style, color);
	}
	public NColor getBorderLeftColor(){
		return new NColor(book,style.getLeftBorderColorColor());
	}
	
	public void setBorderBottomColor(Color color){
		BookHelper.setBottomBorderColor(style, color);
	}
	public NColor getBorderBottomColor(){
		return new NColor(book,style.getBottomBorderColorColor());
	}
	
	public void setBorderRightColor(Color color){
		BookHelper.setRightBorderColor(style, color);
	}
	public NColor getBorderRightColor(){
		return new NColor(book,style.getRightBorderColorColor());
	}
}
