package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
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
	
	
	ModelRef<Book> bookRef;
	ModelRef<CellStyle> styleRef;
	
	NFont nfont;
	
	public NCellStyle(ModelRef<Book> book,ModelRef<CellStyle> style) {
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
		NCellStyle other = (NCellStyle) obj;
		if (styleRef == null) {
			if (other.styleRef != null)
				return false;
		} else if (!styleRef.equals(other.styleRef))
			return false;
		return true;
	}

	public NFont getFont() {
		if(nfont!=null){
			return nfont;
		}
		Book book = bookRef.get();
		Font font = book.getFontAt(getNative().getFontIndex());
		return nfont = new NFont(bookRef,new SimpleRef<Font>(font));
	}

	public void setFont(NFont nfont) {
		this.nfont = nfont; 
		getNative().setFont(nfont==null?null:nfont.getNative());
	}
	
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(getNative()).append("]");
		sb.append("font:[").append(getFont()).append("]");
		return sb.toString();
	}

	public void cloneAttribute(NCellStyle src) {
		getNative().cloneStyleFrom(src.getNative());
	}

	public NColor getBackgroundColor() {
		Color srcColor = getNative().getFillForegroundColorColor();
		return new NColor(bookRef,new SimpleRef<Color>(srcColor));
	}
	
	public void setBackgroundColor(NColor color) {
		BookHelper.setFillForegroundColor(getNative(), color.getNative());
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
		BookHelper.setTopBorderColor(getNative(), color);
	}
	public NColor getBorderTopColor(){
		return new NColor(bookRef,new SimpleRef<Color>(getNative().getTopBorderColorColor()));
	}
	
	public void setBorderLeftColor(Color color){
		BookHelper.setLeftBorderColor(getNative(), color);
	}
	public NColor getBorderLeftColor(){
		return new NColor(bookRef,new SimpleRef<Color>(getNative().getLeftBorderColorColor()));
	}
	
	public void setBorderBottomColor(Color color){
		BookHelper.setBottomBorderColor(getNative(), color);
	}
	public NColor getBorderBottomColor(){
		return new NColor(bookRef,new SimpleRef<Color>(getNative().getBottomBorderColorColor()));
	}
	
	public void setBorderRightColor(Color color){
		BookHelper.setRightBorderColor(getNative(), color);
	}
	public NColor getBorderRightColor(){
		return new NColor(bookRef,new SimpleRef<Color>(getNative().getRightBorderColorColor()));
	}
}
