package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zss.model.Book;

public class NCellStyle {

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

//	public void setFontColor(NColor color) {
//		//set color form here will not go through BookHelper, cause set color issue of a theme color in XSSFont.
//		//use font set color
//		style.setFontColorColor(color.getNative());
//	}
	

}
