/* CellStyleImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.model.sys.XBook;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
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
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(getNative()).append("]");
		sb.append("font:[").append(getFont()).append("]");
		return sb.toString();
	}

	public ColorImpl getBackgroundColor() {
		org.zkoss.poi.ss.usermodel.Color srcColor = getNative().getFillForegroundColorColor();
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(srcColor));
	}

	public FillPattern getFillPattern(){
		return EnumUtil.toStyleFillPattern(getNative().getFillPattern());
	}

	public Alignment getAlignment(){
		return EnumUtil.toStyleAlignemnt(getNative().getAlignment());
	}

	public VerticalAlignment getVerticalAlignment(){
		return EnumUtil.toStyleVerticalAlignemnt(getNative().getVerticalAlignment());
	}

	public boolean isWrapText() {
		return getNative().getWrapText();
	}
	
	public BorderType getBorderLeft(){
		return EnumUtil.toStyleBorderType(getNative().getBorderLeft());
	}

	public BorderType getBorderTop(){
		return EnumUtil.toStyleBorderType(getNative().getBorderTop());
	}

	public BorderType getBorderRight(){
		return EnumUtil.toStyleBorderType(getNative().getBorderRight());
	}

	public BorderType getBorderBottom(){
		return EnumUtil.toStyleBorderType(getNative().getBorderBottom());
	}

	public Color getBorderTopColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getTopBorderColorColor()));
	}

	public Color getBorderLeftColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getLeftBorderColorColor()));
	}

	public Color getBorderBottomColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getBottomBorderColorColor()));
	}

	public Color getBorderRightColor(){
		return new ColorImpl(bookRef,new SimpleRef<org.zkoss.poi.ss.usermodel.Color>(getNative().getRightBorderColorColor()));
	}
	
	public String getDataFormat(){
		return getNative().getDataFormatString();
	}

	public boolean isLocked() {
		return getNative().getLocked();
	}

	public boolean isHidden() {
		return getNative().getHidden();
	}
}
