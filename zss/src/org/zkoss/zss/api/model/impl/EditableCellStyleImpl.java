/* EditableCellStyleImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/6/4 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.DataFormat;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class EditableCellStyleImpl extends CellStyleImpl implements EditableCellStyle{
	
	public EditableCellStyleImpl(ModelRef<XBook> book,ModelRef<org.zkoss.poi.ss.usermodel.CellStyle> style) {
		super(book,style);
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
		EditableCellStyleImpl other = (EditableCellStyleImpl) obj;
		if (styleRef == null) {
			if (other.styleRef != null)
				return false;
		} else if (!styleRef.equals(other.styleRef))
			return false;
		return true;
	}
	public void setFont(Font nfont) {
		this.nfont = (FontImpl)nfont; 
		getNative().setFont(nfont==null?null:this.nfont.getNative());
	}

	public void copyAttributeFrom(CellStyle src) {
		getNative().cloneStyleFrom(((CellStyleImpl)src).getNative());
	}

	public void setBackgroundColor(Color color) {
		BookHelper.setFillForegroundColor(getNative(), ((ColorImpl)color).getNative());
	}

	public void setFillPattern(FillPattern pattern) {
		getNative().setFillPattern(EnumUtil.toStyleFillPattern(pattern));	
	}

	public void setAlignment(Alignment alignment){
		getNative().setAlignment(EnumUtil.toStyleAlignemnt(alignment));
	}

	public void setVerticalAlignment(VerticalAlignment alignment){
		getNative().setVerticalAlignment(EnumUtil.toStyleVerticalAlignemnt(alignment));
	}
	
	public void setWrapText(boolean wraptext) {
		getNative().setWrapText(wraptext);
	}

	public void setBorderLeft(BorderType borderType){
		getNative().setBorderLeft(EnumUtil.toStyleBorderType(borderType));
	}

	public void setBorderTop(BorderType borderType){
		getNative().setBorderTop(EnumUtil.toStyleBorderType(borderType));
	}

	public void setBorderRight(BorderType borderType){
		getNative().setBorderRight(EnumUtil.toStyleBorderType(borderType));
	}
	
	public void setBorderBottom(BorderType borderType){
		getNative().setBorderBottom(EnumUtil.toStyleBorderType(borderType));
	}
	
	public void setBorderTopColor(Color color){
		BookHelper.setTopBorderColor(getNative(),  ((ColorImpl)color).getNative());
	}
	
	public void setBorderLeftColor(Color color){
		BookHelper.setLeftBorderColor(getNative(), ((ColorImpl)color).getNative());
	}

	public void setBorderBottomColor(Color color){
		BookHelper.setBottomBorderColor(getNative(), ((ColorImpl)color).getNative());
	}
	
	public void setBorderRightColor(Color color){
		BookHelper.setRightBorderColor(getNative(), ((ColorImpl)color).getNative());
	}
	
	public void setDataFormat(String format){
		if(getDataFormat().equals(format)){
			return;
		}
		//this api doesn't create a new df, it just return the one in book
		DataFormat df = bookRef.get().createDataFormat();

		short index = df.getFormat(format);		
		getNative().setDataFormat(index);
	}

	public void setLocked(boolean locked) {
		getNative().setLocked(locked);
	}

	public void setHidden(boolean hidden) {
		getNative().setHidden(hidden);;
	}
}
