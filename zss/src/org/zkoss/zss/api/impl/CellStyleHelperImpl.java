package org.zkoss.zss.api.impl;

import org.zkoss.zss.api.CellStyleHelper;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.CellStyleImpl;
import org.zkoss.zss.api.model.impl.FontImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;

public class CellStyleHelperImpl implements CellStyleHelper{

	RangeImpl range;
	
	public CellStyleHelperImpl(RangeImpl range) {
		this.range = range;
	}
	
	/**
	 * create a new cell style and clone from src if it is not null
	 * @param src the source to clone, could be null
	 * @return the new cell style
	 */
	public CellStyle createCellStyle(CellStyle src){
		XBook book = range.getNative().getSheet().getBook();
//		synchronized(book){//should be protected in range.batch,visit
			CellStyle style = new CellStyleImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.CellStyle>(book.createCellStyle()));
			if(src!=null){
				style.cloneAttribute(src);
			}
			return style;
//		}
	}

	public Font createFont(Font src) {
		XBook book = range.getNative().getSheet().getBook();
//		synchronized(book){
		org.zkoss.poi.ss.usermodel.Font font = book.createFont();

			Font nf = new FontImpl(((BookImpl)range.getBook()).getRef(),new SimpleRef<org.zkoss.poi.ss.usermodel.Font>(font));
			if(src!=null){
				nf.cloneAttribute(src);
			}
			
			return nf;
//		}
	}
	
//	public NCellAnchor createCellAnchor(int row,int col,int width,int height){
//		NCellAnchor anchor = null;
//		Worksheet sheet = range.getNative().getSheet();
//		sheet.getColumnWidth(columnIndex)
//		
//		
//		return anchor;
//	}
}
