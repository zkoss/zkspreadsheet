package org.zkoss.zss.api.impl;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.NCellStyleHelper;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.impl.NBookImpl;
import org.zkoss.zss.api.model.impl.NCellStyleImpl;
import org.zkoss.zss.api.model.impl.NFontImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.Book;

public class NCellStyleHelperImpl implements NCellStyleHelper{

	NRangeImpl range;
	
	public NCellStyleHelperImpl(NRangeImpl range) {
		this.range = range;
	}
	
	/**
	 * create a new cell style and clone from src if it is not null
	 * @param src the source to clone, could be null
	 * @return the new cell style
	 */
	public NCellStyle createCellStyle(NCellStyle src){
		Book book = range.getNative().getSheet().getBook();
//		synchronized(book){//should be protected in range.batch,visit
			NCellStyle style = new NCellStyleImpl(((NBookImpl)range.getBook()).getRef(),new SimpleRef<CellStyle>(book.createCellStyle()));
			if(src!=null){
				style.cloneAttribute(src);
			}
			return style;
//		}
	}

	public NFont createFont(NFont src) {
		Book book = range.getNative().getSheet().getBook();
//		synchronized(book){
			Font font = book.createFont();

			NFont nf = new NFontImpl(((NBookImpl)range.getBook()).getRef(),new SimpleRef<Font>(font));
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
