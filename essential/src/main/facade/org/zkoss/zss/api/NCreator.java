package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;

public class NCreator {

	NRange range;
	
	public NCreator(NRange range) {
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
			NCellStyle style = new NCellStyle(book,book.createCellStyle());
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

			NFont nf = new NFont(book,font);
			if(src!=null){
				nf.cloneAttribute(src);
			}
			
			return nf;
//		}
			

	}
}
