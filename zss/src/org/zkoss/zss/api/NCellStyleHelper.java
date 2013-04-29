package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.impl.NBookImpl;
import org.zkoss.zss.api.model.impl.NCellStyleImpl;
import org.zkoss.zss.api.model.impl.NFontImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.Book;

//TODO rename to cell helper?
public interface NCellStyleHelper {

	/**
	 * create a new cell style and clone from src if it is not null
	 * @param src the source to clone, could be null
	 * @return the new cell style
	 */
	public NCellStyle createCellStyle(NCellStyle src);

	public NFont createFont(NFont src);
}
