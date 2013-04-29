package org.zkoss.zss.api;

import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Font;

//TODO rename to cell helper?
public interface CellStyleHelper {

	/**
	 * create a new cell style and clone from src if it is not null
	 * @param src the source to clone, could be null
	 * @return the new cell style
	 */
	public CellStyle createCellStyle(CellStyle src);

	public Font createFont(Font src);
}
